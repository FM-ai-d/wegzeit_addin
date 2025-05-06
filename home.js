// HINWEIS: Du musst deinen eigenen OpenRouteService API-Key einsetzen.
function berechneWegzeit() {
  const start = document.getElementById("start").value;
  const ziel = document.getElementById("ziel").value;
  const pufferMin = parseInt(document.getElementById("puffer").value || "0");
  const apiKey = "5b3ce3597851110001cf6248452ca34c9da540f3af9efe62e26ea48e";

  Promise.all([
    fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(start)}`).then(res => res.json()),
    fetch(`https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(ziel)}`).then(res => res.json())
  ]).then(([startData, zielData]) => {
    if (!startData.length || !zielData.length) {
      alert("Adressen konnten nicht gefunden werden.");
      return;
    }

    const startCoords = [parseFloat(startData[0].lon), parseFloat(startData[0].lat)];
    const zielCoords = [parseFloat(zielData[0].lon), parseFloat(zielData[0].lat)];

    fetch("https://api.openrouteservice.org/v2/directions/driving-car", {
      method: "POST",
      headers: {
        "Authorization": apiKey,
        "Content-Type": "application/json"
      },
      body: JSON.stringify({ coordinates: [startCoords, zielCoords] })
    })
    .then(res => res.json())
    .then(data => {
      const durationSec = data.features[0].properties.summary.duration;
      const totalSec = durationSec + pufferMin * 60;
      const totalMin = Math.ceil(totalSec / 60);

      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (tokenResult) {
        if (tokenResult.status !== "succeeded") return;

        Office.context.mailbox.item.start.getAsync((startResult) => {
          const meetingStart = new Date(startResult.value);
          const anfahrtStart = new Date(meetingStart.getTime() - totalSec * 1000);

          const newEvent = {
            Subject: `Anfahrt: ${start} → ${ziel}`,
            Start: anfahrtStart.toISOString(),
            End: meetingStart.toISOString(),
            ShowAs: "OOF",
            Body: { ContentType: "Text", Content: `Wegzeit inkl. Puffer: ca. ${totalMin} Min.` },
            Categories: ["Wegzeit"]
          };

          const url = `${Office.context.mailbox.restUrl}/v2.0/me/events`;
          fetch(url, {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${tokenResult.value}`,
              "Accept": "application/json",
              "Content-Type": "application/json"
            },
            body: JSON.stringify(newEvent)
          })
          .then(resp => resp.json())
          .then(() => {
            const resultDiv = document.getElementById("wegzeitErgebnis");
            resultDiv.innerText = `Anfahrtstermin erstellt: ${totalMin} Min. (inkl. Puffer).`;
            resultDiv.style.display = "block";
          });
        });
      });
    });
  });
}
