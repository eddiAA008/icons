Office.onReady(() => {
  // Office.js is ready
});

function reportPhishing() {
  const item = Office.context.mailbox.item;

  item.getAllInternetHeadersAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const headers = result.value;

      fetch("https://your-api-endpoint.example.com/report", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ headers })
      })
      .then(() => alert("Phishing-Meldung wurde gesendet."))
      .catch((err) => alert("Fehler beim Senden: " + err));
    } else {
      alert("Fehler beim Abrufen der Mail-Header.");
    }
  });
}
