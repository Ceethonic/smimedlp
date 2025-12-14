async function sendToClasifier(url = '', data = {}, event) {
  printLog("Sending event to classifier");

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 35000);

  // jeśli agent trzyma request (confirm) — pokaż, że czekamy
  let waitInterval = null;
  const waitStart = setTimeout(() => {
    waitInterval = setInterval(() => {
      printLog("Waiting for DLP decision (confirm popup may be active)...");
    }, 1000);
  }, 1500);

  fetch(url, {
    signal: controller.signal,
    method: 'POST',
    mode: 'cors',
    cache: 'no-cache',
    credentials: 'same-origin',
    headers: { 'Content-Type': 'application/json' },
    redirect: 'follow',
    referrerPolicy: 'no-referrer',
    body: JSON.stringify(data)
  }).then(async (response) => {
    printLog("Classifier HTTP status: " + response.status);

    // czytelny RAW (to pozwoli zobaczyć 0/1/2/confirm itd.)
    const raw = await response.text().catch(() => "");
    printLog("Engine raw response: " + raw);

    if (!response.ok) {
      printLog("Engine returned error: " + response.status);
      handleError(response.status, event);
      return null;
    }

    try { return JSON.parse(raw); } catch (e) { return null; }
  }).then((responseJson) => {
    clearTimeout(timeout);
    clearTimeout(waitStart);
    if (waitInterval) clearInterval(waitInterval);

    if (!responseJson) {
      printLog("Engine response is not JSON");
      handleError("invalid_json", event);
      return;
    }

    handleResponse(responseJson, event);
  }).catch(e => {
    clearTimeout(timeout);
    clearTimeout(waitStart);
    if (waitInterval) clearInterval(waitInterval);

    printLog("Request crashed");
    printLog(e && e.name ? e.name : "unknown_error");
    handleError(e, event);
  });
}
