/* global Office */

(() => {
  "use strict";

  const LOG_KEY = "DLP_DEV_LOGS_V1";
  const UI_LOG_KEY = "DLP_DEV_LOG_ENABLE";
  const NO_CT_KEY  = "DLP_DEV_NO_CONTENT_TYPE";

  function getBool(key) {
    try { return localStorage.getItem(key) === "1"; } catch { return false; }
  }
  function setBool(key, v) {
    try { localStorage.setItem(key, v ? "1" : "0"); } catch {}
  }

  function readLogs() {
    try {
      const raw = localStorage.getItem(LOG_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  function writeLogs(lines) {
    try { localStorage.setItem(LOG_KEY, JSON.stringify(lines)); } catch {}
  }

  function render() {
    const out = document.getElementById("out");
    const lines = readLogs();
    out.textContent = lines.length ? lines.join("\n") : "(brak logów)";
    out.scrollTop = out.scrollHeight;
  }

  function download() {
    const lines = readLogs();
    const blob = new Blob([lines.join("\n")], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "dlp_addin_logs.txt";
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  Office.onReady().then(() => {
    const uiLog = document.getElementById("uiLog");
    const noCT = document.getElementById("noCT");

    uiLog.checked = getBool(UI_LOG_KEY);
    noCT.checked  = getBool(NO_CT_KEY);

    uiLog.addEventListener("change", () => setBool(UI_LOG_KEY, uiLog.checked));
    noCT.addEventListener("change", () => setBool(NO_CT_KEY, noCT.checked));

    document.getElementById("refresh").addEventListener("click", render);
    document.getElementById("clear").addEventListener("click", () => { writeLogs([]); render(); });
    document.getElementById("download").addEventListener("click", download);

    // live stream if available
    try {
      const bc = new BroadcastChannel("DLP_DEV_LOGS_CH");
      bc.onmessage = () => render();
    } catch (e) {}

    render();
    setInterval(render, 1000);
  });
})();
