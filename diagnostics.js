(() => {
  "use strict";

  const LOG_KEY = "smimeDlp.logs.v1";
  const DEBUG_KEY = "smimeDlp.debug.v1";

  function loadLogs() {
    try {
      const raw = localStorage.getItem(LOG_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  function saveLogs(arr) {
    try { localStorage.setItem(LOG_KEY, JSON.stringify(arr)); } catch { /* ignore */ }
  }

  function isDebug() {
    try { return localStorage.getItem(DEBUG_KEY) === "1"; } catch { return false; }
  }

  function setDebug(v) {
    try { localStorage.setItem(DEBUG_KEY, v ? "1" : "0"); } catch { /* ignore */ }
  }

  function formatEntry(e) {
    const meta = e && e.meta ? " " + JSON.stringify(e.meta) : "";
    return `${e.ts || ""} [${e.level || ""}] ${e.msg || ""}${meta}`;
  }

  function render() {
    const logEl = document.getElementById("log");
    const dbgPill = document.getElementById("dbgPill");

    const logs = loadLogs();
    dbgPill.textContent = `debug: ${isDebug() ? "ON" : "OFF"}`;

    logEl.textContent = logs.map(formatEntry).join("\n");
    // keep scrolled near bottom
    logEl.scrollTop = logEl.scrollHeight;
  }

  function download() {
    const logs = loadLogs();
    const content = logs.map(formatEntry).join("\n") + "\n";
    const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "smimeDlp.logs.txt";
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  function wireUi() {
    document.getElementById("btnRefresh").addEventListener("click", render);

    document.getElementById("btnToggleDebug").addEventListener("click", () => {
      setDebug(!isDebug());
      render();
    });

    document.getElementById("btnClear").addEventListener("click", () => {
      saveLogs([]);
      render();
    });

    document.getElementById("btnDownload").addEventListener("click", download);

    // Live updates (best effort)
    try {
      const bc = new BroadcastChannel("smimeDlpLogs");
      bc.onmessage = () => render();
    } catch { /* ignore */ }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => {
      wireUi();
      render();
    });
  } else {
    wireUi();
    render();
  }
})();