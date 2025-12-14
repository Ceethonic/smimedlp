/* global Office */
(function () {
  "use strict";

  var KEY_CFG  = "DLP_DEV_CFG";
  var KEY_LOGS = "DLP_DEV_LOGS";

  function $(id) { return document.getElementById(id); }
  function safeJsonParse(s) { try { return JSON.parse(s); } catch (e) { return null; } }

  function loadCfg() {
    var cfg = safeJsonParse(localStorage.getItem(KEY_CFG) || "{}") || {};
    $("dbgLevel").value   = (cfg.debugLevel != null) ? cfg.debugLevel : 3;
    $("failClosed").value = (cfg.failClosed === true) ? "true" : "false";
    $("agentPort").value  = (cfg.agentPort != null) ? cfg.agentPort : 55299;
    $("logSinkUrl").value = cfg.logSinkUrl || "";
    return cfg;
  }

  function saveCfg() {
    var cfg = safeJsonParse(localStorage.getItem(KEY_CFG) || "{}") || {};
    cfg.debugLevel = parseInt($("dbgLevel").value, 10);
    if (isNaN(cfg.debugLevel)) cfg.debugLevel = 3;
    cfg.failClosed = ($("failClosed").value === "true");
    cfg.agentPort  = parseInt($("agentPort").value, 10);
    if (isNaN(cfg.agentPort)) cfg.agentPort = 55299;
    cfg.logSinkUrl = $("logSinkUrl").value || "";

    localStorage.setItem(KEY_CFG, JSON.stringify(cfg));

    // jeśli akurat jesteśmy w tym samym kontekście co app.js (nie zawsze), spróbujmy przekazać
    try {
      if (window.opener && window.opener.DLP_DEV_CFG_SAVE) window.opener.DLP_DEV_CFG_SAVE(cfg);
    } catch (e) {}

    return cfg;
  }

  function renderLogs() {
    var arr = safeJsonParse(localStorage.getItem(KEY_LOGS) || "[]") || [];
    var lines = [];
    for (var i = 0; i < arr.length; i++) {
      var e = arr[i];
      lines.push((e.ts || "") + " [" + (e.lvl || "") + "] [" + (e.tx || "") + "] " + (e.msg || "") +
        (e.data ? (" " + JSON.stringify(e.data)) : ""));
    }
    $("logs").value = lines.join("\n");
  }

  function copyLogs() {
    var text = $("logs").value || "";
    // IE/legacy fallback
    try {
      $("logs").focus();
      $("logs").select();
      document.execCommand("copy");
    } catch (e) {}
  }

  function downloadLogs() {
    // Uwaga: Outlook Desktop bywa restrykcyjny z downloadami. :contentReference[oaicite:3]{index=3}
    var text = $("logs").value || "";
    try {
      var blob = new Blob([text], { type: "text/plain;charset=utf-8" });
      var url = URL.createObjectURL(blob);
      var a = document.createElement("a");
      a.href = url;
      a.download = "dlp-dev-logs.txt";
      document.body.appendChild(a);
      a.click();
      setTimeout(function () {
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }, 100);
    } catch (e) {}
  }

  function clearLogs() {
    try { localStorage.removeItem(KEY_LOGS); } catch (e) {}
    renderLogs();
  }

  function wire() {
    $("btnSave").onclick = function () { saveCfg(); };
    $("btnRefresh").onclick = function () { renderLogs(); };
    $("btnCopy").onclick = function () { copyLogs(); };
    $("btnDownload").onclick = function () { downloadLogs(); };
    $("btnClear").onclick = function () { clearLogs(); };

    loadCfg();
    renderLogs();
  }

  try {
    Office.onReady(function () { wire(); });
  } catch (e) {
    wire();
  }
})();
