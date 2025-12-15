/* global Office, OfficeRuntime */

"use strict";

const LOG_KEY = "DLP_DIAG_LOGS";
const FLAG_LOG_UI = "DLP_DEV_LOG_ENABLE";

function $(id){ return document.getElementById(id); }

async function readLogs() {
  // Prefer OfficeRuntime.storage (shared)
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      const raw = await OfficeRuntime.storage.getItem(LOG_KEY);
      if (raw) return JSON.parse(raw);
    }
  } catch (e) {}

  // Fallback localStorage
  try {
    const raw = localStorage.getItem(LOG_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e2) {
    return [];
  }
}

async function writeFlagUi(v) {
  const val = v ? "1" : "0";
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.setItem) {
      await OfficeRuntime.storage.setItem(FLAG_LOG_UI, val);
      return;
    }
  } catch (e) {}
  try { localStorage.setItem(FLAG_LOG_UI, val); } catch (e2) {}
}

async function readFlagUi() {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      const v = await OfficeRuntime.storage.getItem(FLAG_LOG_UI);
      if (v !== null && v !== undefined) return String(v) === "1";
    }
  } catch (e) {}
  try { return localStorage.getItem(FLAG_LOG_UI) === "1"; } catch (e2) { return false; }
}

async function clearLogs() {
  const empty = JSON.stringify([]);
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.setItem) {
      await OfficeRuntime.storage.setItem(LOG_KEY, empty);
    }
  } catch (e) {}
  try { localStorage.setItem(LOG_KEY, empty); } catch (e2) {}
}

function downloadText(filename, text) {
  const blob = new Blob([text], { type:"text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 0);
}

async function render() {
  const arr = await readLogs();
  const lines = (arr || []).map(e => {
    const extra = e.extra ? " " + JSON.stringify(e.extra) : "";
    return `${e.ts || ""} [${e.lvl || "INF"}] [${e.tx || ""}] ${e.msg || ""}${extra}`;
  });
  $("out").value = lines.join("\n");
  $("out").scrollTop = $("out").scrollHeight;
}

Office.onReady().then(async () => {
  $("chkUi").checked = await readFlagUi();

  $("chkUi").onchange = async () => {
    await writeFlagUi($("chkUi").checked);
  };

  $("btnClear").onclick = async () => { await clearLogs(); await render(); };
  $("btnRefresh").onclick = render;

  $("btnDownload").onclick = async () => {
    const arr = await readLogs();
    const lines = (arr || []).map(e => {
      const extra = e.extra ? " " + JSON.stringify(e.extra) : "";
      return `${e.ts || ""} [${e.lvl || "INF"}] [${e.tx || ""}] ${e.msg || ""}${extra}`;
    }).join("\n");
    downloadText("dlp-dev-logs.txt", lines);
  };

  setInterval(render, 1000);
  await render();
});
