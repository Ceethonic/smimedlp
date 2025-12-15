/* global Office, OfficeRuntifme */

const LOG_KEY = "DLP_DIAG_LOGS";
const CFG_LOG_ENABLE = "DLP_DEV_LOG_ENABLE";
const CFG_NO_CT = "DLP_DEV_NO_CONTENT_TYPE";

let bc = null;
try { if (typeof BroadcastChannel !== "undefined") bc = new BroadcastChannel("dlp_diag"); } catch (e) {}

function $(id) { return document.getElementById(id); }

function safeParse(s) { try { return s ? JSON.parse(s) : []; } catch (e) { return []; } }

async function getLogs() {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      const raw = await OfficeRuntime.storage.getItem(LOG_KEY);
      return safeParse(raw);
    }
  } catch (e) {}
  // fallback localStorage (if storage not available)
  return safeParse(localStorage.getItem(LOG_KEY));
}

async function setFlag(key, val) {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.setItem) {
      await OfficeRuntime.storage.setItem(key, String(val));
      return;
    }
  } catch (e) {}
  try { localStorage.setItem(key, String(val)); } catch (e2) {}
}

async function clearLogs() {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.setItem) {
      await OfficeRuntime.storage.setItem(LOG_KEY, JSON.stringify([]));
      return;
    }
  } catch (e) {}
  try { localStorage.setItem(LOG_KEY, JSON.stringify([])); } catch (e2) {}
}

function render(arr) {
  const lines = (arr || []).map(e => {
    const extra = e.extra ? " " + (typeof e.extra === "string" ? e.extra : JSON.stringify(e.extra)) : "";
    return `${e.ts || ""} [${e.lvl || "INF"}] ${e.msg || ""}${extra}`;
  });
  $("out").textContent = lines.join("\n");
}

async function refresh() {
  const logs = await getLogs();
  render(logs);
}

Office.onReady().then(async () => {
  $("btnRefresh").onclick = refresh;
  $("btnClear").onclick = async () => { await clearLogs(); await refresh(); };

  $("btnEnableLogs").onclick = async () => { await setFlag(CFG_LOG_ENABLE, "1"); await refresh(); };
  $("btnDisableLogs").onclick = async () => { await setFlag(CFG_LOG_ENABLE, "0"); await refresh(); };

  $("btnNoCtOn").onclick = async () => { await setFlag(CFG_NO_CT, "1"); await refresh(); };
  $("btnNoCtOff").onclick = async () => { await setFlag(CFG_NO_CT, "0"); await refresh(); };

  if (bc) {
    bc.onmessage = async () => { await refresh(); };
  }

  // poll as fallback
  setInterval(refresh, 1000);
  await refresh();
});
