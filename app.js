/* global Office, OfficeRuntime */

"use strict";

/**
 * appvprem.js - classic Outlook (outlook.exe) ItemSend
 * - Ping timeout: 30000 ms (Forcepoint)
 * - POST timeout: 35000 ms (Forcepoint)
 * - NO retry POST (retry can break agent confirm/session)
 * - Logs written to OfficeRuntime.storage (shared) + fallback localStorage
 */

const VERSION = "v1.2-prem-forcepoint-timeouts";

const PORT_WIN = 55299;
const PORT_MAC = 55296;

const PING_TIMEOUT_MS = 30000;   // as original
const POST_TIMEOUT_MS = 35000;   // as original
const FIELD_TIMEOUT_MS = 3000;
const BODY_TIMEOUT_MS = 5000;
const ATTS_LIST_TIMEOUT_MS = 5000;
const ATT_CONTENT_TIMEOUT_MS = 30000;

const LOG_KEY = "DLP_DIAG_LOGS";           // array of entries
const FLAG_LOG_UI = "DLP_DEV_LOG_ENABLE";  // "1" / "0"

let urlDseRoot = "https://localhost:55299/";
let logUiEnable = false;

let logBuf = [];
let flushTimer = null;
let flushing = false;
const MAX_LOGS = 400;

Office.initialize = function () {};

function osPlatform() {
  try {
    const p = Office.context.diagnostics.platform;
    if (p === "Mac") return "Mac";
    if (p === "PC" || p === "OfficeOnline") return "Win";
  } catch (e) {}
  return "Other";
}

async function loadFlags() {
  // Prefer OfficeRuntime.storage to share with diagnostics
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      const v = await OfficeRuntime.storage.getItem(FLAG_LOG_UI).catch(() => null);
      if (v !== null) logUiEnable = (String(v) === "1");
      return;
    }
  } catch (e) {}

  try { logUiEnable = (localStorage.getItem(FLAG_LOG_UI) === "1"); } catch (e2) {}
}

function nowIso() { return new Date().toISOString(); }

function pushLog(level, tx, msg, extra) {
  const entry = { ts: nowIso(), lvl: level, tx, msg: String(msg), extra: extra || null };

  try {
    const line = `${entry.ts} [${entry.lvl}] [${entry.tx}] ${entry.msg}`;
    if (level === "ERR") console.error(line, entry.extra || "");
    else if (level === "DBG") console.debug(line, entry.extra || "");
    else console.log(line, entry.extra || "");
  } catch (e) {}

  // UI progressIndicator (optional, no sleep/busy wait)
  if (logUiEnable && typeof msg === "string") {
    try {
      Office.context.mailbox.item.notificationMessages.replaceAsync("dlpDev", {
        type: "progressIndicator",
        message: msg.substring(0, Math.min(msg.length, 250)),
      });
    } catch (e2) {}
  }

  logBuf.push(entry);
  if (logBuf.length > MAX_LOGS) logBuf = logBuf.slice(logBuf.length - MAX_LOGS);

  if (!flushTimer) {
    flushTimer = setTimeout(() => {
      flushTimer = null;
      flushLogs().catch(() => {});
    }, 250);
  }
}

async function flushLogs() {
  if (flushing) return;
  flushing = true;
  try {
    const raw = JSON.stringify(logBuf);

    // write shared storage
    try {
      if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.setItem) {
        await OfficeRuntime.storage.setItem(LOG_KEY, raw);
      }
    } catch (e) {}

    // fallback localStorage
    try { localStorage.setItem(LOG_KEY, raw); } catch (e2) {}
  } finally {
    flushing = false;
  }
}

function mkLogger(tx) {
  return {
    inf: (m, x) => pushLog("INF", tx, m, x),
    dbg: (m, x) => pushLog("DBG", tx, m, x),
    err: (m, x) => pushLog("ERR", tx, m, x),
  };
}

function fetchWithTimeout(url, init, timeoutMs) {
  if (typeof AbortController === "undefined") return fetch(url, init);
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);
  const req = Object.assign({}, init || {}, { signal: controller.signal });
  return fetch(url, req).finally(() => clearTimeout(t));
}

// Optional cleanup for Word/MSO junk (keeps Forcepoint behaviour: strip tags only, keep &nbsp;)
function extractBodyAndCleanMso(html) {
  let s = String(html || "");
  s = s.replace(/<head[\s\S]*?<\/head>/gi, "");
  s = s.replace(/<!--\s*\[if[\s\S]*?<!\s*\[endif\]\s*-->/gi, "");
  const m = s.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  return (m && m[1] !== undefined) ? m[1] : s;
}
function normalizeHtmlToPlainText(htmlBody) {
  const cleaned = extractBodyAndCleanMso(htmlBody);
  return cleaned.replace(/<[^>]+>/g, "");
}

function getIfVal(result) {
  return (result && result.status === Office.AsyncResultStatus.Succeeded) ? result.value : "";
}

function completeOnceFactory(event, log) {
  let done = false;
  return function completeOnce(allow, reason) {
    if (done) return;
    done = true;
    log.inf("completed", { allow: !!allow, reason: reason || "" });
    try { event.completed({ allowEvent: !!allow }); } catch (e) {}
  };
}

async function httpServerCheck(log) {
  log.inf("Checking the server");
  const url = urlDseRoot + "FirefoxExt/_1";

  const r = await fetchWithTimeout(url, {
    method: "GET",
    mode: "cors",
    cache: "no-cache",
    credentials: "same-origin",
    redirect: "follow",
    referrerPolicy: "no-referrer",
  }, PING_TIMEOUT_MS);

  if (!r.ok) throw new Error("ping_http_" + r.status);
  log.inf("Server is UP");
}

function handleResponse(data, event, log, completeOnce) {
  log.inf("Handling response from engine");
  const item = Office.context.mailbox.item;

  // Forcepoint semantics: action === 1 => BLOCK, else => ALLOW
  if (data && data["action"] === 1) {
    try {
      item.notificationMessages.addAsync("NoSend", { type: "errorMessage", message: "Blocked by DLP engine" });
    } catch (e) {}
    log.inf("DLP block", data);
    completeOnce(false, "blocked");
  } else {
    log.inf("DLP allow", data);
    completeOnce(true, "allowed");
  }
}

async function sendToClassifier(data, event, log, completeOnce) {
  log.inf("Sending event to classifier");

  // heartbeat co 5s (tylko informacyjne)
  const hb = setInterval(() => {
    log.inf("Waiting for DLP decision (confirm popup may be active)...");
  }, 5000);

  try {
    const url = urlDseRoot + "OutlookAddin";

    const resp = await fetchWithTimeout(url, {
      method: "POST",
      mode: "cors",
      cache: "no-cache",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      redirect: "follow",
      referrerPolicy: "no-referrer",
      body: JSON.stringify(data),
    }, POST_TIMEOUT_MS);

    if (!resp.ok) {
      log.err("Engine returned error", { status: resp.status });
      completeOnce(true, "engine_http_error_fail_open");
      return;
    }

    // Forcepoint does response.json()
    const json = await resp.json().catch(() => null);
    if (!json) {
      log.err("Engine response is not JSON");
      completeOnce(true, "invalid_json_fail_open");
      return;
    }

    handleResponse(json, event, log, completeOnce);
  } catch (e) {
    log.err("Request crashed", { name: e && e.name ? e.name : "error", msg: e && e.message ? e.message : String(e) });
    completeOnce(true, "classify_error_fail_open");
  } finally {
    clearInterval(hb);
  }
}

async function tryPost(event, log, completeOnce, subject, from, to, cc, bcc, location, body, attachments) {
  log.inf("Trying to post");
  const payload = { subject, body, from, to, cc, bcc, location, attachments };
  log.dbg("Payload (truncated)", { subjectLen: (subject || "").length, bodyLen: (body || "").length, attCount: (attachments || []).length });
  await sendToClassifier(payload, event, log, completeOnce);
}

async function postMessage(message, event, log, completeOnce, subject, from, to, cc, bcc, location, body, attachmentsAsyncResult) {
  log.inf("Posting message");

  if (attachmentsAsyncResult !== null) {
    const list = attachmentsAsyncResult.value || [];
    log.dbg("Attachment list size: " + list.length);

    if (list.length > 0 && typeof message.getAttachmentContentAsync === "function") {
      const mapped = await Promise.all(
        list.map(att => new Promise((resolve) => {
          let finished = false;
          const t = setTimeout(() => { if (!finished) { finished = true; resolve(null); } }, ATT_CONTENT_TIMEOUT_MS);

          message.getAttachmentContentAsync(att.id, (data) => {
            if (finished) return;
            finished = true;
            clearTimeout(t);

            try {
              let base64 = data.value.content;
              if (data.value.format !== "base64") {
                base64 = btoa(data.value.content);
                log.dbg("Encoded attachment in base64");
              }
              resolve({ file_name: att.name, data: base64, content_type: att.contentType });
            } catch (e) {
              resolve(null);
            }
          });
        }))
      );

      await tryPost(event, log, completeOnce, subject, from, to, cc, bcc, location, body, mapped.filter(Boolean));
      return;
    }
  }

  await tryPost(event, log, completeOnce, subject, from, to, cc, bcc, location, body, []);
}

async function validate(event, log, completeOnce) {
  const message = Office.context.mailbox.item;
  const isAppointment = message.itemType === "appointment";
  log.inf(`Validating ${isAppointment ? "appointment" : "message"}`);

  const fields = isAppointment ? [
    message.subject.getAsync.bind(message.subject),
    message.organizer.getAsync.bind(message.organizer),
    message.requiredAttendees.getAsync.bind(message.requiredAttendees),
    message.optionalAttendees.getAsync.bind(message.optionalAttendees),
    message.location.getAsync.bind(message.location)
  ] : [
    message.subject.getAsync.bind(message.subject),
    message.from.getAsync.bind(message.from),
    message.to.getAsync.bind(message.to),
    message.cc.getAsync.bind(message.cc),
    message.bcc.getAsync.bind(message.bcc)
  ];

  const values = await Promise.all([
    new Promise((resolve, reject) => { httpServerCheck(log).then(resolve).catch(reject); }),

    ...fields.map(fn => new Promise(resolve => {
      setTimeout(() => resolve(""), FIELD_TIMEOUT_MS);
      fn(result => resolve(getIfVal(result)));
    })),

    new Promise(resolve => {
      setTimeout(() => resolve(""), BODY_TIMEOUT_MS);
      message.body.getAsync(Office.CoercionType.Html, {}, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const htmlBody = result.value || "";
          log.dbg("=== Raw HTML Body ===");
          log.dbg(htmlBody);
          const plain = normalizeHtmlToPlainText(htmlBody);
          log.dbg("=== Normalized Text ===");
          log.dbg(plain);
          resolve(plain);
        } else {
          resolve("");
        }
      });
    }),

    new Promise(resolve => {
      setTimeout(() => resolve(null), ATTS_LIST_TIMEOUT_MS);
      if (typeof message.getAttachmentsAsync !== "function") { resolve(null); return; }
      message.getAttachmentsAsync(result => {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) resolve(result);
        else resolve(null);
      });
    })
  ]);

  const [alive, ...rest] = values;
  const [subject, from, to, cc, bcc, location, body, attachments] = isAppointment
    ? [rest[0], rest[1], rest[2], rest[3], "", rest[4], rest[5], rest[6]]
    : [rest[0], rest[1], rest[2], rest[3], rest[4], "", rest[5], rest[6]];

  await postMessage(message, event, log, completeOnce, subject, from, to, cc, bcc, location, body, attachments);
}

function handleError(err, event, log, completeOnce) {
  log.err("handleError", { err: err && err.message ? err.message : String(err) });
  // Forcepoint-like fail-open
  completeOnce(true, "error_fail_open");
}

function onMessageSendHandler(event) {
  Office.onReady().then(async () => {
    const tx = "TX-" + Date.now() + "-" + Math.floor(Math.random() * 1000000);
    await loadFlags();
    const log = mkLogger(tx);
    const completeOnce = completeOnceFactory(event, log);

    log.inf(`FP email validation started - [${VERSION}]`);

    const os = osPlatform();
    if (os === "Mac") {
      urlDseRoot = `https://localhost:${PORT_MAC}/`;
      log.inf("MacOS detected");
    } else if (os === "Win") {
      urlDseRoot = `https://localhost:${PORT_WIN}/`;
      log.inf("WindowsOS detected");
    } else {
      log.err("OS is not MacOS or WindowsOS");
      completeOnce(true, "unsupported_os_fail_open");
      return;
    }

    validate(event, log, completeOnce).catch(err => handleError(err, event, log, completeOnce));
  });
}

window.onMessageSendHandler = onMessageSendHandler;
