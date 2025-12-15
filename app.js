/* global Office */

// SMIME DLP DEV for classic Outlook (outlook.exe)
// - ItemSend handler: onMessageSendHandler(event)
// - POST to local agent: https://localhost:55299/OutlookAddin
// - Logs persisted to localStorage (Diagnostics can read)
// - No retry POST (retry breaks agent confirm/session behaviour)
// - Long classify timeout (default 180s) to allow CONFIRM popup time

"use strict";

const VERSION = "v1.3-prem";

// ---- Storage keys (Diagnostics toggles) ----
const K_LOG_ENABLE = "DLP_DEV_LOG_ENABLE";               // "1" / "0"
const K_NO_CONTENT_TYPE = "DLP_DEV_NO_CONTENT_TYPE";     // "1" => omit Content-Type
const K_PRESERVE_BREAKS = "DLP_DEV_PRESERVE_BREAKS";     // "1" => insert \n for block tags
const K_CLASSIFY_TIMEOUT = "DLP_DEV_CLASSIFY_TIMEOUT_MS";// e.g. "180000"
const K_TOTAL_TIMEOUT = "DLP_DEV_TOTAL_TIMEOUT_MS";      // e.g. "240000"
const K_LOG_BUFFER = "DLP_DEV_LOGS";

// ---- Defaults ----
const DEFAULT_CLASSIFY_TIMEOUT_MS = 180000;  // 3 min (CONFIRM może trwać)
const DEFAULT_TOTAL_TIMEOUT_MS = 240000;     // 4 min safety (fail-open)
const SERVER_CHECK_TIMEOUT_MS = 30000;
const ATTACHMENT_READ_TIMEOUT_MS = 30000;

// Forcepoint-like ports
let urlDseRoot = "https://localhost:55299/";

// Settings loaded per-send (so Diagnostics changes apply immediately)
function loadSettings() {
  const get = (k) => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  const bool = (k) => get(k) === "1";
  const int = (k, def) => {
    const v = parseInt(get(k) || "", 10);
    return Number.isFinite(v) && v > 0 ? v : def;
  };

  return {
    logEnable: bool(K_LOG_ENABLE),
    noContentType: bool(K_NO_CONTENT_TYPE),
    preserveBreaks: bool(K_PRESERVE_BREAKS),
    classifyTimeoutMs: int(K_CLASSIFY_TIMEOUT, DEFAULT_CLASSIFY_TIMEOUT_MS),
    totalTimeoutMs: int(K_TOTAL_TIMEOUT, DEFAULT_TOTAL_TIMEOUT_MS),
  };
}

// ---- Log buffer (Diagnostics reads it) ----
function appendToLogBuffer(line) {
  try {
    let buf = localStorage.getItem(K_LOG_BUFFER) || "";
    buf += line + "\n";
    // keep last ~300KB
    const MAX = 300000;
    if (buf.length > MAX) buf = buf.slice(buf.length - MAX);
    localStorage.setItem(K_LOG_BUFFER, buf);
  } catch (e) {}

  // Optional: also try OfficeRuntime.storage (if available)
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.setItem) {
      OfficeRuntime.storage.getItem(K_LOG_BUFFER).then((v) => {
        let buf = (v || "") + line + "\n";
        const MAX = 300000;
        if (buf.length > MAX) buf = buf.slice(buf.length - MAX);
        return OfficeRuntime.storage.setItem(K_LOG_BUFFER, buf);
      }).catch(() => {});
    }
  } catch (e) {}
}

function makeLogger(tx, settings) {
  let lastUiTs = 0;
  return function log(level, msg, obj) {
    const ts = new Date().toISOString();
    let line = `${ts} [${level}] [${tx}] ${msg}`;
    if (obj !== undefined) {
      try {
        const s = JSON.stringify(obj);
        line += " " + (s.length > 600 ? s.slice(0, 600) + "...(trunc)" : s);
      } catch (e) {}
    }

    try { console.log(line); } catch (e) {}
    appendToLogBuffer(line);

    if (settings.logEnable) {
      const now = Date.now();
      if (now - lastUiTs > 700) {
        lastUiTs = now;
        try {
          Office.context.mailbox.item.notificationMessages.replaceAsync("dlpDev", {
            type: "progressIndicator",
            message: (msg || "").toString().slice(0, 250),
          });
        } catch (e) {}
      }
    }
  };
}

function operatingSystem() {
  try {
    const platform = Office.context.diagnostics.platform;
    if (platform === "Mac") return "MacOS";
    if (platform === "PC" || platform === "OfficeOnline") return "WindowsOS";
    return "Other";
  } catch (e) {
    return "Other";
  }
}

// ---- HTML normalization ----
// Goal: similar to Forcepoint, but optionally preserve breaks for readability.
function extractBodyAndCleanMso(html) {
  let s = String(html || "");
  s = s.replace(/<head[\s\S]*?<\/head>/gi, "");
  s = s.replace(/<!--\s*\[if[\s\S]*?<!\s*\[endif\]\s*-->/gi, "");
  const m = s.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  return (m && m[1] !== undefined) ? m[1] : s;
}

function normalizeHtml(htmlBody, preserveBreaks) {
  let s = extractBodyAndCleanMso(htmlBody);

  if (preserveBreaks) {
    s = s.replace(/<\s*br\s*\/?>/gi, "\n");
    s = s.replace(/<\/\s*(p|div|tr|li|h[1-6])\s*>/gi, "\n");
  }

  // Strip tags, keep entities like &nbsp; literally
  return s.replace(/<[^>]+>/g, "");
}

// ---- Networking ----
function fetchWithTimeout(url, options, timeoutMs) {
  if (typeof AbortController === "undefined") return fetch(url, options);
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  const opts = Object.assign({}, options || {}, { signal: controller.signal });
  return fetch(url, opts).finally(() => clearTimeout(timeout));
}

function makeCompleter(event, log) {
  let done = false;
  return function completeOnce(allowEvent, reason) {
    if (done) return;
    done = true;
    log("INF", "completed", { allow: !!allowEvent, reason: reason || "" });
    try { event.completed({ allowEvent: !!allowEvent }); } catch (e) {}
  };
}

async function httpServerCheck(log) {
  log("INF", "Checking the server");
  const url = urlDseRoot + "FirefoxExt/_1";

  const r = await fetchWithTimeout(url, {
    method: "GET",
    mode: "cors",
    cache: "no-cache",
    credentials: "same-origin",
    redirect: "follow",
    referrerPolicy: "no-referrer",
  }, SERVER_CHECK_TIMEOUT_MS);

  if (!r.ok) throw new Error("server_down_http_" + r.status);
  log("INF", "Server is UP");
}

function handleResponse(data, log, completeOnce) {
  log("INF", "Handling response from engine");
  const message = Office.context.mailbox.item;

  // Semantyka jak Forcepoint: action === 1 => BLOCK, reszta => ALLOW
  if (data && data.action === 1) {
    try {
      message.notificationMessages.addAsync("NoSend", {
        type: "errorMessage",
        message: "Blocked by DLP engine",
      });
    } catch (e) {}
    log("INF", "DLP block", data);
    completeOnce(false, "blocked");
  } else {
    log("INF", "DLP allow.", data);
    completeOnce(true, "allowed");
  }
}

async function sendToClassifier(payload, settings, log, completeOnce) {
  log("INF", "Sending event to classifier");

  // Heartbeat while waiting (CONFIRM popup might block the HTTP response)
  let hb = null;
  const hbStart = setTimeout(() => {
    hb = setInterval(() => {
      log("INF", "Waiting for DLP decision (confirm popup may be active)...");
    }, 5000);
  }, 1500);

  const headers = {};
  if (!settings.noContentType) headers["Content-Type"] = "application/json";

  const url = urlDseRoot + "OutlookAddin";

  // Total timeout safety (fail-open) – żeby nie wisieć w nieskończoność
  let totalTimer = null;
  totalTimer = setTimeout(() => {
    log("ERR", "TOTAL timeout exceeded -> fail-open", { ms: settings.totalTimeoutMs });
    if (hb) clearInterval(hb);
    clearTimeout(hbStart);
    completeOnce(true, "total_timeout_fail_open");
  }, settings.totalTimeoutMs);

  try {
    log("DBG", "Payload (truncated)", {
      subjectLen: (payload.subject || "").length,
      bodyLen: (payload.body || "").length,
      attCount: (payload.attachments || []).length,
      noContentType: settings.noContentType,
      classifyTimeoutMs: settings.classifyTimeoutMs,
    });

    const resp = await fetchWithTimeout(url, {
      method: "POST",
      mode: "cors",
      cache: "no-cache",
      credentials: "same-origin",
      headers,
      redirect: "follow",
      referrerPolicy: "no-referrer",
      body: JSON.stringify(payload),
    }, settings.classifyTimeoutMs);

    const raw = await resp.text().catch(() => "");
    log("DBG", "Engine raw response (truncated)", raw.length > 800 ? raw.slice(0, 800) + "...(trunc)" : raw);

    if (!resp.ok) {
      log("ERR", "Engine returned HTTP error", { status: resp.status });
      completeOnce(true, "engine_http_error_fail_open");
      return;
    }

    let json = null;
    try { json = raw ? JSON.parse(raw) : null; } catch (e) {}
    if (!json) {
      log("ERR", "Engine response is not JSON");
      completeOnce(true, "engine_invalid_json_fail_open");
      return;
    }

    handleResponse(json, log, completeOnce);
  } catch (e) {
    // if AbortController fired or host killed the fetch:
    log("ERR", "sendToClassifier failed", { err: (e && e.message) ? e.message : String(e) });
    completeOnce(true, "classify_error_fail_open");
  } finally {
    if (hb) clearInterval(hb);
    clearTimeout(hbStart);
    if (totalTimer) clearTimeout(totalTimer);
  }
}

// ---- Collect data ----
function getIfVal(result) {
  return (result && result.status === Office.AsyncResultStatus.Succeeded) ? result.value : "";
}

function getAsyncWithTimeout(getterFn, ms, fallback) {
  return new Promise((resolve) => {
    let done = false;
    const t = setTimeout(() => {
      if (done) return;
      done = true;
      resolve(fallback);
    }, ms);

    try {
      getterFn((result) => {
        if (done) return;
        done = true;
        clearTimeout(t);
        resolve(getIfVal(result));
      });
    } catch (e) {
      if (done) return;
      done = true;
      clearTimeout(t);
      resolve(fallback);
    }
  });
}

function bodyGetHtmlWithTimeout(message, ms, log, settings) {
  return new Promise((resolve) => {
    let done = false;
    const t = setTimeout(() => {
      if (done) return;
      done = true;
      resolve("");
    }, ms);

    message.body.getAsync(Office.CoercionType.Html, {}, (result) => {
      if (done) return;
      done = true;
      clearTimeout(t);

      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const htmlBody = result.value || "";
        log("DBG", "=== Raw HTML Body ===");
        log("DBG", htmlBody.length > 2000 ? htmlBody.slice(0, 2000) + "...(trunc)" : htmlBody);

        const plainText = normalizeHtml(htmlBody, settings.preserveBreaks);

        log("DBG", "=== Normalized Text ===");
        log("DBG", plainText.length > 2000 ? plainText.slice(0, 2000) + "...(trunc)" : plainText);

        resolve(plainText);
      } else {
        resolve("");
      }
    });
  });
}

function getAttachmentsAsyncSafe(message) {
  return new Promise((resolve) => {
    if (typeof message.getAttachmentsAsync !== "function") {
      resolve(null);
      return;
    }
    message.getAttachmentsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
        resolve(result.value);
        return;
      }
      resolve(null);
    });
  });
}

function readAttachmentContentSafe(message, attachment) {
  return new Promise((resolve) => {
    if (typeof message.getAttachmentContentAsync !== "function") {
      resolve(null);
      return;
    }

    let done = false;
    const t = setTimeout(() => {
      if (done) return;
      done = true;
      resolve(null);
    }, ATTACHMENT_READ_TIMEOUT_MS);

    message.getAttachmentContentAsync(attachment.id, (data) => {
      if (done) return;
      done = true;
      clearTimeout(t);

      try {
        if (!data || data.status !== Office.AsyncResultStatus.Succeeded || !data.value) {
          resolve(null);
          return;
        }

        let base64EncodedContent = data.value.content;
        if (data.value.format !== "base64") {
          // fallback (may fail on binary) - dev only
          base64EncodedContent = btoa(data.value.content);
        }

        resolve({
          file_name: attachment.name,
          data: base64EncodedContent,
          content_type: attachment.contentType,
        });
      } catch (e) {
        resolve(null);
      }
    });
  });
}

async function validateAndPost(event, settings, log, completeOnce) {
  const message = Office.context.mailbox.item;

  if (!message) {
    log("ERR", "No mailbox item");
    completeOnce(true, "no_item_fail_open");
    return;
  }

  if (message.itemType !== "message" && message.itemType !== "appointment") {
    log("ERR", "Unknown itemType", { itemType: message.itemType });
    completeOnce(true, "unknown_item_fail_open");
    return;
  }

  log("INF", message.itemType === "message" ? "Validating message" : "Validating appointment");

  // server check first (fast fail)
  try {
    await httpServerCheck(log);
  } catch (e) {
    log("ERR", "Server check failed", { err: e.message || String(e) });
    completeOnce(true, "server_down_fail_open");
    return;
  }

  if (message.itemType === "message") {
    const subject = await getAsyncWithTimeout(message.subject.getAsync.bind(message.subject), 3000, "");
    const from = await getAsyncWithTimeout(message.from.getAsync.bind(message.from), 3000, "");
    const to = await getAsyncWithTimeout(message.to.getAsync.bind(message.to), 3000, "");
    const cc = await getAsyncWithTimeout(message.cc.getAsync.bind(message.cc), 3000, "");
    const bcc = await getAsyncWithTimeout(message.bcc.getAsync.bind(message.bcc), 3000, "");
    const body = await bodyGetHtmlWithTimeout(message, 5000, log, settings);

    const atts = await getAttachmentsAsyncSafe(message);
    const attList = [];

    if (atts && atts.length > 0) {
      log("DBG", "Attachment list size: " + atts.length);
      const contents = await Promise.all(atts.map((a) => readAttachmentContentSafe(message, a)));
      for (const c of contents) if (c) attList.push(c);
    } else {
      log("DBG", "Attachment list size: 0");
    }

    const payload = { subject, body, from, to, cc, bcc, location: "", attachments: attList };
    log("INF", "Posting message");
    log("INF", "Trying to post");
    await sendToClassifier(payload, settings, log, completeOnce);
    return;
  }

  // appointment
  const subject = await getAsyncWithTimeout(message.subject.getAsync.bind(message.subject), 3000, "");
  const organizer = await getAsyncWithTimeout(message.organizer.getAsync.bind(message.organizer), 3000, "");
  const required = await getAsyncWithTimeout(message.requiredAttendees.getAsync.bind(message.requiredAttendees), 3000, "");
  const optional = await getAsyncWithTimeout(message.optionalAttendees.getAsync.bind(message.optionalAttendees), 3000, "");
  const location = await getAsyncWithTimeout(message.location.getAsync.bind(message.location), 3000, "");
  const body = await bodyGetHtmlWithTimeout(message, 5000, log, settings);

  const atts = await getAttachmentsAsyncSafe(message);
  const attList = [];
  if (atts && atts.length > 0) {
    log("DBG", "Attachment list size: " + atts.length);
    const contents = await Promise.all(atts.map((a) => readAttachmentContentSafe(message, a)));
    for (const c of contents) if (c) attList.push(c);
  } else {
    log("DBG", "Attachment list size: 0");
  }

  const payload = {
    subject,
    body,
    from: organizer,
    to: required,
    cc: optional,
    bcc: [],
    location,
    attachments: attList,
  };

  log("INF", "Posting message");
  log("INF", "Trying to post");
  await sendToClassifier(payload, settings, log, completeOnce);
}

// ---- Entry point called by manifest ----
function onMessageSendHandler(event) {
  const tx = "TX-" + Date.now() + "-" + Math.floor(Math.random() * 1000000);
  const settings = loadSettings();
  const log = makeLogger(tx, settings);
  const completeOnce = makeCompleter(event, log);

  log("INF", `FP email validation started - [${VERSION}]`);
  const os = operatingSystem();

  if (os === "MacOS") {
    urlDseRoot = "https://localhost:55296/";
    log("INF", "MacOS detected");
  } else if (os === "WindowsOS") {
    urlDseRoot = "https://localhost:55299/";
    log("INF", "WindowsOS detected");
  } else {
    log("ERR", "Unsupported OS/platform");
    completeOnce(true, "unsupported_os_fail_open");
    return;
  }

  // Run (no retries, no double complete)
  Promise.resolve()
    .then(() => validateAndPost(event, settings, log, completeOnce))
    .catch((e) => {
      log("ERR", "Unhandled exception", { err: (e && e.message) ? e.message : String(e) });
      completeOnce(true, "exception_fail_open");
    });
}

// Required for Outlook to find the function
window.onMessageSendHandler = onMessageSendHandler;

// Office.initialize is optional here but harmless
Office.initialize = function () {};
