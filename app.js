/* global Office */

/**
 * appvprem.js (outlook.exe) - Forcepoint-like
 * - logs persisted to OfficeRuntime.storage for diagnostics
 * - POST to agent with auto-fallback (retry without Content-Type) to reduce CORS/preflight issues
 * - confirm is native agent UI: request can be pending -> heartbeat logs
 */

let urlDseRoot = "https://localhost:55299/"; // Windows default
let logEnable = true;
let compatNoContentType = false;

const LOG_KEY = "DLP_DIAG_LOGS";
const CFG_LOG_ENABLE = "DLP_DEV_LOG_ENABLE";
const CFG_NO_CT = "DLP_DEV_NO_CONTENT_TYPE";

const MAX_LOGS = 600;
let _lastUiLogTs = 0;
let _bc = null;

try {
  if (typeof BroadcastChannel !== "undefined") _bc = new BroadcastChannel("dlp_diag");
} catch (e) {}

Office.initialize = function () {};

function _ts() { return new Date().toISOString(); }

function _safe(obj) {
  try { return JSON.stringify(obj); } catch (e) { return String(obj); }
}

function _appendLog(entry) {
  // Broadcast (best effort)
  try { if (_bc) _bc.postMessage(entry); } catch (e) {}

  // Persist (best effort): OfficeRuntime.storage -> fallback localStorage
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      OfficeRuntime.storage.getItem(LOG_KEY).then((raw) => {
        let arr = [];
        try { arr = raw ? JSON.parse(raw) : []; } catch (_) { arr = []; }
        arr.push(entry);
        if (arr.length > MAX_LOGS) arr = arr.slice(arr.length - MAX_LOGS);
        return OfficeRuntime.storage.setItem(LOG_KEY, JSON.stringify(arr));
      }).catch(() => {});
    } else {
      const raw = localStorage.getItem(LOG_KEY);
      let arr = [];
      try { arr = raw ? JSON.parse(raw) : []; } catch (_) { arr = []; }
      arr.push(entry);
      if (arr.length > MAX_LOGS) arr = arr.slice(arr.length - MAX_LOGS);
      localStorage.setItem(LOG_KEY, JSON.stringify(arr));
    }
  } catch (e) {}
}

function printLog(text, extra) {
  const entry = { ts: _ts(), lvl: "INF", msg: String(text), extra: extra || null };
  try { console.log(`[${entry.ts}] ${entry.msg}`, entry.extra || ""); } catch (e) {}
  _appendLog(entry);

  // Optional UI progress indicator (throttled, non-blocking)
  if (logEnable) {
    const now = Date.now();
    if (now - _lastUiLogTs > 500) {
      _lastUiLogTs = now;
      try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("dlpdiag", {
          type: "progressIndicator",
          message: entry.msg.substring(0, 250),
        });
      } catch (e2) {}
    }
  }
}

function printDbg(text, extra) {
  const entry = { ts: _ts(), lvl: "DBG", msg: String(text), extra: extra || null };
  try { console.log(`[${entry.ts}] ${entry.msg}`, entry.extra || ""); } catch (e) {}
  _appendLog(entry);
}

function printErr(text, extra) {
  const entry = { ts: _ts(), lvl: "ERR", msg: String(text), extra: extra || null };
  try { console.error(`[${entry.ts}] ${entry.msg}`, entry.extra || ""); } catch (e) {}
  _appendLog(entry);
}

function handleError(reason, event) {
  printErr("handleError", { reason: String(reason) });
  // Forcepoint-like fail-open
  try { event.completed({ allowEvent: true }); } catch (e) {}
}

function operatingSytem() {
  try {
    const platform = Office.context.diagnostics.platform;
    if (platform === "Mac") return "MacOS";
    if (platform === "PC" || platform === "OfficeOnline") return "WindowsOS";
    return "Other";
  } catch (e) {
    return "Other";
  }
}

// Forcepoint-like: keep entities like &nbsp; literally; strip tags only.
// outlook.exe fix: drop <head> + MSO conditional comments + extract <body> first.
function extractBodyAndCleanMso(html) {
  let s = String(html || "");
  s = s.replace(/<head[\s\S]*?<\/head>/gi, "");
  s = s.replace(/<!--\s*\[if[\s\S]*?<!\s*\[endif\]\s*-->/gi, "");
  const m = s.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  return (m && m[1] !== undefined) ? m[1] : s;
}
function normalizeHtmlToForcepointPlainText(htmlBody) {
  const cleaned = extractBodyAndCleanMso(htmlBody);
  return cleaned.replace(/<[^>]+>/g, "");
}

function fetchWithTimeout(url, options, timeoutMs) {
  if (typeof AbortController === "undefined") return fetch(url, options);
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);
  options = options || {};
  options.signal = controller.signal;
  return fetch(url, options).finally(() => clearTimeout(t));
}

async function httpServerCheck(resolve, reject) {
  printLog("Checking the server");
  fetchWithTimeout(urlDseRoot + "FirefoxExt/_1", {
    method: "GET",
    mode: "cors",
    cache: "no-cache",
    credentials: "same-origin",
    redirect: "follow",
    referrerPolicy: "no-referrer",
  }, 10000).then((resp) => {
    if (!resp.ok) {
      printErr("Server is down", { status: resp.status });
      reject(false);
    } else {
      printLog("Server is UP");
      resolve(true);
    }
  }).catch((e) => {
    printErr("Ping crashed", { err: e && e.name ? e.name : String(e) });
    reject(false);
  });
}

async function postClassifier(url, payload, event) {
  printLog("Sending event to classifier");

  // confirm popup => pending; log heartbeat
  let waitInterval = null;
  const waitStart = setTimeout(() => {
    waitInterval = setInterval(() => {
      printLog("Waiting for DLP decision (confirm popup may be active)...");
    }, 1000);
  }, 1200);

  const body = JSON.stringify(payload);

  async function doPost(withContentType) {
    const headers = {};
    if (withContentType) headers["Content-Type"] = "application/json";

    // If user forced "no Content-Type", obey it
    if (compatNoContentType) {
      return fetchWithTimeout(url, {
        method: "POST",
        mode: "cors",
        cache: "no-cache",
        credentials: "same-origin",
        redirect: "follow",
        referrerPolicy: "no-referrer",
        body: body
      }, 30000);
    }

    return fetchWithTimeout(url, {
      method: "POST",
      mode: "cors",
      cache: "no-cache",
      credentials: "same-origin",
      headers: headers,
      redirect: "follow",
      referrerPolicy: "no-referrer",
      body: body
    }, 30000);
  }

  try {
    // 1) primary try (Content-Type JSON) unless compatNoContentType forces otherwise
    let resp = await doPost(true);

    // read raw (helps diagnose confirm/permit/block)
    const raw = await resp.text().catch(() => "");
    printDbg("Classifier HTTP status", { status: resp.status });
    printDbg("Engine raw response", { raw: raw ? raw.slice(0, 2000) : "" });

    if (!resp.ok) {
      // 2) fallback only for network-ish / preflight-like cases is hard to detect with fetch,
      // but in practice many failures throw before getting here. Still: treat non-2xx as error.
      throw new Error("HTTP_" + resp.status);
    }

    let json = null;
    try { json = JSON.parse(raw); } catch (_) { json = null; }
    if (!json) throw new Error("invalid_json");

    clearTimeout(waitStart); if (waitInterval) clearInterval(waitInterval);
    return json;

  } catch (e1) {
    // 2) fallback retry without Content-Type (if not forced already)
    printErr("POST failed, retry without Content-Type", { err: e1 && e1.message ? e1.message : String(e1) });

    try {
      let resp2 = await doPost(false);
      const raw2 = await resp2.text().catch(() => "");
      printDbg("Classifier retry HTTP status", { status: resp2.status });
      printDbg("Engine raw response (retry)", { raw: raw2 ? raw2.slice(0, 2000) : "" });

      if (!resp2.ok) throw new Error("HTTP_" + resp2.status);

      let json2 = null;
      try { json2 = JSON.parse(raw2); } catch (_) { json2 = null; }
      if (!json2) throw new Error("invalid_json");

      clearTimeout(waitStart); if (waitInterval) clearInterval(waitInterval);
      return json2;

    } catch (e2) {
      clearTimeout(waitStart); if (waitInterval) clearInterval(waitInterval);
      handleError(e2 && e2.message ? e2.message : "post_failed", event);
      return null;
    }
  }
}

function handleResponse(data, event) {
  printLog("Handling response from engine", data);

  const message = Office.context.mailbox.item;

  // Forcepoint semantics (as per original): action === 1 => BLOCK else ALLOW
  if (data && data["action"] === 1) {
    try {
      message.notificationMessages.addAsync("NoSend", {
        type: "errorMessage",
        message: "Blocked by DLP engine"
      });
    } catch (e) {}
    printLog("DLP block");
    try { event.completed({ allowEvent: false }); } catch (e2) {}
  } else {
    printLog("DLP allow");
    try { event.completed({ allowEvent: true }); } catch (e3) {}
  }
}

async function tryPost(event, subject, from, to, cc, bcc, location, body, attachments) {
  printLog("Trying to post");
  const payload = { subject, body, from, to, cc, bcc, location, attachments };
  printDbg("Payload (truncated)", { subjectLen: (subject || "").length, bodyLen: (body || "").length, attCount: attachments ? attachments.length : 0 });

  const url = urlDseRoot + "OutlookAddin";
  const resp = await postClassifier(url, payload, event);
  if (resp) handleResponse(resp, event);
}

async function postMessage(message, event, subject, from, to, cc, bcc, location, body, attachments) {
  printLog("Posting message");

  if (attachments !== null && attachments && attachments.value && attachments.value.length > 0) {
    const list = attachments.value;
    printLog("Attachment list size: " + list.length);

    if (typeof message.getAttachmentContentAsync !== "function") {
      await tryPost(event, subject, from, to, cc, bcc, location, body, []);
      return;
    }

    const mapped = await Promise.all(list.map(att => new Promise((resolve) => {
      let done = false;
      const t = setTimeout(() => { if (!done) { done = true; resolve(null); } }, 30000);

      message.getAttachmentContentAsync(att.id, (r) => {
        if (done) return;
        done = true;
        clearTimeout(t);

        try {
          if (!r || r.status !== Office.AsyncResultStatus.Succeeded || !r.value) {
            resolve(null); return;
          }
          let base64 = r.value.content;
          if (r.value.format !== "base64") {
            // may fail for binary/unicode; keep Forcepoint-like behavior
            base64 = btoa(r.value.content);
            printDbg("Encoded attachment in base64");
          }
          resolve({ file_name: att.name, data: base64, content_type: att.contentType });
        } catch (e) {
          resolve(null);
        }
      });
    })));

    await tryPost(event, subject, from, to, cc, bcc, location, body, mapped.filter(Boolean));
    return;
  }

  await tryPost(event, subject, from, to, cc, bcc, location, body, []);
}

function getIfVal(result) {
  return (result && result.status === Office.AsyncResultStatus.Succeeded) ? result.value : "";
}

async function validate(event) {
  const message = Office.context.mailbox.item;

  if (message.itemType !== "message" && message.itemType !== "appointment") {
    printErr("Unknown itemType", { itemType: message.itemType });
    handleError("unknown_item_type", event);
    return;
  }

  if (message.itemType === "message") {
    printLog("Validating message");

    await Promise.all([
      new Promise((resolve, reject) => httpServerCheck(resolve, reject)),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.subject.getAsync(r => resolve(getIfVal(r)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.from.getAsync(r => resolve(getIfVal(r)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.to.getAsync(r => resolve(getIfVal(r)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.cc.getAsync(r => resolve(getIfVal(r)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.bcc.getAsync(r => resolve(getIfVal(r)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 5000);
        message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, (r) => {
          if (r.status === Office.AsyncResultStatus.Succeeded) {
            const html = r.value || "";
            printDbg("=== Raw HTML Body ===");
            printDbg(html.length > 6000 ? (html.slice(0, 6000) + "\n...[truncated]...") : html);

            const plain = normalizeHtmlToForcepointPlainText(html);
            printDbg("=== Normalized Text ===");
            printDbg(plain.length > 6000 ? (plain.slice(0, 6000) + "\n...[truncated]...") : plain);
            resolve(plain);
          } else {
            resolve("");
          }
        });
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(null), 5000);
        if (typeof message.getAttachmentsAsync !== "function") { resolve(null); return; }
        message.getAttachmentsAsync((r) => {
          if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.length > 0) resolve(r);
          else resolve(null);
        });
      })
    ]).then(([alive, subject, from, to, cc, bcc, body, attachments]) => {
      return postMessage(message, event, subject, from, to, cc, bcc, "", body, attachments);
    }).catch((e) => {
      printErr("Validate failed", { err: e && e.message ? e.message : String(e) });
      handleError("validate_failed", event);
    });

    return;
  }

  // appointment
  printLog("Validating appointment");

  await Promise.all([
    new Promise((resolve, reject) => httpServerCheck(resolve, reject)),

    new Promise((resolve) => {
      setTimeout(() => resolve(""), 3000);
      message.subject.getAsync(r => resolve(getIfVal(r)));
    }),

    new Promise((resolve) => {
      setTimeout(() => resolve(""), 3000);
      message.organizer.getAsync(r => resolve(getIfVal(r)));
    }),

    new Promise((resolve) => {
      setTimeout(() => resolve(""), 3000);
      message.requiredAttendees.getAsync(r => resolve(getIfVal(r)));
    }),

    new Promise((resolve) => {
      setTimeout(() => resolve(""), 3000);
      message.optionalAttendees.getAsync(r => resolve(getIfVal(r)));
    }),

    new Promise((resolve) => {
      setTimeout(() => resolve(""), 3000);
      message.location.getAsync(r => resolve(getIfVal(r)));
    }),

    new Promise((resolve) => {
      setTimeout(() => resolve(""), 5000);
      message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, (r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          const html = r.value || "";
          const plain = normalizeHtmlToForcepointPlainText(html);
          resolve(plain);
        } else resolve("");
      });
    }),

    new Promise((resolve) => {
      setTimeout(() => resolve(null), 5000);
      if (typeof message.getAttachmentsAsync !== "function") { resolve(null); return; }
      message.getAttachmentsAsync((r) => {
        if (r.status === Office.AsyncResultStatus.Succeeded && r.value && r.value.length > 0) resolve(r);
        else resolve(null);
      });
    })
  ]).then(([alive, subject, organizer, req, opt, location, body, attachments]) => {
    return postMessage(message, event, subject, organizer, req, opt, [], location, body, attachments);
  }).catch((e) => {
    printErr("Validate appointment failed", { err: e && e.message ? e.message : String(e) });
    handleError("validate_failed", event);
  });
}

async function loadFlagsBestEffort() {
  // Prefer OfficeRuntime.storage for shared flags across runtimes
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage && OfficeRuntime.storage.getItem) {
      const v1 = await OfficeRuntime.storage.getItem(CFG_LOG_ENABLE).catch(() => null);
      const v2 = await OfficeRuntime.storage.getItem(CFG_NO_CT).catch(() => null);
      if (v1 !== null) logEnable = (String(v1) === "1");
      if (v2 !== null) compatNoContentType = (String(v2) === "1");
      return;
    }
  } catch (e) {}

  // Fallback localStorage
  try {
    logEnable = (localStorage.getItem(CFG_LOG_ENABLE) === "1");
    compatNoContentType = (localStorage.getItem(CFG_NO_CT) === "1");
  } catch (e2) {}
}

function onMessageSendHandler(event) {
  Office.onReady().then(async function () {
    await loadFlagsBestEffort();

    printLog("FP email validation started - [v1.2]");
    const os = operatingSytem();

    if (os === "MacOS") {
      printLog("MacOS detected");
      urlDseRoot = "https://localhost:55296/";
      validate(event);
    } else if (os === "WindowsOS") {
      printLog("WindowsOS detected");
      urlDseRoot = "https://localhost:55299/";
      validate(event);
    } else {
      printErr("OS not supported", { os });
      handleError("unsupported_os", event);
    }
  }).catch(() => {
    // If Office.onReady fails, fail-open
    handleError("office_onready_failed", event);
  });
}

// Expose for manifest action mapping
window.onMessageSendHandler = onMessageSendHandler;
