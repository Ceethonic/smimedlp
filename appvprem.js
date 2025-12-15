/* global Office */

// SMIME DLP (Outlook.exe) - app.js
// - Collects subject/from/to/cc/bcc/location/body + attachments
// - POSTs to local Forcepoint agent: https://localhost:55299/OutlookAddin
// - Expects JSON response: { action: 0|1 } (0=allow, 1=block)
// - Logs to console + localStorage (view in diagnostics.html)

(() => {
  "use strict";

  const VERSION = "vprem-1.3";
  const LOG_KEY = "smimeDlp.logs.v1";
  const DEBUG_KEY = "smimeDlp.debug.v1";         // "1" enables verbose UI + extra logs
  const POST_TIMEOUT_KEY = "smimeDlp.postTimeoutMs.v1"; // optional override

  let urlDseRoot = "https://localhost:55299/"; // Windows default (outlook.exe)
  const PING_PATH = "FirefoxExt/_1";
  const CLASSIFY_PATH = "OutlookAddin";

  // ---------- Logging (console + storage + BroadcastChannel) ----------
  const MAX_LOG_ITEMS = 2000;

  function isoNow() {
    return new Date().toISOString();
  }

  function safeJson(obj) {
    try { return JSON.stringify(obj); } catch { return "\"<unserializable>\""; }
  }

  function isDebug() {
    try { return localStorage.getItem(DEBUG_KEY) === "1"; } catch { return false; }
  }

  function getPostTimeoutMs() {
    // Default: 120s (confirm can take time). Original Forcepoint used 35s.
    const def = 120000;
    try {
      const v = Number(localStorage.getItem(POST_TIMEOUT_KEY));
      return Number.isFinite(v) && v > 0 ? v : def;
    } catch {
      return def;
    }
  }

  function loadLogBuffer() {
    try {
      const raw = localStorage.getItem(LOG_KEY);
      if (!raw) return [];
      const arr = JSON.parse(raw);
      return Array.isArray(arr) ? arr : [];
    } catch {
      return [];
    }
  }

  function saveLogBuffer(arr) {
    try { localStorage.setItem(LOG_KEY, JSON.stringify(arr)); } catch { /* ignore */ }
  }

  function appendLog(entry) {
    const buf = loadLogBuffer();
    buf.push(entry);
    while (buf.length > MAX_LOG_ITEMS) buf.shift();
    saveLogBuffer(buf);

    try {
      const bc = new BroadcastChannel("smimeDlpLogs");
      bc.postMessage(entry);
      bc.close();
    } catch { /* ignore */ }
  }

  let lastUiTs = 0;
  function uiProgress(message) {
    if (!isDebug()) return;
    const now = Date.now();
    if (now - lastUiTs < 800) return;
    lastUiTs = now;

    try {
      const item = Office.context.mailbox.item;
      item.notificationMessages.replaceAsync("smimeDlpProgress", {
        type: "progressIndicator",
        message: String(message).substring(0, 250),
      });
    } catch { /* ignore */ }
  }

  function printLog(level, msg, meta) {
    const entry = {
      ts: isoNow(),
      level,
      msg: String(msg),
      meta: meta ?? null
    };

    // Console
    try {
      const line = `${entry.ts} [${level}] ${entry.msg}` + (entry.meta ? ` ${safeJson(entry.meta)}` : "");
      if (level === "ERR") console.error(line);
      else if (level === "DBG") console.debug(line);
      else console.log(line);
    } catch { /* ignore */ }

    // Storage
    appendLog(entry);

    // UI (optional)
    uiProgress(entry.msg);
  }

  function logInf(msg, meta) { printLog("INF", msg, meta); }
  function logDbg(msg, meta) { printLog("DBG", msg, meta); }
  function logErr(msg, meta) { printLog("ERR", msg, meta); }

  // ---------- Helpers ----------
  function sleepBusy(ms) {
    const start = Date.now();
    while (Date.now() - start < ms) { /* busy */ }
  }

  function operatingSystem() {
    // Outlook desktop on Windows => "PC"
    // New Outlook / OWA often => "OfficeOnline"
    const platform = Office?.context?.diagnostics?.platform;
    if (platform === "Mac") return "MacOS";
    if (platform === "PC") return "WindowsDesktop";
    if (platform === "OfficeOnline") return "OfficeOnline";
    return "Other";
  }

  function getIfVal(result) {
    return result && result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
  }

  function withTimeout(ms, promise, onTimeoutValue) {
    return new Promise((resolve) => {
      let done = false;
      const t = setTimeout(() => {
        if (done) return;
        done = true;
        resolve(onTimeoutValue);
      }, ms);

      Promise.resolve(promise)
        .then((v) => {
          if (done) return;
          done = true;
          clearTimeout(t);
          resolve(v);
        })
        .catch(() => {
          if (done) return;
          done = true;
          clearTimeout(t);
          resolve(onTimeoutValue);
        });
    });
  }

  function extractBodyHtml(html) {
    // Try to keep only <body>...</body> (classic Outlook can return full Word HTML doc).
    try {
      const lower = html.toLowerCase();
      const b0 = lower.indexOf("<body");
      if (b0 >= 0) {
        const bStart = lower.indexOf(">", b0);
        const bEnd = lower.lastIndexOf("</body>");
        if (bStart >= 0 && bEnd > bStart) {
          html = html.substring(bStart + 1, bEnd);
        }
      }

      // Remove style/script blocks + conditional comments that add noise
      html = html.replace(/<style[\s\S]*?<\/style>/gi, "");
      html = html.replace(/<script[\s\S]*?<\/script>/gi, "");
      html = html.replace(/<!--[\s\S]*?-->/g, "");
      html = html.replace(/<!\[if[\s\S]*?\]>/gi, "");
      html = html.replace(/~~themedata~~/gi, "");
      html = html.replace(/~~colorschememapping~~/gi, "");
      return html;
    } catch {
      return html;
    }
  }

  function normalizeHtmlToText(html) {
    // Forcepoint-like: strip tags, DO NOT decode entities (keeps &nbsp;)
    // Also do NOT trim, to avoid eating intentional spaces.
    try {
      return String(html).replace(/<[^>]+>/g, "");
    } catch {
      return "";
    }
  }

  // ---------- Networking ----------
  async function httpServerCheck() {
    logInf("Checking the server");
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);

    try {
      const resp = await fetch(urlDseRoot + PING_PATH, {
        signal: controller.signal,
        method: "GET",
        mode: "cors",
        cache: "no-cache",
        credentials: "same-origin",
        redirect: "follow",
        referrerPolicy: "no-referrer",
      });
      clearTimeout(timeout);

      if (!resp.ok) {
        logErr("Server is down", { status: resp.status });
        return false;
      }
      logInf("Server is UP");
      return true;
    } catch (e) {
      clearTimeout(timeout);
      logErr("Server check crashed", { name: e?.name, message: e?.message });
      return false;
    }
  }

  async function sendToClassifier(url, data, event, tx) {
    logInf("Sending event to classifier");
    const controller = new AbortController();
    const timeoutMs = getPostTimeoutMs();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);

    // Optional "waiting..." debug heartbeat (no busy-wait)
    let waitTimer = null;
    if (isDebug()) {
      waitTimer = setInterval(() => {
        logInf("Waiting for DLP decision (confirm popup may be active)...", { tx });
      }, 1000);
    }

    try {
      const resp = await fetch(url, {
        signal: controller.signal,
        method: "POST",
        mode: "cors",
        cache: "no-cache",
        credentials: "same-origin",
        headers: { "Content-Type": "application/json" },
        redirect: "follow",
        referrerPolicy: "no-referrer",
        body: JSON.stringify(data),
      });

      if (!resp.ok) {
        throw new Error(`HTTP ${resp.status}`);
      }

      const json = await resp.json();
      clearTimeout(timeout);
      if (waitTimer) clearInterval(waitTimer);
      handleResponse(json, event, tx);
    } catch (e) {
      clearTimeout(timeout);
      if (waitTimer) clearInterval(waitTimer);
      logErr("Classifier request crashed", { tx, name: e?.name, message: e?.message });
      handleError(e, event, tx);
    }
  }

  // ---------- DLP decision ----------
  function handleResponse(data, event, tx) {
    logInf("Handling response from engine", { tx });
    logDbg("Engine raw response (truncated)", { tx, data });

    const action = Number(data && data.action);
    const item = Office.context.mailbox.item;

    if (action === 1) {
      try {
        item.notificationMessages.addAsync("NoSend", {
          type: "errorMessage",
          message: "Blocked by DLP engine"
        });
      } catch { /* ignore */ }

      logInf("DLP block.", { tx });
      event.completed({ allowEvent: false });
      logInf("completed", { allow: false, reason: "blocked", tx });
    } else {
      logInf("DLP allow.", { tx });
      event.completed({ allowEvent: true });
      logInf("completed", { allow: true, reason: "allowed", tx });
    }
  }

  function handleError(err, event, tx) {
    // Fail open to avoid blocking mail if local agent is down
    logErr("handleError", { tx, err: err?.message || String(err) });
    try { event.completed({ allowEvent: true }); } catch { /* ignore */ }
    logInf("completed", { allow: true, reason: "classify_error_fail_open", tx });
  }

  // ---------- Message collection + attachments ----------
  async function tryPost(event, subject, from, to, cc, bcc, location, body, attachments, tx) {
    logInf("Trying to post", { tx });

    const data = { subject, body, from, to, cc, bcc, location, attachments: attachments || [] };

    if (data.attachments) logDbg("Attachment list size: " + data.attachments.length, { tx });

    // Light payload stats (safe)
    logDbg("Payload (truncated)", {
      tx,
      subjectLen: (subject && subject.length) || 0,
      bodyLen: (body && body.length) || 0,
      attCount: (data.attachments && data.attachments.length) || 0,
    });

    sendToClassifier(urlDseRoot + CLASSIFY_PATH, data, event, tx);
  }

  async function postMessage(message, event, subject, from, to, cc, bcc, location, body, attachments, tx) {
    logInf("Posting message", { tx });

    if (attachments !== null && attachments && Array.isArray(attachments.value)) {
      const items = attachments.value;

      const results = await Promise.all(
        items.map((att) => new Promise((resolve) => {
          let done = false;
          const t = setTimeout(() => {
            if (done) return;
            done = true;
            resolve(null);
          }, 30000);

          try {
            message.getAttachmentContentAsync(att.id, (data) => {
              if (done) return;
              done = true;
              clearTimeout(t);

              try {
                let base64EncodedContent = data?.value?.content || "";
                if (data?.value?.format !== "base64") {
                  // NOTE: btoa is ASCII-only; Forcepoint uses it anyway.
                  try {
                    base64EncodedContent = btoa(base64EncodedContent);
                    logDbg("Encoded attachment in base64", { tx, name: att.name });
                  } catch {
                    // last resort: send as-is
                  }
                }

                resolve({
                  file_name: att.name,
                  data: base64EncodedContent,
                  content_type: att.contentType
                });
              } catch {
                resolve(null);
              }
            });
          } catch {
            clearTimeout(t);
            resolve(null);
          }
        }))
      );

      tryPost(event, subject, from, to, cc, bcc, location, body, results.filter(Boolean), tx);
    } else {
      tryPost(event, subject, from, to, cc, bcc, location, body, [], tx);
    }
  }

  async function validate(event, tx) {
    const message = Office.context.mailbox.item;
    const isAppointment = message.itemType === "appointment";
    logInf(`Validating ${isAppointment ? "appointment" : "message"}`, { tx });

    const fields = isAppointment ? [
      message.subject.getAsync.bind(message.subject),
      message.organizer.getAsync.bind(message.organizer),
      message.requiredAttendees.getAsync.bind(message.requiredAttendees),
      message.optionalAttendees.getAsync.bind(message.optionalAttendees),
      message.location.getAsync.bind(message.location),
    ] : [
      message.subject.getAsync.bind(message.subject),
      message.from.getAsync.bind(message.from),
      message.to.getAsync.bind(message.to),
      message.cc.getAsync.bind(message.cc),
      message.bcc.getAsync.bind(message.bcc),
    ];

    const values = await Promise.all([
      httpServerCheck(),

      // Subject/from/to/cc/bcc/location
      ...fields.map((fn) => new Promise((resolve) => {
        let done = false;
        const t = setTimeout(() => {
          if (done) return;
          done = true;
          resolve("");
        }, 3000);

        try {
          fn((result) => {
            if (done) return;
            done = true;
            clearTimeout(t);
            resolve(getIfVal(result));
          });
        } catch {
          if (done) return;
          done = true;
          clearTimeout(t);
          resolve("");
        }
      })),

      // Body HTML normalization
      new Promise((resolve) => {
        let done = false;
        const t = setTimeout(() => {
          if (done) return;
          done = true;
          resolve("");
        }, 5000);

        try {
          message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, (result) => {
            if (done) return;
            done = true;
            clearTimeout(t);

            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const rawHtml = result.value || "";
              logDbg("=== Raw HTML Body ===", { tx });
              logDbg(rawHtml, { tx });

              const extracted = extractBodyHtml(rawHtml);
              const plainText = normalizeHtmlToText(extracted);

              logDbg("=== Normalized Text ===", { tx });
              logDbg(plainText, { tx });

              resolve(plainText);
            } else {
              resolve("");
            }
          });
        } catch {
          if (done) return;
          done = true;
          clearTimeout(t);
          resolve("");
        }
      }),

      // Attachments
      new Promise((resolve) => {
        let done = false;
        const t = setTimeout(() => {
          if (done) return;
          done = true;
          resolve(null);
        }, 5000);

        try {
          message.getAttachmentsAsync((result) => {
            if (done) return;
            done = true;
            clearTimeout(t);

            if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
              resolve(result);
            } else {
              resolve(null);
            }
          });
        } catch {
          if (done) return;
          done = true;
          clearTimeout(t);
          resolve(null);
        }
      }),
    ]);

    const [alive, ...rest] = values;
    if (!alive) throw new Error("Server might be down");

    const [subject, from, to, cc, bcc, location, body, attachments] = isAppointment
      ? [rest[0], rest[1], rest[2], rest[3], "", rest[4], rest[5], rest[6]]
      : [rest[0], rest[1], rest[2], rest[3], rest[4], "", rest[5], rest[6]];

    await postMessage(message, event, subject, from, to, cc, bcc, location, body, attachments, tx);
  }

  // ---------- Entry point ----------
  function onMessageSendHandler(event) {
    const tx = `TX-${Date.now()}-${Math.floor(Math.random() * 1000000)}`;

    Office.onReady().then(() => {
      logInf(`FP email validation started - [${VERSION}]`, { tx });

      const os = operatingSystem();
      if (os === "WindowsDesktop") {
        logInf("WindowsOS detected", { tx });
        urlDseRoot = "https://localhost:55299/";
      } else if (os === "MacOS") {
        logInf("MacOS detected", { tx });
        urlDseRoot = "https://localhost:55296/";
      } else if (os === "OfficeOnline") {
        // New Outlook / OWA – keep Windows port by default unless you want separate routing
        logInf("OfficeOnline detected", { tx });
        urlDseRoot = "https://localhost:55299/";
      } else {
        logErr("OS is not supported", { tx, os });
        handleError("Unsupported OS", event, tx);
        return;
      }

      validate(event, tx).catch((err) => handleError(err, event, tx));
    });
  }

  // Some hosts require Office.actions.associate for event handlers
  try {
    if (Office?.actions?.associate) {
      Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    }
  } catch { /* ignore */ }

  // Also expose as global for older runtimes
  try { window.onMessageSendHandler = onMessageSendHandler; } catch { /* ignore */ }

  // Office.initialize is still referenced by some hosts
  try { Office.initialize = function () {}; } catch { /* ignore */ }
})();
