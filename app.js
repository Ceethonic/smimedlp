/* global Office */

/**
 * appvprem.js - Forcepoint-like (on-prem) for classic Outlook (outlook.exe)
 * Based on uploaded Forcepoint app.js logic (olk.exe), adapted for outlook.exe runtime stability.
 *
 * Key behaviors preserved:
 * - Ping: GET https://localhost:<port>/FirefoxExt/_1
 * - Classify: POST https://localhost:<port>/OutlookAddin with JSON payload
 * - Decision: if action === 1 => BLOCK, else => ALLOW
 * - Confirm popup is handled by native agent; add-in waits for response (heartbeat logs)
 *
 * Dev toggles:
 * - localStorage DLP_DEV_LOG_ENABLE="1" -> also shows progressIndicator (throttled)
 * - localStorage DLP_DEV_NO_CONTENT_TYPE="1" -> omit Content-Type header to reduce preflight issues
 */

let logEnable = true;

// Default ports as in Forcepoint behavior
// - Windows: 55299
// - Mac: 55296
let urlDseRoot = "https://localhost:55299/";

// Optional compat: omit Content-Type to avoid preflight/CORS issues in some classic Outlook setups
let compatNoContentType = false;

// Throttle UI progressIndicator updates
let _lastUiLogTs = 0;

(function loadDevToggles() {
  try {
    logEnable = (localStorage.getItem("DLP_DEV_LOG_ENABLE") === "1");
    compatNoContentType = (localStorage.getItem("DLP_DEV_NO_CONTENT_TYPE") === "1");
  } catch (e) {}
})();

Office.initialize = function () {};

function printLog(text) {
  const line = `[${new Date().toISOString()}] ${text}`;
  try { console.log(line); } catch (e) {}

  // Optional: show in Outlook UI (non-blocking, throttled)
  if (logEnable && (typeof text === "string" || text instanceof String)) {
    const now = Date.now();
    if (now - _lastUiLogTs > 500) {
      _lastUiLogTs = now;
      try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("succeeded", {
          type: "progressIndicator",
          message: String(text).substring(0, Math.min(String(text).length, 250)),
        });
      } catch (e2) {}
    }
  }
}

function handleError(data, event) {
  try { printLog(String(data)); } catch (e) {}
  printLog("Completing event (fail-open)");
  try { event.completed({ allowEvent: true }); } catch (e2) {}
  printLog("Event Completed");
}

function operatingSytem() {
  // In classic Outlook it's typically "PC" (Windows) or "Mac"
  try {
    const platform = Office.context.diagnostics.platform;
    if (platform === "Mac") return "MacOS";
    if (platform === "PC" || platform === "OfficeOnline") return "WindowsOS";
    return "Other";
  } catch (e) {
    return "Other";
  }
}

/**
 * Outlook.exe sometimes returns full Word HTML doc including <head> and MSO conditional XML comments.
 * Forcepoint original does: html.replace(/<[^>]+>/g,'') which can leak junk text.
 * Fix: remove <head>, remove MSO conditional comments, try to extract <body>, then do same strip-tags.
 * IMPORTANT: We keep entities like &nbsp; literally (do NOT decode), matching Forcepoint behavior.
 */
function extractBodyAndCleanMso(html) {
  let s = String(html || "");
  // drop head content
  s = s.replace(/<head[\s\S]*?<\/head>/gi, "");
  // drop MSO conditional comments (contain <xml> blocks that become "Clean/DocumentEmail/false...")
  s = s.replace(/<!--\s*\[if[\s\S]*?<!\s*\[endif\]\s*-->/gi, "");
  // extract body if present
  const m = s.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  if (m && m[1] !== undefined) return m[1];
  return s;
}

function normalizeHtmlToForcepointPlainText(htmlBody) {
  const cleaned = extractBodyAndCleanMso(htmlBody);
  return cleaned.replace(/<[^>]+>/g, "");
}

// fetch with AbortController timeout
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
  }, 30000)
    .then((response) => {
      if (!response.ok) {
        printLog("Server is down");
        reject(false);
      } else {
        printLog("Server is UP");
        resolve(true);
      }
    })
    .catch(() => {
      printLog("Request crashed");
      reject(false);
    });
}

async function sendToClasifier(url = "", data = {}, event) {
  printLog("Sending event to classifier");

  // Heartbeat while agent confirm UI is active / request is pending
  let waitInterval = null;
  const waitStart = setTimeout(() => {
    waitInterval = setInterval(() => {
      printLog("Waiting for DLP decision (confirm popup may be active)...");
    }, 1000);
  }, 1500);

  const headers = {};
  if (!compatNoContentType) headers["Content-Type"] = "application/json";

  fetchWithTimeout(url, {
    method: "POST",
    mode: "cors",
    cache: "no-cache",
    credentials: "same-origin",
    headers: headers,
    redirect: "follow",
    referrerPolicy: "no-referrer",
    body: JSON.stringify(data),
  }, 60000) // allow confirm flows; host-level timeout still applies
    .then(async (response) => {
      printLog("Classifier HTTP status: " + response.status);

      const raw = await response.text().catch(() => "");
      printLog("Engine raw response: " + raw);

      if (!response.ok) {
        printLog("Engine returned error: " + response.status);
        handleError("HTTP_" + response.status, event);
        return null;
      }

      try { return JSON.parse(raw); } catch (e) { return null; }
    })
    .then((respJson) => {
      clearTimeout(waitStart);
      if (waitInterval) clearInterval(waitInterval);

      if (!respJson) {
        printLog("Engine response is not JSON");
        handleError("invalid_json", event);
        return;
      }

      handleResponse(respJson, event);
    })
    .catch((e) => {
      clearTimeout(waitStart);
      if (waitInterval) clearInterval(waitInterval);

      printLog("Request crashed");
      try { printLog(e && e.name ? e.name : "unknown_error"); } catch (_) {}
      handleError(e && e.name ? e.name : "request_crashed", event);
    });
}

function handleResponse(data, event) {
  printLog("Handling response from engine");
  const message = Office.context.mailbox.item;

  // Forcepoint semantics (from attachment):
  // action === 1 => BLOCK
  // else => ALLOW
  if (data["action"] === 1) {
    try {
      message.notificationMessages.addAsync("NoSend", {
        type: "errorMessage",
        message: "Blocked by DLP engine",
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
  const data = { subject, body, from, to, cc, bcc, location, attachments };
  if (attachments) printLog("Attachment list size: " + attachments.length);
  sendToClasifier(urlDseRoot + "OutlookAddin", data, event);
}

async function postMessage(message, event, subject, from, to, cc, bcc, location, body, attachments) {
  printLog("Posting message");

  // attachments: AsyncResult from getAttachmentsAsync OR null
  if (attachments !== null && attachments && attachments.value && attachments.value.length > 0) {
    // If host doesn't support getAttachmentContentAsync, fail gracefully (send no attachments)
    if (typeof message.getAttachmentContentAsync !== "function") {
      tryPost(event, subject, from, to, cc, bcc, location, body, []);
      return;
    }

    await Promise.all(
      attachments.value.map(
        (attachment) =>
          new Promise((resolve) => {
            let resolved = false;

            // Forcepoint-style per attachment timeout
            const t = setTimeout(() => {
              if (resolved) return;
              resolved = true;
              resolve(null);
            }, 30000);

            message.getAttachmentContentAsync(attachment.id, (data) => {
              if (resolved) return;
              resolved = true;
              clearTimeout(t);

              try {
                if (!data || data.status !== Office.AsyncResultStatus.Succeeded || !data.value) {
                  resolve(null);
                  return;
                }

                let base64EncodedContent = data.value.content;

                if (data.value.format !== "base64") {
                  base64EncodedContent = btoa(data.value.content);
                  printLog("Encoded attachment in base64");
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
          })
      )
    ).then((result) => {
      tryPost(event, subject, from, to, cc, bcc, location, body, result.filter(Boolean));
    });
  } else {
    tryPost(event, subject, from, to, cc, bcc, location, body, []);
  }
}

function getIfVal(result) {
  return result && result.status === Office.AsyncResultStatus.Succeeded ? result.value : "";
}

async function validate(event) {
  const message = Office.context.mailbox.item;

  if (message.itemType === "appointment") {
    printLog("Validating appointment");

    await Promise.all([
      new Promise((resolve, reject) => { httpServerCheck(resolve, reject); }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.subject.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.organizer.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.requiredAttendees.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.optionalAttendees.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.location.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 5000);
        message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const htmlBody = result.value;

            printLog("=== Raw HTML Body ===");
            printLog(htmlBody);

            const plainText = normalizeHtmlToForcepointPlainText(htmlBody);

            printLog("=== Normalized Text ===");
            printLog(plainText);

            resolve(plainText);
          } else {
            resolve("");
          }
        });
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(null), 5000);
        if (typeof message.getAttachmentsAsync !== "function") {
          resolve(null);
          return;
        }
        message.getAttachmentsAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
            resolve(result);
            return;
          }
          resolve(null);
        });
      }),
    ])
      .then(([alive, subject, organizer, requiredAttendees, optionalAttendees, location, body, attachments]) => {
        postMessage(message, event, subject, organizer, requiredAttendees, optionalAttendees, [], location, body, attachments);
      })
      .catch(() => {
        handleError("Server might be down", event);
      });

  } else if (message.itemType === "message") {
    printLog("Validating message");

    await Promise.all([
      new Promise((resolve, reject) => { httpServerCheck(resolve, reject); }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.subject.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.from.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.to.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.cc.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 3000);
        message.bcc.getAsync((result) => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(""), 5000);
        message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const htmlBody = result.value;

            printLog("=== Raw HTML Body ===");
            printLog(htmlBody);

            const plainText = normalizeHtmlToForcepointPlainText(htmlBody);

            printLog("=== Normalized Text ===");
            printLog(plainText);

            resolve(plainText);
          } else {
            resolve("");
          }
        });
      }),

      new Promise((resolve) => {
        setTimeout(() => resolve(null), 5000);
        if (typeof message.getAttachmentsAsync !== "function") {
          resolve(null);
          return;
        }
        message.getAttachmentsAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
            resolve(result);
            return;
          }
          resolve(null);
        });
      }),
    ])
      .then(([alive, subject, from, to, cc, bcc, body, attachments]) => {
        postMessage(message, event, subject, from, to, cc, bcc, "", body, attachments);
      })
      .catch((err) => {
        try { printLog(err && err.message ? err.message : "validate_error"); } catch (e) {}
        handleError("Server might be down", event);
      });

  } else {
    printLog("message item type unknown");
    try { printLog(message.itemType); } catch (e) {}
    handleError("Unknown Message Type", event);
  }
}

function onMessageSendHandler(event) {
  // Avoid long Office.onReady waits inside send event (but keep safe fallback)
  function start() {
    printLog("FP email validation started - [v1.2]");

    const os = operatingSytem();
    if (os === "MacOS") {
      printLog("MacOS detected");
      urlDseRoot = "https://localhost:55296/";
      validate(event).catch((err) => handleError(err, event));
    } else if (os === "WindowsOS") {
      printLog("WindowsOS detected");
      urlDseRoot = "https://localhost:55299/";
      validate(event).catch((err) => handleError(err, event));
    } else {
      printLog("OS is not MacOS or WindowsOS");
      handleError("Not MacOS or WindowsOS", event);
    }
  }

  try {
    if (Office && Office.context && Office.context.mailbox) start();
    else Office.onReady().then(start);
  } catch (e) {
    // last resort
    start();
  }
}

// Expose for manifest action mapping
window.onMessageSendHandler = onMessageSendHandler;
