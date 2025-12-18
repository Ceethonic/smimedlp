
let logEnable = false;

let urlDseRoot = 'https://localhost:55299/';


let compatNoContentType = false;

// Throttle UI progressIndicator updates (avoid UI hangs)
let _lastUiLogTs = 0;

// Shared log storage for Diagnostics taskpane
const DLP_LOG_KEY = "DLP_DEV_LOGS_V1";
const DLP_LOG_MAX = 800;
let _bc = null;
try { _bc = new BroadcastChannel("DLP_DEV_LOGS_CH"); } catch (e) {}

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

  // persist logs so diagnostics.html can read them
  try {
    const raw = localStorage.getItem(DLP_LOG_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    const list = Array.isArray(arr) ? arr : [];
    list.push(line);
    while (list.length > DLP_LOG_MAX) list.shift();
    localStorage.setItem(DLP_LOG_KEY, JSON.stringify(list));
  } catch (e) {}

  // stream to diagnostics (if available)
  try { if (_bc) _bc.postMessage(line); } catch (e) {}


  if (logEnable && (typeof text === 'string' || text instanceof String)) {
    const now = Date.now();
    if (now - _lastUiLogTs > 500) {
      _lastUiLogTs = now;
      try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("succeeded", {
          type: "progressIndicator",
          message: text.substring(0, Math.min(text.length, 250)),
        });
      } catch (e) {}
    }
  }
}

function handleError(data, event) {
  printLog(String(data));
  printLog("Completing event (fail-open)");
  try { event.completed({ allowEvent: true }); } catch (e) {}
  printLog("Event Completed");
}

function operatingSytem() {
  try {
    var platform = Office.context.diagnostics.platform;
    if (platform === 'Mac') return 'MacOS';
    if (platform === 'OfficeOnline' || platform === 'PC') return 'WindowsOS';
    return 'Other';
  } catch (e) {
    return 'Other';
  }
}


function extractBodyAndCleanMso(html) {
  let s = String(html || "");
  s = s.replace(/<head[\s\S]*?<\/head>/gi, "");
  s = s.replace(/<!--\s*\[if[\s\S]*?<!\s*\[endif\]\s*-->/gi, "");
  const m = s.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  if (m && m[1] !== undefined) return m[1];
  return s;
}

function normalizeHtmlToForcepointPlainText(htmlBody) {
  const cleaned = extractBodyAndCleanMso(htmlBody);
  return cleaned.replace(/<[^>]+>/g, '');
}


function fetchWithTimeout(url, options, timeoutMs) {
  if (typeof AbortController === "undefined") {
    return fetch(url, options);
  }
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  options = options || {};
  options.signal = controller.signal;
  return fetch(url, options).finally(() => clearTimeout(timeout));
}

async function httpServerCheck(resolve, reject) {
  printLog("Checking the server");

  fetchWithTimeout(urlDseRoot + 'FirefoxExt/_1', {
    method: 'GET',
    mode: 'cors',
    cache: 'no-cache',
    credentials: 'same-origin',
    redirect: 'follow',
    referrerPolicy: 'no-referrer',
  }, 30000).then(response => {
    if (!response.ok) {
      printLog("Server is down");
      reject(false);
    } else {
      printLog("Server is UP");
      resolve(true);
    }
  }).catch(e => {
    printLog("Request crashed");
    reject(false);
  });
}

async function sendToClasifier(url = '', data = {}, event) {
  printLog("Sending event to classifier");


  let waitInterval = null;
  const waitStart = setTimeout(() => {
    waitInterval = setInterval(() => {
      printLog("Waiting for DLP decision (confirm popup may be active)...");
    }, 1000);
  }, 1500);

  const headers = {};
  if (!compatNoContentType) headers['Content-Type'] = 'application/json';

  fetchWithTimeout(url, {
    method: 'POST',
    mode: 'cors',
    cache: 'no-cache',
    credentials: 'same-origin',
    headers: headers,
    redirect: 'follow',
    referrerPolicy: 'no-referrer',
    body: JSON.stringify(data)
  }, 35000).then(async (response) => {
    printLog("Classifier HTTP status: " + response.status);

   
    const raw = await response.text().catch(() => "");
    printLog("Engine raw response: " + raw);

    if (!response.ok) {
      printLog("Engine returned error: " + response.status);
      handleError(response.status, event);
      return null;
    }

    try { return JSON.parse(raw); } catch (e) { return null; }
  }).then((responseJson) => {
    clearTimeout(waitStart);
    if (waitInterval) clearInterval(waitInterval);

    if (!responseJson) {
      printLog("Engine response is not JSON");
      handleError("invalid_json", event);
      return;
    }
    handleResponse(responseJson, event);
  }).catch(e => {
    clearTimeout(waitStart);
    if (waitInterval) clearInterval(waitInterval);

    printLog("Request crashed");
    try { printLog(e && e.name ? e.name : "unknown_error"); } catch (_) {}
    handleError(e && e.name ? e.name : "request_crashed", event);
  });
}

function handleResponse(data, event) {
  printLog("Handling response from engine");
  let message = Office.context.mailbox.item;


  if (data["action"] === 1) {
    try {
      message.notificationMessages.addAsync('NoSend', { type: 'errorMessage', message: 'Blocked by DLP engine' });
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
  let data = { subject, body, from, to, cc, bcc, location, attachments };
  if (attachments) printLog("Attachment list size: " + attachments.length);
  sendToClasifier(urlDseRoot + 'OutlookAddin', data, event);
}

async function postMessage(message, event, subject, from, to, cc, bcc, location, body, attachments) {
  printLog("Posting message");

  // attachments can be null (no attachments) or AsyncResult from getAttachmentsAsync
  if (attachments !== null && attachments && attachments.value && attachments.value.length > 0) {
    await Promise.all(
      attachments.value.map(attachment => new Promise((resolve) => {
        if (typeof message.getAttachmentContentAsync !== "function") {
          resolve(null);
          return;
        }

        let resolved = false;

        const t = setTimeout(() => {
          if (resolved) return;
          resolved = true;
          resolve(null);
        }, 30000);

        message.getAttachmentContentAsync(attachment.id, data => {
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
              content_type: attachment.contentType
            });
          } catch (e) {
            resolve(null);
          }
        });
      }))
    ).then(result => {
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
        setTimeout(() => { resolve(""); }, 3000);
        message.subject.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.organizer.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.requiredAttendees.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.optionalAttendees.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.location.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 5000);
        message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, result => {
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
        setTimeout(() => { resolve(null); }, 5000);

        if (typeof message.getAttachmentsAsync !== "function") {
          resolve(null);
          return;
        }

        message.getAttachmentsAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
            resolve(result);
            return;
          }
          resolve(null);
        });
      })
    ]).then(([alive, subject, organizer, requiredAttendees, optionalAttendees, location, body, attachments]) => {
      postMessage(message, event, subject, organizer, requiredAttendees, optionalAttendees, [], location, body, attachments);
    }).catch(err => {
      handleError("Server might be down", event);
    });

  } else if (message.itemType === "message") {
    printLog("Validating message");

    await Promise.all([
      new Promise((resolve, reject) => { httpServerCheck(resolve, reject); }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.subject.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.from.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.to.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.cc.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 3000);
        message.bcc.getAsync(result => resolve(getIfVal(result)));
      }),

      new Promise((resolve) => {
        setTimeout(() => { resolve(""); }, 5000);
        message.body.getAsync(Office.CoercionType.Html, { asyncContext: event }, result => {
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
        setTimeout(() => { resolve(null); }, 5000);

        if (typeof message.getAttachmentsAsync !== "function") {
          resolve(null);
          return;
        }

        message.getAttachmentsAsync(result => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.length > 0) {
            resolve(result);
            return;
          }
          resolve(null);
        });
      })
    ]).then(([alive, subject, from, to, cc, bcc, body, attachments]) => {
      postMessage(message, event, subject, from, to, cc, bcc, "", body, attachments);
    }).catch(err => {
      try { printLog(err.message); } catch (e) {}
      handleError("Server might be down", event);
    });

  } else {
    printLog("message item type unknown");
    printLog(message.itemType);
    handleError("Unknown Message Type", event);
  }
}

function onMessageSendHandler(event) {
  Office.onReady().then(function () {
    printLog("FP email validation started - [v1.2]");

    var os = operatingSytem();
    if (os === "MacOS") {
      printLog("MacOS detected");
      urlDseRoot = 'https://localhost:55296/';
      validate(event).catch(err => { handleError(err, event); });
    } else if (os === "WindowsOS") {
      printLog("WindowsOS detected");
      urlDseRoot = 'https://localhost:55299/';
      validate(event).catch(err => { handleError(err, event); });
    } else {
      printLog("OS is not MacOS or WindowsOS");
      handleError("Not MacOS or WindowsOS", event);
    }
  });
}

window.onMessageSendHandler = onMessageSendHandler;
