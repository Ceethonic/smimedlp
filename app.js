/* global Office */
(function () {
  "use strict";

  // =========================================================
  // DEV CONFIG (nadpisywana przez localStorage: DLP_DEV_CFG)
  // =========================================================
  var CFG = {
    agentPort: 55299,
    agentBase: null,

    pingPath: "FirefoxExt/_1",
    classifyPath: "OutlookAddin",

    // przy błędach sieci/parsowania:
    // false = fail-open, true = fail-closed
    failClosed: false,

    // timeouty
    hardTimeoutMs: 12000,
    pingTimeoutMs: 1500,
    // per-field timeouts (jak w olk.exe)
    fieldTimeoutMs: 3000,
    bodyTimeoutMs: 5000,
    attachmentsListTimeoutMs: 5000,
    attachmentContentTimeoutMs: 30000,
    classifyTimeoutMs: 7000,

    // content-type POST (CORS/preflight)
    // application/json; charset=utf-8  -> zwykle preflight
    // text/plain; charset=utf-8        -> często bez preflight (serwer musi przyjąć)
    postContentType: "application/json; charset=utf-8",

    // logging
    debugLevel: 3, // 0 OFF, 1 ERR, 2 INF, 3 DBG
    persistLocalStorage: true,
    localStorageKeyLogs: "DLP_DEV_LOGS",
    localStorageKeyCfg: "DLP_DEV_CFG",

    // opcjonalny log sink na localhost (zapis do pliku)
    logSinkUrl: "",

    // log body
    logBodyHtml: true,
    maxBodyLogChars: 6000
  };

  // =========================================================
  // Storage CFG
  // =========================================================
  function safeJsonParse(s) { try { return JSON.parse(s); } catch (e) { return null; } }

  function loadCfg() {
    try {
      var raw = localStorage.getItem(CFG.localStorageKeyCfg);
      var obj = raw ? safeJsonParse(raw) : null;
      if (!obj) return;
      for (var k in obj) if (obj.hasOwnProperty(k) && CFG.hasOwnProperty(k)) CFG[k] = obj[k];
    } catch (e) {}
  }

  function saveCfg() {
    try { localStorage.setItem(CFG.localStorageKeyCfg, JSON.stringify(CFG)); } catch (e) {}
  }

  loadCfg();
  CFG.agentBase = "https://localhost:" + CFG.agentPort + "/";

  window.DLP_DEV_CFG_SAVE = function (newCfg) {
    if (!newCfg) return;
    for (var k in newCfg) if (newCfg.hasOwnProperty(k) && CFG.hasOwnProperty(k)) CFG[k] = newCfg[k];
    CFG.agentBase = "https://localhost:" + CFG.agentPort + "/";
    saveCfg();
  };

  // =========================================================
  // Logging
  // =========================================================
  var _buf = [];

  function nowIso() {
    try { return new Date().toISOString(); } catch (e) { return "" + (new Date()); }
  }

  function truncateForLog(s) {
    if (!s) return "";
    var t = String(s);
    if (t.length <= CFG.maxBodyLogChars) return t;
    return t.slice(0, CFG.maxBodyLogChars) + "\n...[truncated]...";
  }

  function pushLog(level, tx, msg, data) {
    var entry = { ts: nowIso(), lvl: level, tx: tx, msg: msg, data: data || null };

    _buf.push(entry);
    if (_buf.length > 800) _buf.shift();

    if (CFG.persistLocalStorage) {
      try { localStorage.setItem(CFG.localStorageKeyLogs, JSON.stringify(_buf)); } catch (e) {}
    }

    try {
      var line = entry.ts + " [" + level + "] [" + tx + "] " + msg;
      if (data !== undefined && data !== null) console.log(line, data);
      else console.log(line);
    } catch (e2) {}

    if (CFG.logSinkUrl) {
      try {
        var x = new XMLHttpRequest();
        x.open("POST", CFG.logSinkUrl, true);
        x.timeout = 300;
        x.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        x.send(JSON.stringify(entry));
      } catch (e3) {}
    }
  }

  function logE(tx, msg, data) { if (CFG.debugLevel >= 1) pushLog("ERR", tx, msg, data); }
  function logI(tx, msg, data) { if (CFG.debugLevel >= 2) pushLog("INF", tx, msg, data); }
  function logD(tx, msg, data) { if (CFG.debugLevel >= 3) pushLog("DBG", tx, msg, data); }

  window.DLP_DEV_DUMP = function () { try { return _buf.slice(); } catch (e) { return []; } };

  // =========================================================
  // Outlook notifications (jak Forcepoint)
  // =========================================================
  function notifyBlocked(messageText) {
    try {
      var item = Office.context.mailbox.item;
      item.notificationMessages.addAsync("NoSend", {
        type: "errorMessage",
        message: messageText || "Blocked by DLP engine"
      });
    } catch (e) {}
  }

  function notifyInfo(messageText) {
    try {
      var item = Office.context.mailbox.item;
      item.notificationMessages.replaceAsync("DlpInfo", {
        type: "informationalMessage",
        message: messageText || "",
        icon: "icon16",
        persistent: false
      });
    } catch (e) {}
  }

  // =========================================================
  // XHR helper - callback ONCE
  // =========================================================
  function xhr(method, url, payload, timeoutMs, headers, cb) {
    var x = new XMLHttpRequest();
    var done = false;

    function finish(err, text) {
      if (done) return;
      done = true;
      try {
        x.onreadystatechange = null;
        x.onerror = null;
        x.ontimeout = null;
        x.onabort = null;
      } catch (e) {}
      cb(err, text);
    }

    try {
      x.open(method, url, true);
      x.timeout = timeoutMs;

      x.onreadystatechange = function () {
        if (x.readyState !== 4) return;
        if (x.status >= 200 && x.status < 300) finish(null, x.responseText || "");
        else finish(new Error("HTTP " + x.status), null);
      };

      x.onerror = function () { finish(new Error("network"), null); };
      x.ontimeout = function () { finish(new Error("timeout"), null); };
      x.onabort = function () { finish(new Error("abort"), null); };

      if (headers) {
        for (var h in headers) if (headers.hasOwnProperty(h)) x.setRequestHeader(h, headers[h]);
      }

      if (payload !== undefined && payload !== null) x.send(payload);
      else x.send();
    } catch (e) {
      finish(e, null);
    }
  }

  // =========================================================
  // HTML -> plain text (bardziej jak olk.exe + fix spacji)
  // =========================================================
  function decodeHtmlEntities(s) {
    try {
      // szybki decode: &nbsp; &lt; itd.
      var ta = document.createElement("textarea");
      ta.innerHTML = s;
      return ta.value;
    } catch (e) {
      return s;
    }
  }

  function htmlToPlainTextForcepointLike(html) {
    var s = String(html || "");

    // log raw
    // (logowanie robimy wyżej, tutaj tylko transformacja)

    // usuń style/script (żeby nie wpychać CSS do body)
    s = s.replace(/<style[\s\S]*?<\/style>/gi, " ");
    s = s.replace(/<script[\s\S]*?<\/script>/gi, " ");

    // zachowaj separatory zanim wytniesz tagi (żeby słowa się nie sklejały)
    s = s.replace(/<br\s*\/?>/gi, "\n");
    s = s.replace(/<\/(p|div|tr|td|th|li|h[1-6])>/gi, "\n");

    // strip tagów jak w olk.exe (ale po powyższych separatorach)
    s = s.replace(/<[^>]+>/g, "");

    // decode encji + NBSP -> normal space
    s = decodeHtmlEntities(s);
    s = s.replace(/[\u00A0\u2007\u202F]/g, " "); // NBSP/NNBSP itd.

    // normalizacja whitespace:
    s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    s = s.replace(/[ \t]+/g, " ");
    s = s.replace(/\n[ \t]+/g, "\n");
    s = s.replace(/\n{3,}/g, "\n\n");

    return s.trim();
  }

  // =========================================================
  // Helpers: getAsync z timeoutem (Forcepoint-style)
  // =========================================================
  function getAsyncValue(getter, timeoutMs, onValue) {
    var done = false;
    var timer = setTimeout(function () {
      if (done) return;
      done = true;
      onValue(null);
    }, timeoutMs);

    try {
      getter(function (result) {
        if (done) return;
        done = true;
        clearTimeout(timer);
        if (result && result.status === Office.AsyncResultStatus.Succeeded) onValue(result.value);
        else onValue(null);
      });
    } catch (e) {
      if (done) return;
      done = true;
      clearTimeout(timer);
      onValue(null);
    }
  }

  function mapRecipsToObjects(arr) {
    var out = [];
    try {
      for (var i = 0; i < (arr || []).length; i++) {
        out.push({
          emailAddress: arr[i].emailAddress || "",
          displayName: arr[i].displayName || "",
          recipientType: arr[i].recipientType || "user"
        });
      }
    } catch (e) {}
    return out;
  }

  function normalizeFromObj(v) {
    if (!v) return { emailAddress: "", displayName: "" };
    // może już być w formacie {emailAddress, displayName}
    if (v.emailAddress || v.displayName) {
      return { emailAddress: v.emailAddress || "", displayName: v.displayName || "" };
    }
    // lub string
    if (typeof v === "string") {
      return { emailAddress: v, displayName: "" };
    }
    return { emailAddress: "", displayName: "" };
  }

  // =========================================================
  // Attachments: jak w olk.exe (jeśli API dostępne)
  // =========================================================
  function getAttachmentsLikeForcepoint(item, tx, cb) {
    // jeśli API nie istnieje, od razu []
    if (!item || typeof item.getAttachmentsAsync !== "function") {
      cb([]);
      return;
    }

    getAsyncValue(function (done) { item.getAttachmentsAsync(done); }, CFG.attachmentsListTimeoutMs, function (attList) {
      if (!attList || !attList.length) {
        cb([]);
        return;
      }

      logD(tx, "Attachment list size: " + attList.length);

      // jeśli brak API do pobierania contentu
      if (typeof item.getAttachmentContentAsync !== "function") {
        cb([]);
        return;
      }

      var results = [];
      var pending = attList.length;
      var finished = false;

      function finish() {
        if (finished) return;
        finished = true;
        // filtruj null (jak w olk.exe)
        cb(results.filter(Boolean));
      }

      for (var i = 0; i < attList.length; i++) {
        (function (att) {
          var resolved = false;

          // timeout per attachment (jak w olk.exe)
          var t = setTimeout(function () {
            if (resolved) return;
            resolved = true;
            results.push(null);
            pending--;
            if (pending <= 0) finish();
          }, CFG.attachmentContentTimeoutMs);

          try {
            item.getAttachmentContentAsync(att.id, function (r) {
              if (resolved) return;
              resolved = true;
              clearTimeout(t);

              if (r && r.status === Office.AsyncResultStatus.Succeeded && r.value) {
                var val = r.value;
                var base64 = val.content;

                // Forcepoint: jeśli nie base64 -> btoa
                try {
                  if (val.format !== "base64") {
                    base64 = btoa(val.content);
                    logD(tx, "Encoded attachment in base64");
                  }
                } catch (e0) {}

                results.push({
                  file_name: att.name,
                  data: base64,
                  content_type: att.contentType
                });
              } else {
                results.push(null);
              }

              pending--;
              if (pending <= 0) finish();
            });
          } catch (e) {
            if (resolved) return;
            resolved = true;
            clearTimeout(t);
            results.push(null);
            pending--;
            if (pending <= 0) finish();
          }
        })(attList[i]);
      }

      // safety: gdyby nic nie zawołało
      setTimeout(finish, CFG.attachmentContentTimeoutMs + 1000);
    });
  }

  // =========================================================
  // Collect payload (Message/Appointment)
  // =========================================================
  function collectPayload(item, tx, cb) {
    // wspólny szkielet
    var payload = {
      subject: "",
      body: "",
      from: { emailAddress: "", displayName: "" },
      to: [],
      cc: [],
      bcc: [],
      location: "",
      attachments: []
    };

    // body html -> plain text (Forcepoint-like)
    function getBody(done) {
      if (!item || !item.body || typeof item.body.getAsync !== "function") {
        done("");
        return;
      }
      getAsyncValue(
        function (cb2) { item.body.getAsync(Office.CoercionType.Html, cb2); },
        CFG.bodyTimeoutMs,
        function (html) {
          html = html || "";
          if (CFG.debugLevel >= 3 && CFG.logBodyHtml) {
            logD(tx, "=== Raw HTML Body ===");
            logD(tx, truncateForLog(html));
          }
          var plain = htmlToPlainTextForcepointLike(html);
          if (CFG.debugLevel >= 3) {
            logD(tx, "=== Normalized Text ===");
            logD(tx, truncateForLog(plain));
          }
          done(plain);
        }
      );
    }

    // MESSAGE
    if (item.itemType === "message") {
      logI(tx, "Validating message");

      // server check jest robiony wcześniej pingiem, tu tylko dane
      getAsyncValue(function (cb2) { item.subject.getAsync(cb2); }, CFG.fieldTimeoutMs, function (subject) {
        payload.subject = subject || "";

        // from: preferuj message.from.getAsync (jak olk.exe), fallback: userProfile
        var gotFrom = false;
        if (item.from && typeof item.from.getAsync === "function") {
          getAsyncValue(function (cb3) { item.from.getAsync(cb3); }, CFG.fieldTimeoutMs, function (fromVal) {
            gotFrom = true;
            payload.from = normalizeFromObj(fromVal);
            afterFrom();
          });
        } else {
          afterFrom();
        }

        function afterFrom() {
          if (!gotFrom) {
            try {
              var up = Office.context.mailbox.userProfile;
              payload.from = { emailAddress: up.emailAddress || "", displayName: up.displayName || "" };
            } catch (e0) {}
          }

          getAsyncValue(function (cb4) { item.to.getAsync(cb4); }, CFG.fieldTimeoutMs, function (toVal) {
            payload.to = mapRecipsToObjects(toVal || []);

            getAsyncValue(function (cb5) { item.cc.getAsync(cb5); }, CFG.fieldTimeoutMs, function (ccVal) {
              payload.cc = mapRecipsToObjects(ccVal || []);

              getAsyncValue(function (cb6) { item.bcc.getAsync(cb6); }, CFG.fieldTimeoutMs, function (bccVal) {
                payload.bcc = mapRecipsToObjects(bccVal || []);

                getBody(function (bodyText) {
                  payload.body = bodyText || "";

                  getAttachmentsLikeForcepoint(item, tx, function (atts) {
                    payload.attachments = atts || [];
                    cb(payload);
                  });
                });
              });
            });
          });
        }
      });

      return;
    }

    // APPOINTMENT
    if (item.itemType === "appointment") {
      logI(tx, "Validating appointment");

      getAsyncValue(function (cb2) { item.subject.getAsync(cb2); }, CFG.fieldTimeoutMs, function (subject) {
        payload.subject = subject || "";

        // organizer
        if (item.organizer && typeof item.organizer.getAsync === "function") {
          getAsyncValue(function (cb3) { item.organizer.getAsync(cb3); }, CFG.fieldTimeoutMs, function (orgVal) {
            payload.from = normalizeFromObj(orgVal);
            afterOrg();
          });
        } else {
          afterOrg();
        }

        function afterOrg() {
          // required/optional attendees -> mapujemy jako "to"/"cc" (jak w Forcepoint)
          if (item.requiredAttendees && typeof item.requiredAttendees.getAsync === "function") {
            getAsyncValue(function (cb4) { item.requiredAttendees.getAsync(cb4); }, CFG.fieldTimeoutMs, function (req) {
              payload.to = mapRecipsToObjects(req || []);
              afterReq();
            });
          } else {
            afterReq();
          }

          function afterReq() {
            if (item.optionalAttendees && typeof item.optionalAttendees.getAsync === "function") {
              getAsyncValue(function (cb5) { item.optionalAttendees.getAsync(cb5); }, CFG.fieldTimeoutMs, function (opt) {
                payload.cc = mapRecipsToObjects(opt || []);
                afterOpt();
              });
            } else {
              afterOpt();
            }

            function afterOpt() {
              // location
              if (item.location && typeof item.location.getAsync === "function") {
                getAsyncValue(function (cb6) { item.location.getAsync(cb6); }, CFG.fieldTimeoutMs, function (loc) {
                  payload.location = loc || "";
                  afterLoc();
                });
              } else {
                afterLoc();
              }

              function afterLoc() {
                getBody(function (bodyText) {
                  payload.body = bodyText || "";

                  getAttachmentsLikeForcepoint(item, tx, function (atts) {
                    payload.attachments = atts || [];
                    cb(payload);
                  });
                });
              }
            }
          }
        }
      });

      return;
    }

    // fallback
    logE(tx, "message item type unknown", { itemType: item.itemType });
    cb(payload);
  }

  // =========================================================
  // Complete-once
  // =========================================================
  function makeCompleteOnce(event) {
    var done = false;
    return function (allow, reason) {
      if (done) return;
      done = true;
      try { event.completed({ allowEvent: !!allow }); } catch (e) {}
    };
  }

  // =========================================================
  // Handle response (dokładnie jak olk.exe: action==1 => block)
  // =========================================================
  function handleResponse(obj, tx, finish) {
    logI(tx, "Handling response from engine");

    // RAW response (DBG)
    if (CFG.debugLevel >= 3) {
      logD(tx, "Engine raw response (truncated)");
      logD(tx, truncateForLog(JSON.stringify(obj)));
    }

    var actionVal = obj ? obj.action : null;
    if (typeof actionVal === "string") actionVal = actionVal.trim();

    // Forcepoint semantics:
    // action === 1 -> BLOCK
    var isBlock = (actionVal === 1 || actionVal === "1");

    if (isBlock) {
      notifyBlocked("Blocked by DLP engine");
      logI(tx, "DLP block");
      finish(false, "blocked");
    } else {
      logI(tx, "DLP allow");
      finish(true, "allowed");
    }
  }

  // =========================================================
  // Main OnSend handler
  // =========================================================
  window.onMessageSendHandler = function onMessageSendHandler(event) {
    var tx = "TX-" + Date.now() + "-" + Math.floor(Math.random() * 1000000);
    var t0 = Date.now();
    var complete = makeCompleteOnce(event);

    function ms() { return Date.now() - t0; }

    logI(tx, "FP email validation started - [v1.2]");
    try {
      var plat = Office.context && Office.context.diagnostics && Office.context.diagnostics.platform;
      if (plat === "Mac") logI(tx, "MacOS detected");
      else logI(tx, "WindowsOS detected");
    } catch (e0) {}

    notifyInfo("Checking DLP...");

    // watchdog – nigdy nie wieszaj send
    var watchdog = setTimeout(function () {
      logE(tx, "WATCHDOG fired", { ms: ms() });
      notifyInfo("DLP timeout – allow (watchdog).");
      complete(true, "watchdog");
    }, CFG.hardTimeoutMs);

    function finish(allow, reason) {
      clearTimeout(watchdog);
      logI(tx, "completed", { allow: allow, reason: reason, ms: ms() });
      complete(allow, reason);
    }

    // 1) server check
    logI(tx, "Checking the server");
    var pingUrl = CFG.agentBase + CFG.pingPath;

    xhr("GET", pingUrl, null, CFG.pingTimeoutMs, null, function (pingErr) {
      if (pingErr) {
        logE(tx, "Server might be down", { err: pingErr.message, url: pingUrl, ms: ms() });

        if (CFG.failClosed) {
          notifyBlocked("Blocked by DLP engine (server down)");
          finish(false, "server_down_fail_closed");
        } else {
          // fail-open
          finish(true, "server_down_fail_open");
        }
        return;
      }

      logI(tx, "Server is UP");

      // 2) validate/collect data
      var item = Office.context.mailbox.item;
      logI(tx, "Posting message");
      logD(tx, "Trying to post");

      collectPayload(item, tx, function (payloadObj) {
        // DBG: podgląd fragmentu body, żeby widzieć spacje
        if (CFG.debugLevel >= 3) {
          var preview = (payloadObj.body || "").slice(0, 200);
          logD(tx, "Payload body preview", preview);
        }

        logD(tx, "Sending event to classifier");

        // 3) send to classifier
        var classifyUrl = CFG.agentBase + CFG.classifyPath;
        var payloadStr = JSON.stringify(payloadObj);

        var headers = { "Content-Type": CFG.postContentType };

        xhr("POST", classifyUrl, payloadStr, CFG.classifyTimeoutMs, headers, function (classErr, respText) {
          if (classErr) {
            logE(tx, "classify failed", { err: classErr.message, url: classifyUrl, ms: ms() });

            if (CFG.failClosed) {
              notifyBlocked("Blocked by DLP engine (classify error)");
              finish(false, "classify_error_fail_closed");
            } else {
              finish(true, "classify_error_fail_open");
            }
            return;
          }

          // parse JSON
          var obj = safeJsonParse(respText || "");
          if (!obj) {
            // w dev loguj raw string
            if (CFG.debugLevel >= 3) {
              logD(tx, "Engine raw response (string, truncated)");
              logD(tx, truncateForLog(respText || ""));
            }

            logE(tx, "Engine response is not JSON");

            if (CFG.failClosed) {
              notifyBlocked("Blocked by DLP engine (invalid response)");
              finish(false, "invalid_response_fail_closed");
            } else {
              finish(true, "invalid_response_fail_open");
            }
            return;
          }

          handleResponse(obj, tx, finish);
        });
      });
    });
  };

  // keep runtime warm
  try { Office.onReady(function () {}); } catch (e) {}

})();
