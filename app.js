/* global Office */
(function () {
  "use strict";

  // =========================================================
  // DEV CONFIG (nadpisywana przez localStorage: DLP_DEV_CFG)
  // =========================================================
  var CFG = {
    // Agent (Forcepoint engine) na localhost
    agentPort: 55299,
    agentBase: null, // wyliczane

    // Endpointy zgodne z Forcepoint (olk.exe)
    pingPath: "FirefoxExt/_1",
    classifyPath: "OutlookAddin",

    // Fail policy (dla błędów sieci/parsowania itp.)
    // false = fail-open (domyślnie DEV), true = fail-closed
    failClosed: false,

    // Timeouts
    hardTimeoutMs: 12000,    // watchdog całego OnSend
    pingTimeoutMs: 1500,
    collectTimeoutMs: 4000,
    classifyTimeoutMs: 7000,

    // CORS/preflight workaround:
    // - application/json (normalnie, ale robi preflight OPTIONS)
    // - text/plain (często bez preflight; serwer musi umieć parsować JSON z text/plain)
    postContentType: "application/json; charset=utf-8",

    // Logging
    debugLevel: 3, // 0 OFF, 1 ERR, 2 INF, 3 DBG
    persistLocalStorage: true,
    localStorageKeyLogs: "DLP_DEV_LOGS",
    localStorageKeyCfg: "DLP_DEV_CFG",

    // opcjonalny log sink (endpoint na localhost zapisujący log do pliku)
    logSinkUrl: "",

    // Kontrola logowania treści
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

  // Exposed for diagnostics pane
  window.DLP_DEV_CFG_SAVE = function (newCfg) {
    if (!newCfg) return;
    for (var k in newCfg) if (newCfg.hasOwnProperty(k) && CFG.hasOwnProperty(k)) CFG[k] = newCfg[k];
    CFG.agentBase = "https://localhost:" + CFG.agentPort + "/";
    saveCfg();
  };

  // =========================================================
  // Logging (Forcepoint-like + TX)
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
  // UI notifications
  // =========================================================
  function notify(type, message) {
    try {
      var item = Office.context.mailbox.item;
      item.notificationMessages.replaceAsync("NoSend", {
        type: type,
        message: message,
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
  // HTML -> normalized text
  // =========================================================
  function normalizeText(s) {
    if (!s) return "";
    return String(s)
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .replace(/[ \t]+/g, " ")
      .replace(/\n\s*\n+/g, "\n\n")
      .trim();
  }

  function htmlToText(html) {
    try {
      var div = document.createElement("div");
      div.innerHTML = html;

      // kluczowe: usuń style/script, bo textContent je uwzględnia
      var nodes = div.querySelectorAll("style,script");
      for (var i = 0; i < nodes.length; i++) nodes[i].remove();

      var text = div.textContent || div.innerText || "";
      return normalizeText(text);
    } catch (e) {
      return "";
    }
  }

  // =========================================================
  // Collect message data (Forcepoint-ish payload)
  // =========================================================
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

  function mapAttachments(item) {
    var list = [];
    try {
      var a = item.attachments || [];
      for (var i = 0; i < a.length; i++) {
        list.push({
          id: a[i].id,
          name: a[i].name,
          contentType: a[i].contentType,
          size: a[i].size,
          isInline: a[i].isInline
        });
      }
    } catch (e) {}
    return list;
  }

  function collectItemData(item, tx, cb) {
    var fromProfile = (Office.context.mailbox && Office.context.mailbox.userProfile) ? Office.context.mailbox.userProfile : null;

    var data = {
      subject: "",
      body: "", // normalized text
      from: {
        emailAddress: (fromProfile && fromProfile.emailAddress) ? fromProfile.emailAddress : "",
        displayName: (fromProfile && fromProfile.displayName) ? fromProfile.displayName : ""
      },
      to: [],
      cc: [],
      bcc: [],
      attachments: [],
      location: ""
    };

    var pending = 6;
    var done = false;

    function finish() {
      if (done) return;
      done = true;
      cb(data);
    }

    var timer = setTimeout(function () {
      logE(tx, "collect timeout (partial data)", { pending: pending });
      finish();
    }, CFG.collectTimeoutMs);

    function one() {
      pending--;
      if (pending <= 0) {
        clearTimeout(timer);
        finish();
      }
    }

    // subject
    try {
      item.subject.getAsync(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) data.subject = r.value || "";
        one();
      });
    } catch (e) { one(); }

    // body HTML -> normalized text + debug dumps
    try {
      item.body.getAsync(Office.CoercionType.Html, function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) {
          var html = r.value || "";
          if (CFG.debugLevel >= 3 && CFG.logBodyHtml) {
            logD(tx, "=== Raw HTML Body ===");
            logD(tx, truncateForLog(html));
          }
          var norm = htmlToText(html);
          if (CFG.debugLevel >= 3) {
            logD(tx, "=== Normalized Text ===");
            logD(tx, truncateForLog(norm));
          }
          data.body = norm;
        }
        one();
      });
    } catch (e2) { one(); }

    // recipients
    try {
      item.to.getAsync(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) data.to = mapRecipsToObjects(r.value);
        one();
      });
    } catch (e3) { one(); }

    try {
      item.cc.getAsync(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) data.cc = mapRecipsToObjects(r.value);
        one();
      });
    } catch (e4) { one(); }

    try {
      item.bcc.getAsync(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) data.bcc = mapRecipsToObjects(r.value);
        one();
      });
    } catch (e5) { one(); }

    // attachments + location
    try {
      data.attachments = mapAttachments(item);
      logD(tx, "Attachment list size: " + data.attachments.length);
    } catch (e6) {}

    try {
      if (item.itemType === "appointment") data.location = item.location || "";
    } catch (e7) {}

    one();
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
  // Main OnSend handler (ItemSend - classic outlook.exe)
  // =========================================================
  window.onMessageSendHandler = function onMessageSendHandler(event) {
    var tx = "TX-" + Date.now() + "-" + Math.floor(Math.random() * 1000000);
    var t0 = Date.now();
    var complete = makeCompleteOnce(event);

    function ms() { return Date.now() - t0; }

    // Forcepoint-like header logs
    logI(tx, "FP email validation started - [v1.2]");
    try {
      var plat = Office.context && Office.context.diagnostics && Office.context.diagnostics.platform;
      if (plat === "Mac") logI(tx, "MacOS detected");
      else logI(tx, "WindowsOS detected");
    } catch (e0) {}

    logI(tx, "Validating message");
    notify("informationalMessage", "DLP DEV: sprawdzanie...");

    // watchdog – never hang
    var watchdog = setTimeout(function () {
      logE(tx, "WATCHDOG fired", { ms: ms() });
      notify("informationalMessage", "DLP DEV: timeout – wysyłka dozwolona (watchdog).");
      complete(true, "watchdog");
    }, CFG.hardTimeoutMs);

    function finish(allow, reason) {
      clearTimeout(watchdog);
      logI(tx, "completed", { allow: allow, reason: reason, ms: ms() });
      complete(allow, reason);
    }

    // 1) ping server
    logI(tx, "Checking the server");
    var pingUrl = CFG.agentBase + CFG.pingPath;

    xhr("GET", pingUrl, null, CFG.pingTimeoutMs, null, function (pingErr) {
      if (pingErr) {
        logE(tx, "Server is down", { err: pingErr.message, url: pingUrl, ms: ms() });

        if (CFG.failClosed) {
          notify("errorMessage", "DLP DEV: agent nieosiągalny – blokuję (fail-closed).");
          finish(false, "agent_unreachable");
        } else {
          notify("informationalMessage", "DLP DEV: agent nieosiągalny – puszczam (fail-open).");
          finish(true, "agent_unreachable_fail_open");
        }
        return;
      }

      logI(tx, "Server is UP");

      // 2) collect data
      logI(tx, "Posting message");
      logD(tx, "Trying to post");

      var item = Office.context.mailbox.item;

      collectItemData(item, tx, function (payloadObj) {
        logD(tx, "Sending event to classifier");

        // 3) send to classifier
        var classifyUrl = CFG.agentBase + CFG.classifyPath;

        // payload string
        var payloadStr = JSON.stringify(payloadObj);

        var headers = { "Content-Type": CFG.postContentType };

        xhr("POST", classifyUrl, payloadStr, CFG.classifyTimeoutMs, headers, function (classErr, respText) {
          if (classErr) {
            // HTTP 0 / network -> zwykle CORS/preflight/TLS
            logE(tx, "classify failed", { err: classErr.message, url: classifyUrl, ms: ms() });

            if (CFG.failClosed) {
              notify("errorMessage", "DLP DEV: błąd klasyfikacji – blokuję (fail-closed).");
              finish(false, "classify_error");
            } else {
              notify("informationalMessage", "DLP DEV: błąd klasyfikacji – puszczam (fail-open).");
              finish(true, "classify_error_fail_open");
            }
            return;
          }

          // Handle response from engine
          logI(tx, "Handling response from engine");

          if (CFG.debugLevel >= 3) {
            logD(tx, "Engine raw response (truncated)");
            logD(tx, truncateForLog(respText || ""));
          }

          var obj = safeJsonParse(respText || "{}");
          if (!obj) {
            // jeśli odpowiedź nie-JSON
            logE(tx, "Engine response is not JSON");

            if (CFG.failClosed) {
              notify("errorMessage", "Blocked by DLP engine (invalid response).");
              finish(false, "invalid_response_fail_closed");
            } else {
              // fail-open
              logI(tx, "DLP allow. (invalid response -> fail-open)");
              notify("informationalMessage", "DLP DEV: OK – wysyłka dozwolona.");
              finish(true, "invalid_response_fail_open");
            }
            return;
          }

          // =====================================================
          // FORCEPOINT SEMANTICS (jak w Twoim olk.exe):
          // action == 1  => BLOCK
          // else         => ALLOW
          // =====================================================
          var actionVal = obj.action;
          if (typeof actionVal === "string") actionVal = actionVal.trim();

          var isBlock = (actionVal === 1 || actionVal === "1");
          var allow = !isBlock;

          var msg = obj.msg || obj.message || obj.reason || "";

          if (allow) {
            logI(tx, "DLP allow.");
            notify("informationalMessage", "DLP DEV: OK – wysyłka dozwolona.");
            finish(true, "allowed");
          } else {
            logI(tx, "DLP block.", { action: actionVal, msg: msg });
            // jak w olk.exe: stały komunikat, nawet jeśli msg puste
            notify("errorMessage", msg ? ("Blocked by DLP engine: " + msg) : "Blocked by DLP engine");
            finish(false, "blocked");
          }
        });
      });
    });
  };

  // Keep runtime warm
  try { Office.onReady(function () {}); } catch (e) {}

})();
