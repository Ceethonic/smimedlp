/* global Office */
(function () {
  "use strict";

  // =========================================================
  // DEV CONFIG (nadpisywana przez localStorage: DLP_DEV_CFG)
  // =========================================================
  var CFG = {
    // Agent (Forcepoint engine) na localhost
    agentPort: 55299,
    agentBase: null,                 // wyliczane

    // Endpointy zgodne z Forcepoint (olk.exe)
    pingPath: "FirefoxExt/_1",
    classifyPath: "OutlookAddin",

    // Fail policy
    failClosed: false,               // DEV: false = fail-open

    // Timeouts (krótkie żeby nie wieszać send)
    hardTimeoutMs: 12000,            // watchdog całego OnSend
    pingTimeoutMs: 1500,
    collectTimeoutMs: 4000,
    classifyTimeoutMs: 7000,

    // CORS/preflight workaround:
    // - "application/json; charset=utf-8" (normalnie, ale generuje preflight OPTIONS)
    // - "text/plain; charset=utf-8" (często bez preflight; serwer musi umieć parsować JSON z text/plain)
    postContentType: "application/json; charset=utf-8",

    // Logging
    debugLevel: 3,                   // 0 OFF, 1 ERR, 2 INF, 3 DBG
    persistLocalStorage: true,
    localStorageKeyLogs: "DLP_DEV_LOGS",
    localStorageKeyCfg: "DLP_DEV_CFG",

    // log sink (opcjonalnie): endpoint, który zapisuje logi do pliku po stronie localhost
    // np. https://localhost:55298/DevLog/_1
    logSinkUrl: "",

    // Kontrola logowania treści
    logBodyHtml: true,               // loguj surowy HTML (DBG)
    maxBodyLogChars: 6000            // limit na dump (żeby nie zabić konsoli)
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

  function pushLog(level, tx, msg, data) {
    var entry = { ts: nowIso(), lvl: level, tx: tx, msg: msg, data: data || null };

    _buf.push(entry);
    if (_buf.length > 800) _buf.shift();

    if (CFG.persistLocalStorage) {
      try { localStorage.setItem(CFG.localStorageKeyLogs, JSON.stringify(_buf)); } catch (e) {}
    }

    // Console formatting similar-ish to your logs
    try {
      var line = entry.ts + " [" + level + "] [" + tx + "] " + msg;
      if (data !== undefined && data !== null) console.log(line, data);
      else console.log(line);
    } catch (e2) {}

    // Optional log sink (fire-and-forget)
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
  // UI notifications (krótkie)
  // =========================================================
  function notify(type, message) {
    try {
      var item = Office.context.mailbox.item;
      item.notificationMessages.replaceAsync("dlp-dev", {
        type: type,
        message: message,
        icon: "icon16",
        persistent: false
      });
    } catch (e) {}
  }

  // =========================================================
  // XHR helper - callback ONCE (fix dublowania HTTP0/network)
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

      if (payload !== undefined && payload !== null) {
        x.send(payload);
      } else {
        x.send();
      }
    } catch (e) {
      finish(e, null);
    }
  }

  // =========================================================
  // HTML -> normalized text (DEV)
  // =========================================================
  function normalizeText(s) {
    if (!s) return "";
    // collapse whitespace + normalize newlines
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
      var text = div.textContent || div.innerText || "";
      return normalizeText(text);
    } catch (e) {
      return "";
    }
  }

  function truncateForLog(s) {
    if (!s) return "";
    var t = String(s);
    if (t.length <= CFG.maxBodyLogChars) return t;
    return t.slice(0, CFG.maxBodyLogChars) + "\n...[truncated]...";
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
      body: "", // normalized text for classifier
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

    // body HTML (for logs + normalization)
    var gotHtml = false;
    var htmlValue = "";

    try {
      item.body.getAsync(Office.CoercionType.Html, function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) {
          gotHtml = true;
          htmlValue = r.value || "";
          if (CFG.debugLevel >= 3 && CFG.logBodyHtml) {
            logD(tx, "=== Raw HTML Body ===");
            logD(tx, truncateForLog(htmlValue));
          }
          var norm = htmlToText(htmlValue);
          if (CFG.debugLevel >= 3) {
            logD(tx, "=== Normalized Text ===");
            logD(tx, truncateForLog(norm));
          }
          data.body = norm;
        }
        one();
      });
    } catch (e) { one(); }

    // recipients
    try {
      item.to.getAsync(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) data.to = mapRecipsToObjects(r.value);
        one();
      });
    } catch (e) { one(); }

    try {
      item.cc.getAsync(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) data.cc = mapRecipsToObjects(r.value);
        one();
      });
    } catch (e) { one(); }

    try {
      item.bcc.getAsync(function (r) {
        if (r && r.status === Office.AsyncResultStatus.Succeeded) data.bcc = mapRecipsToObjects(r.value);
        one();
      });
    } catch (e) { one(); }

    // attachments + location (sync-ish)
    try {
      data.attachments = mapAttachments(item);
      try { logD(tx, "Attachment list size: " + data.attachments.length); } catch (e2) {}
    } catch (e3) {}

    // location (appointment)
    try {
      if (item.itemType === "appointment") {
        // w praktyce często jest string; jeśli nie, zostaw ""
        data.location = item.location || "";
      }
    } catch (e4) {}

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
      // classic Outlook on Windows typically "PC"
      if (plat === "Mac") logI(tx, "MacOS detected");
      else logI(tx, "WindowsOS detected");
    } catch (e) {}

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

      collectItemData(item, tx, function (payload) {
        logD(tx, "Sending event to classifier");

        // 3) send to classifier
        var classifyUrl = CFG.agentBase + CFG.classifyPath;

        // payload as Forcepoint-ish JSON
        var body = JSON.stringify(payload);
        var headers = { "Content-Type": CFG.postContentType };

        xhr("POST", classifyUrl, body, CFG.classifyTimeoutMs, headers, function (classErr, respText) {
          if (classErr) {
            // Status 0 / network usually means CORS/preflight blocked or TLS issue
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

          var allow = true;
          var msg = "";

          try {
            var obj = safeJsonParse(respText || "{}") || {};
            // Forcepoint convention: action==1 => allow
            allow = (obj.action === 1);
            msg = obj.msg || "";
          } catch (e) {
            allow = !CFG.failClosed;
            msg = "Invalid response";
          }

          if (allow) {
            logI(tx, "DLP allow.");
            notify("informationalMessage", "DLP DEV: OK – wysyłka dozwolona.");
            finish(true, "allowed");
          } else {
            logI(tx, "DLP block.", { msg: msg });
            notify("errorMessage", msg ? ("DLP DEV: Zablokowano: " + msg) : "DLP DEV: Zablokowano.");
            finish(false, "blocked");
          }
        });
      });
    });
  };

  // Keep runtime warm
  try { Office.onReady(function () {}); } catch (e) {}

})();
