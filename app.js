/* global Office */
(function () {
  "use strict";

  // ====== DEV KONFIG (możesz nadpisać w diagnostics pane) ======
  var CFG = {
    agentPort: 55299,
    agentBase: null,          // wyliczane z portu
    pingPath: "FirefoxExt/_1",
    classifyPath: "EmailDataProcess/_1",

    // zachowanie przy błędzie
    failClosed: false,        // DEV: false = fail-open

    // timeouty (krótkie, żeby nie wieszać wysyłki)
    hardTimeoutMs: 8000,      // watchdog całego on-send
    pingTimeoutMs: 1200,
    collectTimeoutMs: 2000,
    classifyTimeoutMs: 3500,

    // logowanie
    debugLevel: 3,            // 0 OFF, 1 ERR, 2 INF, 3 DBG
    persistLocalStorage: true,
    localStorageKeyLogs: "DLP_DEV_LOGS",
    localStorageKeyCfg:  "DLP_DEV_CFG",

    // opcjonalny “log sink” (jeśli masz lokalny serwis, który zapisze log do pliku)
    // np. https://localhost:55299/DevLog/_1
    logSinkUrl: ""
  };

  // ====== STORAGE CFG ======
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

  // ====== LOGGING ======
  var _buf = [];
  function pushLog(level, tx, msg, data) {
    var entry = { ts: new Date().toISOString(), lvl: level, tx: tx, msg: msg, data: data || null };
    _buf.push(entry);
    if (_buf.length > 600) _buf.shift();

    if (CFG.persistLocalStorage) {
      try { localStorage.setItem(CFG.localStorageKeyLogs, JSON.stringify(_buf)); } catch (e) {}
    }

    // console
    try { console.log("[DLP][" + level + "][" + tx + "] " + msg, data || ""); } catch (e) {}

    // opcjonalny sink
    if (CFG.logSinkUrl) {
      try {
        var x = new XMLHttpRequest();
        x.open("POST", CFG.logSinkUrl, true);
        x.timeout = 300;
        x.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        x.send(JSON.stringify(entry));
      } catch (e2) {}
    }
  }
  function logE(tx, msg, data) { if (CFG.debugLevel >= 1) pushLog("ERR", tx, msg, data); }
  function logI(tx, msg, data) { if (CFG.debugLevel >= 2) pushLog("INF", tx, msg, data); }
  function logD(tx, msg, data) { if (CFG.debugLevel >= 3) pushLog("DBG", tx, msg, data); }

  // export dla diagnostyki (gdyby sandbox dzielił kontekst – bywa różnie)
  window.DLP_DEV_DUMP = function () { try { return _buf.slice(); } catch (e) { return []; } };
  window.DLP_DEV_CFG_SAVE = function (newCfg) {
    if (!newCfg) return;
    for (var k in newCfg) if (newCfg.hasOwnProperty(k) && CFG.hasOwnProperty(k)) CFG[k] = newCfg[k];
    CFG.agentBase = "https://localhost:" + CFG.agentPort + "/";
    saveCfg();
  };

  // ====== UI notify (krótko) ======
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

  // ====== Helpers ======
  function makeCompleteOnce(event) {
    var done = false;
    return function (allow, reason) {
      if (done) return;
      done = true;
      try { event.completed({ allowEvent: !!allow }); } catch (e) {}
    };
  }

  function xhr(method, url, payload, timeoutMs, cb) {
    var x = new XMLHttpRequest();
    x.open(method, url, true);
    x.timeout = timeoutMs;
    x.onreadystatechange = function () {
      if (x.readyState !== 4) return;
      if (x.status >= 200 && x.status < 300) cb(null, x.responseText || "");
      else cb(new Error("HTTP " + x.status));
    };
    x.ontimeout = function () { cb(new Error("timeout")); };
    x.onerror = function () { cb(new Error("network")); };
    try {
      if (payload) {
        x.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        x.send(JSON.stringify(payload));
      } else {
        x.send();
      }
    } catch (e) {
      cb(e);
    }
  }

  function collectItemData(maxWaitMs, cb) {
    var item = Office.context.mailbox.item;
    var data = { subject: "", bodyText: "", to: [], cc: [], bcc: [] };

    var pending = 5;
    var done = false;

    function finish() {
      if (done) return;
      done = true;
      cb(data);
    }
    var timer = setTimeout(finish, maxWaitMs);

    function one() {
      pending--;
      if (pending <= 0) {
        clearTimeout(timer);
        finish();
      }
    }

    function mapRecips(arr) {
      var out = [];
      try {
        for (var i = 0; i < (arr || []).length; i++) out.push(arr[i].emailAddress || arr[i].displayName || "");
      } catch (e) {}
      return out;
    }

    try { item.subject.getAsync(function (r) { if (r && r.status === Office.AsyncResultStatus.Succeeded) data.subject = r.value || ""; one(); }); } catch (e) { one(); }
    try { item.body.getAsync(Office.CoercionType.Text, function (r) { if (r && r.status === Office.AsyncResultStatus.Succeeded) data.bodyText = (r.value || ""); one(); }); } catch (e) { one(); }
    try { item.to.getAsync(function (r) { if (r && r.status === Office.AsyncResultStatus.Succeeded) data.to = mapRecips(r.value); one(); }); } catch (e) { one(); }
    try { item.cc.getAsync(function (r) { if (r && r.status === Office.AsyncResultStatus.Succeeded) data.cc = mapRecips(r.value); one(); }); } catch (e) { one(); }
    try { item.bcc.getAsync(function (r) { if (r && r.status === Office.AsyncResultStatus.Succeeded) data.bcc = mapRecips(r.value); one(); }); } catch (e) { one(); }
  }

  // ====== ENTRYPOINT: ItemSend ======
  window.onMessageSendHandler = function onMessageSendHandler(event) {
    var tx = "TX-" + Date.now() + "-" + Math.floor(Math.random() * 1000000);
    var t0 = Date.now();
    var complete = makeCompleteOnce(event);

    logI(tx, "OnSend start", { env: "classic", port: CFG.agentPort });
    notify("informationalMessage", "DLP DEV: sprawdzanie...");

    // watchdog: zawsze zwolnij event
    var watchdog = setTimeout(function () {
      logE(tx, "WATCHDOG fired", { ms: Date.now() - t0 });
      notify("informationalMessage", "DLP DEV: timeout – wysyłka dozwolona (watchdog).");
      complete(true, "watchdog");
    }, CFG.hardTimeoutMs);

    function finish(allow, reason) {
      clearTimeout(watchdog);
      logI(tx, "completed", { allow: allow, reason: reason, ms: Date.now() - t0 });
      complete(allow, reason);
    }

    // 1) ping lokalnego agenta
    var pingUrl = CFG.agentBase + CFG.pingPath;
    xhr("GET", pingUrl, null, CFG.pingTimeoutMs, function (pingErr) {
      if (pingErr) {
        logE(tx, "ping failed", { err: pingErr.message, url: pingUrl });
        if (CFG.failClosed) {
          notify("errorMessage", "DLP DEV: agent nieosiągalny – blokuję (fail-closed).");
          finish(false, "agent_unreachable");
        } else {
          notify("informationalMessage", "DLP DEV: agent nieosiągalny – puszczam (fail-open).");
          finish(true, "agent_unreachable_fail_open");
        }
        return;
      }

      logD(tx, "ping ok", { ms: Date.now() - t0 });

      // 2) pobierz dane (krótko)
      collectItemData(CFG.collectTimeoutMs, function (data) {
        logD(tx, "collect ok", { ms: Date.now() - t0, subjectLen: (data.subject || "").length });

        // 3) klasyfikacja
        var classifyUrl = CFG.agentBase + CFG.classifyPath;
        xhr("POST", classifyUrl, data, CFG.classifyTimeoutMs, function (classErr, respText) {
          if (classErr) {
            logE(tx, "classify failed", { err: classErr.message, url: classifyUrl });
            if (CFG.failClosed) {
              notify("errorMessage", "DLP DEV: błąd klasyfikacji – blokuję (fail-closed).");
              finish(false, "classify_error");
            } else {
              notify("informationalMessage", "DLP DEV: błąd klasyfikacji – puszczam (fail-open).");
              finish(true, "classify_error_fail_open");
            }
            return;
          }

          var allow = true;
          var msg = "";
          try {
            var obj = JSON.parse(respText || "{}");
            allow = (obj.action === 1);  // jak w Forcepoint: action==1 => allow
            msg = obj.msg || "";
          } catch (e) {
            allow = !CFG.failClosed;
            msg = "Invalid response";
          }

          if (allow) {
            notify("informationalMessage", "DLP DEV: OK – wysyłka dozwolona.");
            finish(true, "allowed");
          } else {
            notify("errorMessage", msg ? ("DLP DEV: Zablokowano: " + msg) : "DLP DEV: Zablokowano.");
            finish(false, "blocked");
          }
        });
      });
    });
  };

  try {
    Office.onReady(function () {
      // nic – ale zostawiamy, żeby Office.js się ustabilizował
    });
  } catch (e) {}

})();
