/* global Office */

(() => {
  "use strict";

  const VERSION = "clean-min-1.0";
  const LOG_KEY = "smimeDlp.logs";
  const DEBUG_KEY = "smimeDlp.debug"; // "1" = pokazuj progressIndicator

  const ROOT_WIN = "https://localhost:55299/";
  const ROOT_MAC = "https://localhost:55296/";
  const PING = "FirefoxExt/_1";
  const POST = "OutlookAddin";

  const PING_TIMEOUT_MS = 30000;
  const POST_TIMEOUT_MS = 35000;         // jak Forcepoint
  const FIELD_TIMEOUT_MS = 3000;
  const BODY_TIMEOUT_MS = 5000;
  const ATTS_LIST_TIMEOUT_MS = 5000;
  const ATT_CONTENT_TIMEOUT_MS = 30000;

  // ---------- logging ----------
  function now() { return new Date().toISOString(); }
  function tx() { return `TX-${Date.now()}-${Math.floor(Math.random() * 1e6)}`; }

  function getDebug() {
    try { return localStorage.getItem(DEBUG_KEY) === "1"; } catch { return false; }
  }

  function pushLog(level, t, msg, meta) {
    const e = { ts: now(), level, tx: t, msg: String(msg), meta: meta ?? null };

    try {
      const line = `${e.ts} [${e.level}] [${e.tx}] ${e.msg}`;
      if (level === "ERR") console.error(line, e.meta || "");
      else if (level === "DBG") console.debug(line, e.meta || "");
      else console.log(line, e.meta || "");
    } catch {}

    try {
      const raw = localStorage.getItem(LOG_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      const list = Array.isArray(arr) ? arr : [];
      list.push(e);
      while (list.length > 400) list.shift();
      localStorage.setItem(LOG_KEY, JSON.stringify(list));
    } catch {}

    if (getDebug()) {
      try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("smimeDlp", {
          type: "progressIndicator",
          message: e.msg.substring(0, 250),
        });
      } catch {}
    }
  }

  const INF = (t, m, x) => pushLog("INF", t, m, x);
  const DBG = (t, m, x) => pushLog("DBG", t, m, x);
  const ERR = (t, m, x) => pushLog("ERR", t, m, x);

  // ---------- helpers ----------
  function platform() {
    try { return Office.context.diagnostics.platform; } catch { return "Unknown"; }
  }
  function rootUrl() {
    return platform() === "Mac" ? ROOT_MAC : ROOT_WIN;
  }

  function fetchTimeout(url, init, ms) {
    const c = new AbortController();
    const to = setTimeout(() => c.abort(), ms);
    return fetch(url, { ...(init || {}), signal: c.signal }).finally(() => clearTimeout(to));
  }

  function withTimeout(ms, work, fallback) {
    return new Promise((resolve) => {
      let done = false;
      const t = setTimeout(() => { if (!done) { done = true; resolve(fallback); } }, ms);
      try {
        work((v) => { if (done) return; done = true; clearTimeout(t); resolve(v); });
      } catch {
        if (done) return; done = true; clearTimeout(t); resolve(fallback);
      }
    });
  }

  // Uwaga: w Forcepoint “Normalized Text” to po prostu strip tagów (bez dekodowania &nbsp;)
  function normalizeHtml(html) {
    let s = String(html || "");
    const m = s.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    if (m && m[1] != null) s = m[1];
    s = s.replace(/<style[\s\S]*?<\/style>/gi, "");
    s = s.replace(/<script[\s\S]*?<\/script>/gi, "");
    s = s.replace(/<!--[\s\S]*?-->/g, "");
    return s.replace(/<[^>]+>/g, "");
  }

  function getVal(r, fallback) {
    return (r && r.status === Office.AsyncResultStatus.Succeeded) ? r.value : fallback;
  }

  function completeOnceFactory(event, t) {
    let done = false;
    return (allow, reason) => {
      if (done) return;
      done = true;
      INF(t, "completed", { allow: !!allow, reason: reason || "" });
      try { event.completed({ allowEvent: !!allow }); } catch {}
    };
  }

  // ---------- network ----------
  async function pingServer(t, root) {
    INF(t, "Checking the server");
    const r = await fetchTimeout(root + PING, {
      method: "GET",
      mode: "cors",
      cache: "no-cache",
      credentials: "same-origin",
      redirect: "follow",
      referrerPolicy: "no-referrer",
    }, PING_TIMEOUT_MS);

    if (!r.ok) throw new Error(`PING_HTTP_${r.status}`);
    INF(t, "Server is UP");
  }

  function handleResponse(t, data, completeOnce) {
    INF(t, "Handling response from engine");
    DBG(t, "Engine raw response", data);

    const action = Number(data && data.action);
    if (action === 1) {
      try {
        Office.context.mailbox.item.notificationMessages.addAsync("NoSend", {
          type: "errorMessage",
          message: "Blocked by DLP engine",
        });
      } catch {}
      INF(t, "DLP block.");
      completeOnce(false, "blocked");
    } else {
      INF(t, "DLP allow.");
      completeOnce(true, "allowed");
    }
  }

  async function postToAgent(t, root, payload, completeOnce) {
    INF(t, "Sending event to classifier");

    // 1 request, bez retry (jak Forcepoint)
    const resp = await fetchTimeout(root + POST, {
      method: "POST",
      mode: "cors",
      cache: "no-cache",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      redirect: "follow",
      referrerPolicy: "no-referrer",
      body: JSON.stringify(payload),
    }, POST_TIMEOUT_MS);

    if (!resp.ok) throw new Error(`POST_HTTP_${resp.status}`);
    const json = await resp.json();
    handleResponse(t, json, completeOnce);
  }

  // ---------- main ----------
  async function onSend(event) {
    const t = tx();
    const root = rootUrl();
    const completeOnce = completeOnceFactory(event, t);

    INF(t, `FP email validation started - [${VERSION}]`, { platform: platform(), root });

    try {
      await pingServer(t, root);

      const item = Office.context.mailbox.item;
      const isAppt = item.itemType === "appointment";

      // subject/from/to/cc/bcc/location
      const subject = await withTimeout(FIELD_TIMEOUT_MS, done => item.subject.getAsync(r => done(r)), null).then(r => getVal(r, ""));
      const from = isAppt
        ? await withTimeout(FIELD_TIMEOUT_MS, done => item.organizer.getAsync(r => done(r)), null).then(r => getVal(r, ""))
        : await withTimeout(FIELD_TIMEOUT_MS, done => item.from.getAsync(r => done(r)), null).then(r => getVal(r, ""));

      const to = isAppt
        ? await withTimeout(FIELD_TIMEOUT_MS, done => item.requiredAttendees.getAsync(r => done(r)), null).then(r => getVal(r, []))
        : await withTimeout(FIELD_TIMEOUT_MS, done => item.to.getAsync(r => done(r)), null).then(r => getVal(r, []));

      const cc = isAppt
        ? await withTimeout(FIELD_TIMEOUT_MS, done => item.optionalAttendees.getAsync(r => done(r)), null).then(r => getVal(r, []))
        : await withTimeout(FIELD_TIMEOUT_MS, done => item.cc.getAsync(r => done(r)), null).then(r => getVal(r, []));

      const bcc = isAppt
        ? []
        : await withTimeout(FIELD_TIMEOUT_MS, done => item.bcc.getAsync(r => done(r)), null).then(r => getVal(r, []));

      const location = isAppt
        ? await withTimeout(FIELD_TIMEOUT_MS, done => item.location.getAsync(r => done(r)), null).then(r => getVal(r, ""))
        : "";

      // body
      const html = await withTimeout(BODY_TIMEOUT_MS, done => item.body.getAsync(Office.CoercionType.Html, {}, r => done(r)), null)
        .then(r => getVal(r, ""));

      DBG(t, "=== Raw HTML Body ===");
      DBG(t, html);

      const body = normalizeHtml(html);

      DBG(t, "=== Normalized Text ===");
      DBG(t, body);

      // attachments list
      const attsRes = await withTimeout(ATTS_LIST_TIMEOUT_MS, done => item.getAttachmentsAsync(r => done(r)), null);
      const atts = (attsRes && attsRes.status === Office.AsyncResultStatus.Succeeded && Array.isArray(attsRes.value)) ? attsRes.value : [];

      // attachment content (FIX: filtrujemy null; NIE btoa jeśli już base64)
      const attachments = (atts.length === 0) ? [] : (await Promise.all(
        atts.map(a => new Promise((resolve) => {
          let done = false;
          const to = setTimeout(() => { if (!done) { done = true; resolve(null); } }, ATT_CONTENT_TIMEOUT_MS);

          try {
            item.getAttachmentContentAsync(a.id, (r) => {
              if (done) return;
              done = true;
              clearTimeout(to);

              if (!r || r.status !== Office.AsyncResultStatus.Succeeded) return resolve(null);

              const fmt = r.value && r.value.format;
              const content = r.value && r.value.content;

              // najbezpieczniej: obsługujemy base64; resztę pomijamy (żeby nie psuć requestu)
              if (fmt === "base64" && typeof content === "string") {
                return resolve({ file_name: a.name, data: content, content_type: a.contentType });
              }

              // Forcepoint robi btoa dla nie-base64 — ale to bywa ryzykowne; robimy tylko dla "text"
              if (fmt === "text" && typeof content === "string") {
                try {
                  const b64 = btoa(content);
                  return resolve({ file_name: a.name, data: b64, content_type: a.contentType });
                } catch {
                  return resolve(null);
                }
              }

              // np. fmt === "url" (cloud attachment) -> pomijamy
              return resolve(null);
            });
          } catch {
            clearTimeout(to);
            resolve(null);
          }
        }))
      )).filter(Boolean);

      DBG(t, "Payload (truncated)", {
        subjectLen: (subject || "").length,
        bodyLen: (body || "").length,
        attCount: attachments.length
      });

      const payload = { subject, body, from, to, cc, bcc, location, attachments };

      await postToAgent(t, root, payload, completeOnce);

    } catch (e) {
      ERR(t, "handleError", { name: e?.name, message: e?.message || String(e) });
      completeOnce(true, "classify_error_fail_open");
    }
  }

  function onMessageSendHandler(event) {
    Office.onReady().then(() => onSend(event));
  }

  // event handler binding
  try { Office.actions.associate("onMessageSendHandler", onMessageSendHandler); } catch {}
  try { window.onMessageSendHandler = onMessageSendHandler; } catch {}

})();
