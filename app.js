/* global Office, OfficeRuntime */

(() => {
  "use strict";

  const VERSION = "min-sharedlog-1.0";
  const ROOT_WIN = "https://localhost:55299/";
  const ROOT_MAC = "https://localhost:55296/";
  const PING = "FirefoxExt/_1";
  const POST = "OutlookAddin";

  // Forcepointowe timeouty + watchdog anty-hang
  const PING_TIMEOUT_MS = 30000;
  const POST_TIMEOUT_MS = 35000;
  const WATCHDOG_MS = 60000; // po minucie fail-open, żeby nie wisiało

  const LOG_KEY = "smimeDlp.logs.v2";   // nowy key (żeby nie mieszać ze starymi)
  const DEBUG_KEY = "smimeDlp.debug.v2";

  const now = () => new Date().toISOString();
  const tx = () => `TX-${Date.now()}-${Math.floor(Math.random() * 1e6)}`;

  async function sGet(k) {
    try { if (OfficeRuntime?.storage) return await OfficeRuntime.storage.getItem(k); } catch {}
    try { return localStorage.getItem(k); } catch {}
    return null;
  }
  async function sSet(k, v) {
    try { if (OfficeRuntime?.storage) return await OfficeRuntime.storage.setItem(k, v); } catch {}
    try { localStorage.setItem(k, v); } catch {}
  }

  async function isDebug() { return (await sGet(DEBUG_KEY)) === "1"; }

  async function log(level, id, msg, meta) {
    const e = { ts: now(), level, tx: id, msg: String(msg), meta: meta ?? null };

    try {
      const line = `${e.ts} [${e.level}] [${e.tx}] ${e.msg}`;
      if (level === "ERR") console.error(line, e.meta || "");
      else if (level === "DBG") console.debug(line, e.meta || "");
      else console.log(line, e.meta || "");
    } catch {}

    try {
      const raw = await sGet(LOG_KEY);
      const arr = raw ? JSON.parse(raw) : [];
      const list = Array.isArray(arr) ? arr : [];
      list.push(e);
      while (list.length > 500) list.shift();
      await sSet(LOG_KEY, JSON.stringify(list));
    } catch {}

    if (await isDebug()) {
      try {
        Office.context.mailbox.item.notificationMessages.replaceAsync("smimeDlp", {
          type: "progressIndicator",
          message: e.msg.substring(0, 250),
        });
      } catch {}
    }
  }

  const INF = (t, m, x) => log("INF", t, m, x);
  const DBG = (t, m, x) => log("DBG", t, m, x);
  const ERR = (t, m, x) => log("ERR", t, m, x);

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

  function normalizeHtml(html) {
    // prosto jak Forcepoint: strip tagów, bez dekodowania &nbsp;
    return String(html || "").replace(/<[^>]+>/g, "");
  }

  async function ping(t, root) {
    await INF(t, "Checking the server");
    const r = await fetchTimeout(root + PING, { method: "GET", mode: "cors" }, PING_TIMEOUT_MS);
    if (!r.ok) throw new Error(`PING_HTTP_${r.status}`);
    await INF(t, "Server is UP");
  }

  function handleResponse(t, data, event) {
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
      event.completed({ allowEvent: false });
    } else {
      INF(t, "DLP allow.");
      event.completed({ allowEvent: true });
    }
  }

  async function post(t, root, payload, event) {
    INF(t, "Sending event to classifier");
    const resp = await fetchTimeout(root + POST, {
      method: "POST",
      mode: "cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    }, POST_TIMEOUT_MS);

    if (!resp.ok) throw new Error(`POST_HTTP_${resp.status}`);
    const json = await resp.json();
    handleResponse(t, json, event);
  }

  async function collectPayload(t) {
    const item = Office.context.mailbox.item;

    // minimalnie: subject + body + attachments (reszta możesz dopiąć później)
    const subject = await new Promise((resolve) => item.subject.getAsync(r => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : "")));

    const html = await new Promise((resolve) =>
      item.body.getAsync(Office.CoercionType.Html, {}, r => resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ""))
    );

    const body = normalizeHtml(html);

    // attachments list
    const attsRes = await new Promise((resolve) => item.getAttachmentsAsync(r => resolve(r)));
    const list = (attsRes.status === Office.AsyncResultStatus.Succeeded && Array.isArray(attsRes.value)) ? attsRes.value : [];

    // attachments content – sekwencyjnie, filtr null
    const attachments = [];
    for (const a of list) {
      const one = await new Promise((resolve) => {
        let done = false;
        const to = setTimeout(() => { if (!done) { done = true; resolve(null); } }, 30000);

        item.getAttachmentContentAsync(a.id, (r) => {
          if (done) return;
          done = true;
          clearTimeout(to);

          if (r.status !== Office.AsyncResultStatus.Succeeded) return resolve(null);
          if (r.value?.format !== "base64" || typeof r.value?.content !== "string") return resolve(null);

          resolve({ file_name: a.name, data: r.value.content, content_type: a.contentType });
        });
      });
      if (one) attachments.push(one);
    }

    DBG(t, "Payload (truncated)", { subjectLen: subject.length, bodyLen: body.length, attCount: attachments.length });
    return { subject, body, attachments, from: "", to: [], cc: [], bcc: [], location: "" };
  }

  async function onSend(event) {
    const t = tx();
    const root = rootUrl();

    INF(t, `START [${VERSION}]`, { platform: platform(), root });

    // watchdog anty-hang
    const wd = setTimeout(() => {
      ERR(t, "WATCHDOG fired -> fail-open");
      try { event.completed({ allowEvent: true }); } catch {}
    }, WATCHDOG_MS);

    try {
      await ping(t, root);
      const payload = await collectPayload(t);
      await post(t, root, payload, event);
    } catch (e) {
      ERR(t, "handleError", { name: e?.name, message: e?.message || String(e) });
      try { event.completed({ allowEvent: true }); } catch {}
    } finally {
      clearTimeout(wd);
    }
  }

  // DLA manifestu klasycznego (Events/ItemSend) MUSI być global:
  window.onMessageSendHandler = function (event) {
    Office.onReady().then(() => onSend(event));
  };

})();
