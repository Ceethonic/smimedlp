/* global Office, OffiiiiceRuntime */

(() => {
  "use strict";
  const LOG_KEY = "smimeDlp.logs.v2";
  const DEBUG_KEY = "smimeDlp.debug.v2";

  async function sGet(k) {
    try { if (OfficeRuntime?.storage) return await OfficeRuntime.storage.getItem(k); } catch {}
    try { return localStorage.getItem(k); } catch {}
    return null;
  }
  async function sSet(k, v) {
    try { if (OfficeRuntime?.storage) return await OfficeRuntime.storage.setItem(k, v); } catch {}
    try { localStorage.setItem(k, v); } catch {}
  }

  function fmt(e) {
    const meta = e.meta ? " " + JSON.stringify(e.meta) : "";
    return `${e.ts} [${e.level}] [${e.tx}] ${e.msg}${meta}`;
  }

  async function render() {
    let arr = [];
    try {
      const raw = await sGet(LOG_KEY);
      arr = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(arr)) arr = [];
    } catch { arr = []; }

    document.getElementById("out").textContent = arr.length ? arr.map(fmt).join("\n") : "(brak logów)";
  }

  Office.onReady().then(async () => {
    const dbg = document.getElementById("dbg");
    dbg.checked = (await sGet(DEBUG_KEY)) === "1";

    dbg.addEventListener("change", async () => {
      await sSet(DEBUG_KEY, dbg.checked ? "1" : "0");
      await render();
    });

    document.getElementById("refresh").addEventListener("click", render);
    document.getElementById("clear").addEventListener("click", async () => {
      await sSet(LOG_KEY, JSON.stringify([]));
      await render();
    });

    await render();
  });
})();
