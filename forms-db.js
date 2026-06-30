(function(){
  const DB_NAME = "integra6a_forms_db";
  const DB_VERSION = 1;
  const STORE = "submissions";
  const QUEUE = "sync_queue";
  const CONFIG = window.INTEGRA_FORMS_DB_CONFIG || {};

  function openDb(){
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          const store = db.createObjectStore(STORE, { keyPath:"id" });
          store.createIndex("form", "form", { unique:false });
          store.createIndex("createdAt", "createdAt", { unique:false });
        }
        if (!db.objectStoreNames.contains(QUEUE)) {
          db.createObjectStore(QUEUE, { keyPath:"id" });
        }
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function requestToPromise(req){
    return new Promise((resolve, reject) => {
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function transactionToPromise(tx){
    return new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
      tx.onabort = () => reject(tx.error);
    });
  }

  async function put(storeName, item){
    const db = await openDb();
    try {
      const tx = db.transaction(storeName, "readwrite");
      tx.objectStore(storeName).put(item);
      await transactionToPromise(tx);
    } finally {
      db.close();
    }
    return item;
  }

  async function remove(storeName, id){
    const db = await openDb();
    try {
      const tx = db.transaction(storeName, "readwrite");
      tx.objectStore(storeName).delete(id);
      await transactionToPromise(tx);
    } finally {
      db.close();
    }
  }

  async function getAll(storeName){
    const db = await openDb();
    try {
      const tx = db.transaction(storeName, "readonly");
      return await requestToPromise(tx.objectStore(storeName).getAll());
    } finally {
      db.close();
    }
  }

  function makeId(form){
    const safeForm = String(form || "form").replace(/[^a-z0-9_-]+/gi, "-").toLowerCase();
    return `${safeForm}-${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
  }

  function asCsv(rows){
    const headers = ["id","form","createdAt","syncedAt","payload"];
    const escape = value => `"${String(value ?? "").replace(/"/g, '""')}"`;
    return [
      headers.join(","),
      ...rows.map(row => headers.map(header => {
        const value = header === "payload" ? JSON.stringify(row.payload || {}) : row[header];
        return escape(value);
      }).join(","))
    ].join("\r\n");
  }

  function download(filename, content, type){
    const blob = new Blob([content], { type });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(link.href), 500);
  }

  async function sendRemote(row){
    if (!CONFIG.endpoint) return false;
    await fetch(CONFIG.endpoint, {
      method:"POST",
      mode:"no-cors",
      body:JSON.stringify({ ...row, token:CONFIG.token || "" })
    });
    return true;
  }

  function remoteList(form){
    if (!CONFIG.endpoint) return Promise.resolve([]);
    return new Promise((resolve, reject) => {
      const callback = `integraFormsDb_${Date.now()}_${Math.random().toString(36).slice(2)}`;
      const script = document.createElement("script");
      const params = new URLSearchParams({
        action:"list",
        form:form || "",
        callback
      });
      window[callback] = data => {
        resolve(Array.isArray(data && data.records) ? data.records : []);
        script.remove();
        delete window[callback];
      };
      script.onerror = () => {
        reject(new Error("Nao foi possivel ler o banco online."));
        script.remove();
        delete window[callback];
      };
      script.src = `${CONFIG.endpoint}?${params.toString()}`;
      document.head.appendChild(script);
    });
  }

  async function sync(){
    const queue = await getAll(QUEUE);
    let synced = 0;
    for (const row of queue) {
      try {
        const ok = await sendRemote(row);
        if (!ok) break;
        row.syncedAt = new Date().toISOString();
        await put(STORE, row);
        await remove(QUEUE, row.id);
        synced += 1;
      } catch (err) {
        console.warn("Sincronizacao pendente:", err);
        break;
      }
    }
    return synced;
  }

  async function save(form, payload){
    const row = {
      id: makeId(form),
      form,
      payload,
      createdAt: new Date().toISOString(),
      syncedAt: ""
    };
    await put(STORE, row);
    await put(QUEUE, row);
    await sync();
    return row;
  }

  async function list(form){
    const [localRows, remoteRows] = await Promise.all([
      getAll(STORE),
      remoteList(form).catch(err => {
        console.warn("Banco online indisponivel; usando dados locais.", err);
        return [];
      })
    ]);
    const mapa = {};
    [...localRows, ...remoteRows].forEach(row => {
      if (row && row.id) mapa[row.id] = row;
    });
    return Object.values(mapa)
      .filter(row => !form || row.form === form)
      .sort((a,b) => String(b.createdAt).localeCompare(String(a.createdAt)));
  }

  async function listPayloads(form){
    const rows = await list(form);
    return rows.map(row => row.payload || row);
  }

  async function exportJson(form){
    const rows = await list(form);
    const stamp = new Date().toISOString().slice(0,10);
    download(`formularios_${form || "todos"}_${stamp}.json`, JSON.stringify(rows, null, 2), "application/json;charset=utf-8");
  }

  async function exportCsv(form){
    const rows = await list(form);
    const stamp = new Date().toISOString().slice(0,10);
    download(`formularios_${form || "todos"}_${stamp}.csv`, asCsv(rows), "text/csv;charset=utf-8");
  }

  function renderStatus(targetId){
    const target = document.getElementById(targetId);
    if (!target) return;
    Promise.all([getAll(STORE), getAll(QUEUE)]).then(([rows, queue]) => {
      target.textContent = `${rows.length} registro(s) no banco local | ${queue.length} pendente(s) de envio`;
    }).catch(() => {
      target.textContent = "Banco local indisponivel neste navegador";
    });
  }

  window.IntegraFormsDB = {
    save,
    list,
    listPayloads,
    sync,
    exportJson,
    exportCsv,
    renderStatus
  };

  window.addEventListener("online", sync);
})();
