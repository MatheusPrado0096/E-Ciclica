// ===== E-Cíclica • armazenamento LOCAL (modo demonstração) =====
// Tudo fica no IndexedDB deste navegador. Serve para testar o sistema
// sem configurar o Firebase. Para uso em equipe, configure o Firebase (config.js).

const DB = 'eciclica-local2';
let dbp;
function db() {
  if (!dbp) dbp = new Promise((res, rej) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  });
  return dbp;
}
async function get(k) { const d = await db(); return new Promise((res, rej) => { const t = d.transaction('kv').objectStore('kv').get(k); t.onsuccess = () => res(t.result); t.onerror = () => rej(t.error); }); }
async function put(k, v) { const d = await db(); return new Promise((res, rej) => { const tx = d.transaction('kv', 'readwrite'); tx.objectStore('kv').put(v, k); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); }); }
async function del(k) { const d = await db(); return new Promise((res, rej) => { const tx = d.transaction('kv', 'readwrite'); tx.objectStore('kv').delete(k); tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error); }); }

const bc = 'BroadcastChannel' in window ? new BroadcastChannel('eciclica') : null;
const subs = new Set();
function emit(topic) { subs.forEach((s) => s.topic === topic && s.fn()); bc?.postMessage(topic); }
bc && (bc.onmessage = (e) => subs.forEach((s) => s.topic === e.data && s.fn()));
function on(topic, fn) { const s = { topic, fn }; subs.add(s); return () => subs.delete(s); }

export async function sha256b64(text) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return btoa(String.fromCharCode(...new Uint8Array(buf)));
}
// hash da senha de upload (a senha em si não fica no código)
export const UPLOAD_HASH = 'mvFbM25qlhmShTffMLLmojdlafz51+dz7M7eZWBlKaA=';
const SESSION = 'eciclica-mat';
const UNLOCK = 'eciclica-unlock';

export function createLocalStore(cfg) {
  let authCb = () => {}; let me = null;
  const users = async () => (await get('users')) || {};
  async function refreshMe() {
    let mat = null; try { mat = localStorage.getItem(SESSION); } catch {}
    const u = mat && (await users())[mat];
    me = u || null; authCb(me);
  }
  const unlocked = () => { try { return sessionStorage.getItem(UNLOCK) === '1'; } catch { return false; } };
  const needUnlock = () => { if (!unlocked()) throw new Error('Digite a senha de upload para continuar.'); };

  return {
    mode: 'local',
    onAuth(cb) { authCb = cb; refreshMe(); },
    async lookup(mat) { return (await users())[mat] || null; },
    async login(mat) {
      const u = (await users())[mat]; if (!u) throw new Error('Matrícula não cadastrada.');
      try { localStorage.setItem(SESSION, mat); } catch {}
      await refreshMe();
    },
    async register({ nome, matricula, equipe }) {
      const all = await users();
      if (all[matricula]) throw new Error('Essa matrícula já está cadastrada.');
      all[matricula] = { nome, matricula, equipe: equipe || '', createdAt: Date.now() };
      await put('users', all); emit('users');
      await this.login(matricula);
    },
    async updateMe(patch) {
      const all = await users(); if (!me) return;
      Object.assign(all[me.matricula], patch); await put('users', all); emit('users'); await refreshMe();
    },
    async logout() { try { localStorage.removeItem(SESSION); sessionStorage.removeItem(UNLOCK); } catch {} me = null; authCb(null); },
    async listUsers() { return Object.values(await users()); },

    isUnlocked: unlocked,
    async unlock(senha) {
      if ((await sha256b64(senha)) !== UPLOAD_HASH) throw new Error('Senha incorreta (modo local).');
      try { sessionStorage.setItem(UNLOCK, '1'); } catch {}
    },
    lock() { try { sessionStorage.removeItem(UNLOCK); } catch {} },

    async getConfig() { return (await get('config')) || { activeCycle: null, equipes: cfg.equipes || [] }; },
    async setConfig(patch) { needUnlock(); await put('config', { ...(await this.getConfig()), ...patch }); emit('config'); },
    watchConfig(cb) { const f = async () => cb(await this.getConfig()); f(); return on('config', f); },

    async listCycles() { return ((await get('cycles')) || []).sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0)); },
    async createCycle(meta, rows, bytes) {
      needUnlock();
      const id = (meta.ym || 'ciclo') + '-' + Date.now().toString(36);
      await put('rows:' + id, rows); await put('file:' + id, bytes); await put('fills:' + id, {});
      const c = { ...meta, id, createdAt: Date.now(), createdBy: me?.nome || '' };
      await put('cycles', [...(await this.listCycles()), c]); emit('cycles');
      return id;
    },
    async deleteCycle(id) {
      needUnlock();
      await del('rows:' + id); await del('file:' + id); await del('fills:' + id);
      await put('cycles', (await this.listCycles()).filter((c) => c.id !== id)); emit('cycles');
    },
    async loadRows(id) { return (await get('rows:' + id)) || []; },
    async getFile(id) { const b = await get('file:' + id); return b ? new Uint8Array(b) : null; },
    watchFills(id, cb) { const f = async () => cb((await get('fills:' + id)) || {}); f(); return on('fills:' + id, f); },
    async saveFills(id, list) {
      const all = (await get('fills:' + id)) || {};
      for (const { k, fill } of list) { if (fill) all[k] = fill; else delete all[k]; }
      await put('fills:' + id, all); emit('fills:' + id);
    },
  };
}
