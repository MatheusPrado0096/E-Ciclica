// ===== E-Cíclica • armazenamento LOCAL (modo demonstração) =====
// Tudo fica no IndexedDB deste navegador. Serve para testar o sistema
// sem configurar o Firebase. Para uso em equipe, configure o Firebase (config.js).

const DB = 'eciclica-local';
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

async function hash(pw, salt) {
  const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(salt + '::' + pw));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, '0')).join('');
}
const rid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
const SESSION = 'eciclica-session';
const pub = (u) => u && ({ uid: u.uid, name: u.name, username: u.username, role: u.role, equipe: u.equipe || '', active: u.active !== false });

export function createLocalStore(cfg) {
  let authCb = () => {}; let me = null;
  const users = async () => (await get('users')) || [];

  async function refreshMe() {
    let uid = null; try { uid = localStorage.getItem(SESSION); } catch {}
    const u = uid && (await users()).find((x) => x.uid === uid && x.active !== false && !x.removed);
    me = u ? pub(u) : null; authCb(me);
  }

  return {
    mode: 'local',
    async init() {},
    onAuth(cb) { authCb = cb; refreshMe(); },
    async needsBootstrap() { return !(await users()).some((u) => u.role === 'admin' && !u.removed); },
    async bootstrap({ name, username, password }) {
      if (!(await this.needsBootstrap())) throw new Error('O administrador já foi criado.');
      const salt = rid();
      const u = { uid: rid(), name, username, role: 'admin', equipe: '', active: true, salt, hash: await hash(password, salt) };
      await put('users', [...(await users()), u]);
      await put('config', { activeCycle: null, equipes: cfg.equipes || [] });
      try { localStorage.setItem(SESSION, u.uid); } catch {}
      await refreshMe();
    },
    async login(username, password) {
      const u = (await users()).find((x) => x.username === username && !x.removed);
      if (!u || (await hash(password, u.salt)) !== u.hash) throw new Error('Usuário ou senha incorretos.');
      if (u.active === false) throw new Error('Este acesso está desativado. Fale com o administrador.');
      try { localStorage.setItem(SESSION, u.uid); } catch {}
      await refreshMe();
    },
    async logout() { try { localStorage.removeItem(SESSION); } catch {} me = null; authCb(null); },
    async changePassword(oldPw, newPw) {
      const list = await users(); const u = list.find((x) => x.uid === me?.uid);
      if (!u || (await hash(oldPw, u.salt)) !== u.hash) throw new Error('Senha atual incorreta.');
      u.salt = rid(); u.hash = await hash(newPw, u.salt); await put('users', list);
    },
    async listUsers() { return (await users()).filter((u) => !u.removed).map(pub); },
    async createUser({ name, username, password, role, equipe }) {
      const list = await users();
      if (list.some((u) => u.username === username && !u.removed)) throw new Error('Já existe um usuário com esse login.');
      const salt = rid();
      list.push({ uid: rid(), name, username, role, equipe, active: true, salt, hash: await hash(password, salt) });
      await put('users', list); emit('users');
    },
    async updateUser(uid, patch) {
      const list = await users(); const u = list.find((x) => x.uid === uid); if (!u) return;
      Object.assign(u, patch); await put('users', list); emit('users');
      if (uid === me?.uid) refreshMe();
    },
    async resetPassword(uid, pw) {
      const list = await users(); const u = list.find((x) => x.uid === uid); if (!u) return;
      u.salt = rid(); u.hash = await hash(pw, u.salt); await put('users', list);
    },
    async removeUser(uid) { await this.updateUser(uid, { removed: true, active: false }); },

    async getConfig() { return (await get('config')) || { activeCycle: null, equipes: cfg.equipes || [] }; },
    async setConfig(patch) { await put('config', { ...(await this.getConfig()), ...patch }); emit('config'); },
    watchConfig(cb) { const f = async () => cb(await this.getConfig()); f(); return on('config', f); },

    async listCycles() { return ((await get('cycles')) || []).sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0)); },
    async createCycle(meta, rows, bytes) {
      const id = (meta.ym || 'ciclo') + '-' + Date.now().toString(36);
      await put('rows:' + id, rows); await put('file:' + id, bytes); await put('fills:' + id, {});
      const c = { ...meta, id, createdAt: Date.now(), createdBy: me?.name || '' };
      await put('cycles', [...(await this.listCycles()), c]); emit('cycles');
      return id;
    },
    async deleteCycle(id) {
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
