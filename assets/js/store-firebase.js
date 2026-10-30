// ===== E-Cíclica • armazenamento no FIREBASE (uso em equipe) =====
// Auth (usuário + senha) e Firestore (documento do mês, preenchimentos e usuários).
// Funciona offline: os preenchimentos ficam guardados no aparelho e sobem quando volta o sinal.

const V = '10.12.2';
const CDN = `https://www.gstatic.com/firebasejs/${V}/`;
const CHUNK = 700000; // caracteres base64 por documento (limite do Firestore é ~1 MB)

function b64(bytes) {
  let s = ''; const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) s += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
  return btoa(s);
}
function unb64(str) { const bin = atob(str); const out = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i); return out; }
const tag = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

function niceError(e) {
  const c = e?.code || '';
  const map = {
    'auth/invalid-credential': 'Usuário ou senha incorretos.',
    'auth/wrong-password': 'Usuário ou senha incorretos.',
    'auth/user-not-found': 'Usuário ou senha incorretos.',
    'auth/invalid-email': 'Usuário inválido.',
    'auth/too-many-requests': 'Muitas tentativas. Aguarde alguns minutos e tente de novo.',
    'auth/weak-password': 'A senha precisa ter pelo menos 6 caracteres.',
    'auth/email-already-in-use': 'Esse login já existe.',
    'auth/network-request-failed': 'Sem conexão com a internet.',
    'auth/operation-not-allowed': 'Ative o login por "E-mail/senha" no Firebase (Authentication → Sign-in method).',
    'permission-denied': 'Sem permissão. Confira as regras do Firestore (README) ou se o seu acesso está ativo.',
    'unavailable': 'Sem conexão com o servidor. Tente de novo.',
  };
  const err = new Error(map[c] || e?.message || String(e)); err.code = c; return err;
}
const wrap = (fn) => async (...a) => { try { return await fn(...a); } catch (e) { throw niceError(e); } };

export async function createFirebaseStore(cfg) {
  const [A, AU, F] = await Promise.all([
    import(CDN + 'firebase-app.js'), import(CDN + 'firebase-auth.js'), import(CDN + 'firebase-firestore.js'),
  ]);
  const app = A.initializeApp(cfg.firebase);
  const auth = AU.getAuth(app);
  let fs;
  try {
    fs = F.initializeFirestore(app, { localCache: F.persistentLocalCache({ tabManager: F.persistentMultipleTabManager() }) });
  } catch { fs = F.getFirestore(app); }
  const d = (...p) => F.doc(fs, ...p);
  const col = (...p) => F.collection(fs, ...p);
  const dom = cfg.dominioLogin || 'eciclica.app';
  const emailFor = (u) => `${u}-${tag()}@${dom}`;

  // app secundário: cria contas sem deslogar o administrador
  let sec;
  function secondary() {
    if (!sec) {
      const a2 = A.initializeApp(cfg.firebase, 'secundario');
      sec = AU.initializeAuth(a2, { persistence: AU.inMemoryPersistence });
    }
    return sec;
  }
  async function newAccount(username, password) {
    const s = secondary();
    const cred = await AU.createUserWithEmailAndPassword(s, emailFor(username), password);
    const uid = cred.user.uid; const email = cred.user.email;
    await AU.signOut(s);
    return { uid, email };
  }

  let me = null; let authCb = () => {}; let busy = false;
  async function handle(user) {
    if (!user) { me = null; return authCb(null); }
    try {
      const snap = await F.getDoc(d('users', user.uid));
      const p = snap.exists() ? snap.data() : null;
      if (!p || p.active === false || p.removed) {
        await AU.signOut(auth); me = null;
        return authCb(null, p ? 'Este acesso está desativado. Fale com o administrador.' : 'Usuário sem cadastro no sistema.');
      }
      me = { uid: user.uid, ...p }; authCb(me);
    } catch (e) { authCb(null, niceError(e).message); }
  }
  const store = {
    mode: 'firebase',
    async init() {},
    onAuth(cb) {
      authCb = cb;
      AU.onAuthStateChanged(auth, (user) => { if (!busy) handle(user); });
    },
    needsBootstrap: wrap(async () => !(await F.getDoc(d('config', 'bootstrap'))).exists()),
    bootstrap: wrap(async ({ name, username, password }) => {
      busy = true;
      try {
        const cred = await AU.createUserWithEmailAndPassword(auth, emailFor(username), password);
        const uid = cred.user.uid;
        const b = F.writeBatch(fs);
        b.set(d('users', uid), { name, username, role: 'admin', equipe: '', active: true, createdAt: F.serverTimestamp() });
        b.set(d('logins', username), { email: cred.user.email, uid });
        b.set(d('config', 'bootstrap'), { adminUid: uid, at: F.serverTimestamp() });
        b.set(d('config', 'app'), { activeCycle: null, equipes: cfg.equipes || [] });
        await b.commit();
      } finally { busy = false; }
      await handle(auth.currentUser);
    }),
    login: wrap(async (username, password) => {
      const l = await F.getDoc(d('logins', username));
      if (!l.exists()) throw new Error('Usuário ou senha incorretos.');
      await AU.signInWithEmailAndPassword(auth, l.data().email, password);
    }),
    logout: wrap(async () => AU.signOut(auth)),
    changePassword: wrap(async (oldPw, newPw) => {
      const u = auth.currentUser;
      await AU.reauthenticateWithCredential(u, AU.EmailAuthProvider.credential(u.email, oldPw));
      await AU.updatePassword(u, newPw);
    }),

    listUsers: wrap(async () => (await F.getDocs(col('users')))
      .docs.map((s) => ({ uid: s.id, ...s.data() })).filter((u) => !u.removed && !u.replacedBy)),
    createUser: wrap(async ({ name, username, password, role, equipe }) => {
      if ((await F.getDoc(d('logins', username))).exists()) throw new Error('Já existe um usuário com esse login.');
      const { uid, email } = await newAccount(username, password);
      const b = F.writeBatch(fs);
      b.set(d('users', uid), { name, username, role, equipe: equipe || '', active: true, createdAt: F.serverTimestamp() });
      b.set(d('logins', username), { email, uid });
      await b.commit();
    }),
    updateUser: wrap(async (uid, patch) => F.updateDoc(d('users', uid), patch)),
    // o Firebase não deixa um usuário trocar a senha de outro pelo navegador;
    // então criamos uma conta nova com a nova senha e o mesmo login.
    resetPassword: wrap(async (uid, pw) => {
      const old = (await F.getDoc(d('users', uid))).data();
      const { uid: nu, email } = await newAccount(old.username, pw);
      const b = F.writeBatch(fs);
      b.set(d('users', nu), { ...old, active: true, createdAt: F.serverTimestamp() });
      b.update(d('users', uid), { active: false, replacedBy: nu });
      b.set(d('logins', old.username), { email, uid: nu });
      await b.commit();
    }),
    removeUser: wrap(async (uid) => {
      const u = (await F.getDoc(d('users', uid))).data();
      const b = F.writeBatch(fs);
      b.update(d('users', uid), { active: false, removed: true });
      if (u?.username) b.delete(d('logins', u.username));
      await b.commit();
    }),

    getConfig: wrap(async () => (await F.getDoc(d('config', 'app'))).data() || { activeCycle: null, equipes: cfg.equipes || [] }),
    setConfig: wrap(async (patch) => F.setDoc(d('config', 'app'), patch, { merge: true })),
    watchConfig(cb) { return F.onSnapshot(d('config', 'app'), (s) => cb(s.data() || { activeCycle: null, equipes: cfg.equipes || [] }), (e) => console.warn(e)); },

    listCycles: wrap(async () => (await F.getDocs(col('cycles'))).docs.map((s) => ({ id: s.id, ...s.data() }))
      .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0))),
    createCycle: wrap(async (meta, rows, bytes) => {
      const id = (meta.ym || 'ciclo') + '-' + Date.now().toString(36);
      const groups = {};
      for (const r of rows) (groups[r.g] ||= []).push(r);
      const str = b64(bytes); const parts = [];
      for (let i = 0; i < str.length; i += CHUNK) parts.push(str.slice(i, i + CHUNK));
      // grava em lotes pequenos (limite de 10 MB por envio)
      let b = F.writeBatch(fs); let size = 0;
      const flush = async () => { if (size) { await b.commit(); b = F.writeBatch(fs); size = 0; } };
      for (const [g, list] of Object.entries(groups)) {
        b.set(d('cycles', id, 'rows', g), { rows: list }); size += JSON.stringify(list).length;
        if (size > 3e6) await flush();
      }
      await flush();
      for (let i = 0; i < parts.length; i++) { b.set(d('cycles', id, 'file', String(i)), { d: parts[i] }); size += parts[i].length; await flush(); }
      await F.setDoc(d('cycles', id), { ...meta, parts: parts.length, createdAt: Date.now(), createdBy: me?.name || '' });
      return id;
    }),
    deleteCycle: wrap(async (id) => {
      for (const sub of ['rows', 'file', 'fills']) {
        const s = await F.getDocs(col('cycles', id, sub));
        const b = F.writeBatch(fs); s.docs.forEach((x) => b.delete(x.ref)); await b.commit();
      }
      await F.deleteDoc(d('cycles', id));
    }),
    loadRows: wrap(async (id) => {
      const s = await F.getDocs(col('cycles', id, 'rows'));
      return s.docs.flatMap((x) => x.data().rows || []).sort((a, b) => a.i - b.i);
    }),
    getFile: wrap(async (id) => {
      const meta = (await F.getDoc(d('cycles', id))).data();
      const parts = await Promise.all([...Array(meta.parts || 0)].map((_, i) => F.getDoc(d('cycles', id, 'file', String(i)))));
      return unb64(parts.map((p) => p.data().d).join(''));
    }),
    watchFills(id, cb, onErr) {
      return F.onSnapshot(col('cycles', id, 'fills'), { includeMetadataChanges: true }, (s) => {
        const all = {}; s.docs.forEach((x) => Object.assign(all, x.data()));
        cb(all, { pending: s.metadata.hasPendingWrites, cache: s.metadata.fromCache });
      }, (e) => onErr?.(niceError(e)));
    },
    saveFills: wrap(async (id, list) => {
      const by = {};
      for (const { k, g, fill } of list) (by[g] ||= {})[k] = fill ? fill : F.deleteField();
      const b = F.writeBatch(fs);
      for (const [g, obj] of Object.entries(by)) b.set(d('cycles', id, 'fills', g), obj, { merge: true });
      // não espera o servidor: offline o lote fica na fila e sobe depois
      const p = b.commit();
      await Promise.race([p, new Promise((r) => setTimeout(r, 1500))]);
      p.catch((e) => console.warn('falha ao sincronizar', e));
    }),
  };
  return store;
}
