// ===== E-Cíclica • armazenamento no FIREBASE (uso em equipe) =====
// Entrada por matrícula (login anônimo do Firebase + cadastro nome/matrícula),
// Firestore para documento do mês e preenchimentos. Funciona offline.
// Upload/troca de documento exige a senha de upload (conferida pelas regras do servidor).

const V = '10.12.2';
const CDN = `https://www.gstatic.com/firebasejs/${V}/`;
const CHUNK = 700000; // caracteres base64 por documento (limite do Firestore é ~1 MB)
const SESSION = 'eciclica-mat';
const UNLOCK = 'eciclica-unlock';

function b64(bytes) {
  let s = ''; const step = 0x8000;
  for (let i = 0; i < bytes.length; i += step) s += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
  return btoa(s);
}
function unb64(str) { const bin = atob(str); const out = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i); return out; }

function niceError(e) {
  const c = e?.code || '';
  const map = {
    'auth/operation-not-allowed': 'Ative o login "Anônimo" no Firebase (Authentication → Método de login).',
    'auth/admin-restricted-operation': 'Ative o login "Anônimo" no Firebase (Authentication → Método de login).',
    'auth/network-request-failed': 'Sem conexão com a internet.',
    'auth/too-many-requests': 'Muitas tentativas. Aguarde alguns minutos.',
    'permission-denied': 'Sem permissão. Confira as regras do Firestore (README).',
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
  try { fs = F.initializeFirestore(app, { localCache: F.persistentLocalCache({ tabManager: F.persistentMultipleTabManager() }) }); }
  catch { fs = F.getFirestore(app); }
  const d = (...p) => F.doc(fs, ...p);
  const col = (...p) => F.collection(fs, ...p);

  async function ensureAuth() {
    await auth.authStateReady?.();
    if (!auth.currentUser) await AU.signInAnonymously(auth);
    return auth.currentUser;
  }
  const getMat = () => { try { return localStorage.getItem(SESSION); } catch { return null; } };
  const unlocked = () => { try { return sessionStorage.getItem(UNLOCK) === '1'; } catch { return false; } };
  const needUnlock = () => { if (!unlocked()) throw new Error('Digite a senha de upload para continuar.'); };

  let me = null; let authCb = () => {};
  async function refreshMe() {
    const mat = getMat();
    if (!mat) { me = null; return authCb(null); }
    try {
      await ensureAuth();
      const s = await F.getDoc(d('users', mat));
      me = s.exists() ? s.data() : null;
      if (!me) { try { localStorage.removeItem(SESSION); } catch {} }
      authCb(me);
    } catch (e) { me = null; authCb(null, niceError(e).message); }
  }

  return {
    mode: 'firebase',
    onAuth(cb) { authCb = cb; refreshMe(); },
    lookup: wrap(async (mat) => { await ensureAuth(); const s = await F.getDoc(d('users', mat)); return s.exists() ? s.data() : null; }),
    login: wrap(async (mat) => {
      await ensureAuth();
      if (!(await F.getDoc(d('users', mat))).exists()) throw new Error('Matrícula não cadastrada.');
      try { localStorage.setItem(SESSION, mat); } catch {}
      await refreshMe();
    }),
    register: wrap(async ({ nome, matricula, equipe }) => {
      await ensureAuth();
      const ref = d('users', matricula);
      if ((await F.getDoc(ref)).exists()) throw new Error('Essa matrícula já está cadastrada.');
      await F.setDoc(ref, { nome, matricula, equipe: equipe || '', createdAt: Date.now() });
      try { localStorage.setItem(SESSION, matricula); } catch {}
      await refreshMe();
    }),
    updateMe: wrap(async (patch) => { if (!me) return; await F.updateDoc(d('users', me.matricula), patch); await refreshMe(); }),
    logout: wrap(async () => { try { localStorage.removeItem(SESSION); sessionStorage.removeItem(UNLOCK); } catch {} me = null; authCb(null); }),
    listUsers: wrap(async () => (await F.getDocs(col('users'))).docs.map((s) => s.data())),

    isUnlocked: unlocked,
    // a senha é conferida pelas regras do Firestore (o servidor guarda só o hash)
    unlock: wrap(async (senha) => {
      const u = await ensureAuth();
      try { await F.setDoc(d('uploadGrants', u.uid), { chave: senha, at: F.serverTimestamp() }); }
      catch (e) { if (e.code === 'permission-denied') throw new Error('Senha recusada pelo servidor. Se a senha estiver certa, publique de novo as regras (firestore.rules) no Firebase.'); throw e; }
      try { sessionStorage.setItem(UNLOCK, '1'); } catch {}
    }),
    lock() { try { sessionStorage.removeItem(UNLOCK); } catch {} },

    getConfig: wrap(async () => { await ensureAuth(); return (await F.getDoc(d('config', 'app'))).data() || { activeCycle: null, equipes: cfg.equipes || [] }; }),
    setConfig: wrap(async (patch) => { needUnlock(); await F.setDoc(d('config', 'app'), patch, { merge: true }); }),
    watchConfig(cb) {
      let un = () => {}; let dead = false;
      ensureAuth().then(() => { if (!dead) un = F.onSnapshot(d('config', 'app'), (s) => cb(s.data() || { activeCycle: null, equipes: cfg.equipes || [] }), (e) => console.warn(e)); });
      return () => { dead = true; un(); };
    },

    listCycles: wrap(async () => (await F.getDocs(col('cycles'))).docs.map((s) => ({ id: s.id, ...s.data() }))
      .sort((a, b) => (b.createdAt || 0) - (a.createdAt || 0))),
    createCycle: wrap(async (meta, rows, bytes) => {
      needUnlock();
      const id = (meta.ym || 'ciclo') + '-' + Date.now().toString(36);
      const groups = {};
      for (const r of rows) (groups[r.g] ||= []).push(r);
      const str = b64(bytes); const parts = [];
      for (let i = 0; i < str.length; i += CHUNK) parts.push(str.slice(i, i + CHUNK));
      let b = F.writeBatch(fs); let size = 0;
      const flush = async () => { if (size) { await b.commit(); b = F.writeBatch(fs); size = 0; } };
      for (const [g, list] of Object.entries(groups)) {
        b.set(d('cycles', id, 'rows', g), { rows: list }); size += JSON.stringify(list).length;
        if (size > 3e6) await flush();
      }
      await flush();
      for (let i = 0; i < parts.length; i++) { b.set(d('cycles', id, 'file', String(i)), { d: parts[i] }); size += parts[i].length; await flush(); }
      await F.setDoc(d('cycles', id), { ...meta, parts: parts.length, createdAt: Date.now(), createdBy: me?.nome || '' });
      return id;
    }),
    deleteCycle: wrap(async (id) => {
      needUnlock();
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
      const p = b.commit(); // offline: fica na fila e sobe depois
      await Promise.race([p, new Promise((r) => setTimeout(r, 1500))]);
      p.catch((e) => console.warn('falha ao sincronizar', e));
    }),
  };
}
