// ===== E-Cíclica • aplicativo =====
import { parseWorkbook, buildRows, exportWorkbook, norm } from './xlsx.js?v=4';
import { createLocalStore } from './store-local.js?v=4';

const cfg = window.ECICLICA_CONFIG || { firebase: {}, equipes: [] };
const JSZip = window.JSZip;
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const h = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const ls = {
  get(k, d) { try { const v = localStorage.getItem('eciclica-' + k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem('eciclica-' + k, JSON.stringify(v)); } catch {} },
};
const pad = (n) => String(n).padStart(2, '0');
const isoOf = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const todayISO = () => isoOf(new Date());
const nowHM = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const brDate = (iso) => (iso ? iso.split('-').reverse().join('/') : '');
const ago = (t) => {
  if (!t) return ''; const s = (Date.now() - t) / 1000;
  if (s < 60) return 'agora'; if (s < 3600) return `há ${Math.floor(s / 60)} min`; if (s < 86400) return `há ${Math.floor(s / 3600)} h`;
  const d = new Date(t); return `${pad(d.getDate())}/${pad(d.getMonth() + 1)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
};
const initials = (n) => String(n || '?').trim().split(/\s+/).slice(0, 2).map((p) => p[0]).join('').toUpperCase();
const ICON = (id) => `<svg class="ic"><use href="#i-${id}"/></svg>`;
const cleanMat = (v) => String(v || '').trim().replace(/[^0-9A-Za-z_-]/g, '').toUpperCase();
const splitColab = (s) => String(s || '').split(',').map((x) => x.trim()).filter(Boolean);

// ---------- estado ----------
const S = {
  store: null, me: null, config: { equipes: cfg.equipes || [] }, cycle: null, cycles: [], rows: [], byK: new Map(), fills: {},
  users: [], f: { q: '', patio: ls.get('patio', ''), status: 'todas', aba: '' }, selecting: false, sel: new Set(),
  shown: 0, filtered: [], tab: 'ciclica', unsubFills: null, unsubConfig: null, parsed: null, upBytes: null, upName: '',
  d: { per: 'all', from: '', to: '', patio: '', eq: '', col: '', view: 'geral' },
};

// ---------- utilidades de interface ----------
let toastT;
function toast(msg, ms = 2400) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('show'), ms); }
function setSeg(el, v) { $$('[role=radio]', el).forEach((b) => b.setAttribute('aria-checked', String(b.dataset.v === v))); el.dataset.value = v ?? ''; }
function segVal(el) { return el.dataset.value || ''; }
function bindSeg(el, onChange) { el.addEventListener('click', (e) => { const b = e.target.closest('[role=radio]'); if (!b) return; setSeg(el, b.dataset.v); onChange?.(b.dataset.v); }); }
function download(bytes, name, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const a = Object.assign(document.createElement('a'), { href: url, download: name }); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}
function equipes() { return S.config.equipes?.length ? S.config.equipes : cfg.equipes || []; }
function eqChips(el, v) {
  el.innerHTML = equipes().map((e) => `<button type="button" class="chip" role="radio" data-v="${h(e)}" aria-checked="${e === v}" aria-pressed="${e === v}">${h(e)}</button>`).join('');
  el.dataset.value = v || '';
}
function bindEqChips(el, after) {
  el.addEventListener('click', (e) => { const b = e.target.closest('[data-v]'); if (!b) return; eqChips(el, b.dataset.v === el.dataset.value ? '' : b.dataset.v); after?.(); });
}

// diálogo genérico
function dialog({ title, html = '', yes = 'OK', no = 'Cancelar', danger = false, onOpen, validate }) {
  return new Promise((resolve) => {
    const d = $('#dlg'); $('#dlgTitle').textContent = title; $('#dlgBody').innerHTML = html;
    const y = $('#dlgYes'); y.textContent = yes; y.style.background = danger ? 'var(--danger)' : ''; y.style.color = danger ? '#fff' : '';
    $('#dlgNo').textContent = no; $('#dlgNo').hidden = !no;
    d.hidden = false; onOpen?.($('#dlgBody'));
    setTimeout(() => ($('#dlgBody input') || y).focus(), 60);
    const close = (val) => { d.classList.add('closing'); setTimeout(() => { d.hidden = true; d.classList.remove('closing'); }, 200); cleanup(); resolve(val); };
    const onSubmit = async (e) => { e.preventDefault(); const v = validate ? await validate($('#dlgBody')) : true; if (v !== false && v !== undefined) close(v); };
    const onClick = (e) => { if (e.target.closest('[data-dclose]')) close(null); };
    const onKey = (e) => { if (e.key === 'Escape') close(null); };
    function cleanup() { $('#dlgForm').removeEventListener('submit', onSubmit); d.removeEventListener('click', onClick); document.removeEventListener('keydown', onKey); }
    $('#dlgForm').addEventListener('submit', onSubmit); d.addEventListener('click', onClick); document.addEventListener('keydown', onKey);
  });
}

// ---------- tema ----------
function applyTheme(t) { document.documentElement.dataset.theme = t; $('#themeIcon').setAttribute('href', t === 'dark' ? '#i-sun' : '#i-moon'); document.querySelector('meta[name=theme-color]').content = t === 'dark' ? '#031221' : '#00335e'; }
$('#themeBtn').addEventListener('click', () => { const t = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'; applyTheme(t); try { localStorage.setItem('eciclica-theme', t); } catch {} if (S.tab === 'painel') renderDash(); });
applyTheme(document.documentElement.dataset.theme);

// ---------- inicialização ----------
async function boot() {
  try {
    if (cfg.firebase?.apiKey) {
      $('#bootMsg').textContent = 'Conectando…';
      const { createFirebaseStore } = await import('./store-firebase.js?v=4');
      S.store = await createFirebaseStore(cfg);
    } else S.store = createLocalStore(cfg);
  } catch (e) {
    console.error(e);
    $('#bootMsg').innerHTML = `Não foi possível conectar ao Firebase.<br><small>${h(e.message)}</small><br><br><button class="btn primary" onclick="location.reload()">Tentar de novo</button>`;
    $('.boot .spinner').hidden = true; return;
  }
  if (S.store.mode === 'local') {
    $('#modeNote').hidden = false;
    $('#modeNote').innerHTML = '<b>Modo local (demonstração).</b> Os dados ficam só neste navegador. Para a equipe usar junto, configure o Firebase no <code>config.js</code>.';
    $('#localBanner').hidden = !!ls.get('lbHide', false);
    $('#lbX').addEventListener('click', () => { $('#localBanner').hidden = true; ls.set('lbHide', true); });
  }
  S.store.onAuth(async (me, msg) => {
    S.me = me;
    if (me) await enterApp(); else showAuth(msg);
    $('#boot').classList.add('off'); setTimeout(() => ($('#boot').hidden = true), 450);
  });
}

// ---------- entrada: matrícula -> (cadastro) -> app ----------
let regMode = false;
function setReg(on, msg) {
  regMode = on; $('#regFields').hidden = !on; $('#lBack').hidden = !on;
  $('#lMat').readOnly = on;
  $('#lBtn span').textContent = on ? 'Cadastrar e entrar' : 'Continuar';
  $('#lLead').textContent = on ? 'Matrícula ainda não cadastrada.' : 'Informe sua matrícula para entrar.';
  if (on) { eqChips($('#lEquipe'), ''); setTimeout(() => $('#lNome').focus(), 50); }
  $('#loginErr').textContent = msg || '';
}
function showAuth(msg) {
  $('#app').hidden = true; $('#auth').hidden = false;
  S.unsubFills?.(); S.unsubConfig?.(); S.unsubFills = S.unsubConfig = null;
  S.cycle = null; S.rows = []; S.fills = {}; S.byK = new Map(); setSelecting(false);
  setReg(false, msg); setTimeout(() => $('#lMat').focus(), 100);
}
bindEqChips($('#lEquipe'));
$('#lBack').addEventListener('click', () => { setReg(false); $('#lMat').value = ''; $('#lMat').focus(); });
$('#loginForm').addEventListener('submit', async (e) => {
  e.preventDefault(); $('#loginErr').textContent = '';
  const mat = cleanMat($('#lMat').value);
  if (mat.length < 3) return ($('#loginErr').textContent = 'Digite sua matrícula.');
  const btn = $('#lBtn'); btn.disabled = true;
  try {
    if (!regMode) {
      const u = await S.store.lookup(mat);
      if (u) await S.store.login(mat); else setReg(true);
    } else {
      const nome = $('#lNome').value.trim().replace(/\s+/g, ' ');
      if (nome.length < 3) throw new Error('Digite seu nome.');
      await S.store.register({ nome, matricula: mat, equipe: $('#lEquipe').dataset.value || '' });
    }
  } catch (err) { $('#loginErr').textContent = err.message; }
  btn.disabled = false;
});

// ---------- app ----------
async function enterApp() {
  $('#auth').hidden = true; $('#app').hidden = false;
  const me = S.me;
  $('#meInit').textContent = initials(me.nome); $('#meName').textContent = me.nome;
  $('#meRole').textContent = 'Matrícula ' + me.matricula + (me.equipe ? ' • Equipe ' + me.equipe : '');
  S.store.listUsers().then((u) => { S.users = u; fillColabList(); }).catch(() => {});
  S.unsubConfig?.();
  S.unsubConfig = S.store.watchConfig((c) => {
    S.config = { equipes: cfg.equipes || [], ...c };
    if (!S.cycle || S.cycle.id !== c.activeCycle) loadCycle(c.activeCycle);
    renderEquipes(); renderCycles();
  });
  renderDocLock();
  route();
}

async function loadCycle(id) {
  S.unsubFills?.(); S.unsubFills = null;
  S.sel.clear(); setSelecting(false);
  if (!id) { S.cycle = null; S.rows = []; S.byK = new Map(); S.fills = {}; renderAll(); return; }
  $('#cTitle').textContent = 'Carregando…';
  try {
    S.cycles = await S.store.listCycles();
    S.cycle = S.cycles.find((c) => c.id === id) || { id, titulo: 'Cíclica' };
    S.rows = await S.store.loadRows(id);
    S.byK = new Map(S.rows.map((r) => [r.k, r]));
  } catch (e) { toast(e.message, 5000); S.cycle = null; S.rows = []; renderAll(); return; }
  S.unsubFills = S.store.watchFills(id, (f, meta) => { S.fills = f || {}; renderAll(true); syncState(meta); }, (e) => toast(e.message, 5000));
  renderAll();
}

function syncState(meta) {
  const el = $('#sync'); if (S.store.mode !== 'firebase') return (el.hidden = true);
  const off = !navigator.onLine; el.hidden = !(off || meta?.pending);
  el.classList.toggle('off', off); el.classList.toggle('busy', !off && !!meta?.pending);
  $('span', el).textContent = off ? 'Offline • salvo no aparelho' : 'Sincronizando…';
  $('use', el).setAttribute('href', off ? '#i-offline' : '#i-cloud');
}
addEventListener('online', () => syncState({})); addEventListener('offline', () => syncState({}));

// ---------- abas ----------
const TABS = ['ciclica', 'painel', 'exportar', 'documento'];
function route() {
  let t = location.hash.slice(1);
  if (!TABS.includes(t)) t = 'ciclica';
  S.tab = t;
  $$('.panel').forEach((p) => (p.hidden = p.id !== t));
  $$('.tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === t)));
  moveInd();
  if (t === 'painel') renderDash();
  if (t === 'exportar') renderExport();
  if (t === 'documento') { renderCycles(); renderDocLock(); }
  $('#selbar').hidden = !(S.selecting && t === 'ciclica');
}
function moveInd() {
  const b = $(`.tab[data-tab="${S.tab}"]`), ind = $('.tab-ind'); if (!b || b.hidden) return;
  ind.style.width = b.offsetWidth + 'px'; ind.style.transform = `translateX(${b.offsetLeft}px)`;
}
addEventListener('hashchange', route); addEventListener('resize', moveInd);
$('#tabs').addEventListener('click', (e) => { const b = e.target.closest('.tab'); if (b) location.hash = b.dataset.tab; });

// menu da conta
$('#meBtn').addEventListener('click', (e) => { e.stopPropagation(); $('#meMenu').hidden = !$('#meMenu').hidden; });
document.addEventListener('click', (e) => { if (!e.target.closest('#meMenu')) $('#meMenu').hidden = true; });
$('#meMenu').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-act]'); if (!b) return; $('#meMenu').hidden = true;
  if (b.dataset.act === 'out') await S.store.logout();
  if (b.dataset.act === 'me') {
    const r = await dialog({
      title: 'Meus dados', yes: 'Salvar',
      html: `<label class="fl"><span>Matrícula</span><input value="${h(S.me.matricula)}" disabled></label>
        <label class="fl"><span>Nome</span><input id="mNome" value="${h(S.me.nome)}" maxlength="60" required></label>
        <div class="field"><span class="fl-l">Equipe</span><div class="chips" id="mEq" role="radiogroup"></div></div>`,
      onOpen: (bd) => { eqChips($('#mEq', bd), S.me.equipe || ''); bindEqChips($('#mEq', bd)); },
      validate: (bd) => { const n = $('#mNome', bd).value.trim(); return n.length >= 3 ? { nome: n, equipe: $('#mEq', bd).dataset.value || '' } : false; },
    });
    if (r) { try { await S.store.updateMe(r); toast('Dados atualizados.'); } catch (err) { toast(err.message, 5000); } }
  }
});

// ---------- cálculo ----------
function statusOf(k) { const f = S.fills[k]; return !f ? 'none' : f.ok === 'EXECUTADO' ? 'ok' : 'pe'; }
function counts(rows) { const c = { total: rows.length, ok: 0, pe: 0, none: 0 }; for (const r of rows) c[statusOf(r.k)]++; return c; }

// ---------- render geral ----------
function renderAll(fillsOnly) {
  const has = !!S.cycle && S.rows.length > 0;
  $('#noCycle').hidden = has; $('#listWrap').hidden = !has;
  $('#cTitle').textContent = S.cycle ? 'Cíclica ' + (S.cycle.titulo || '') : 'Sem documento ativo';
  renderHero();
  if (!fillsOnly) renderChips();
  renderList(fillsOnly);
  if (S.tab === 'painel') renderDash();
  if (S.tab === 'exportar') renderExport();
}
function renderHero() {
  const c = counts(S.rows); const t = c.total || 1;
  $('.p-ex').style.width = (c.ok / t) * 100 + '%'; $('.p-pe').style.width = (c.pe / t) * 100 + '%';
  const st = (cls, n, l, f) => `<button class="stat ${cls}" data-st="${f}"><dt>${l}</dt><dd>${n}</dd></button>`;
  $('#cStats').innerHTML = st('', c.total, 'OS no mês', 'todas') + st('s-ok', c.ok, 'Executadas', 'ok') + st('s-pe', c.pe, 'Pendentes', 'pe') + st('s-none', c.none, 'Sem preencher', 'none');
}
$('#cStats').addEventListener('click', (e) => { const b = e.target.closest('[data-st]'); if (b) { S.f.status = b.dataset.st; renderChips(); renderList(); $('#listWrap').scrollIntoView({ behavior: 'smooth' }); } });

function patios() {
  const m = new Map();
  for (const r of S.rows) { const o = m.get(r.patio) || { n: 0, done: 0 }; o.n++; if (S.fills[r.k]) o.done++; m.set(r.patio, o); }
  return [...m.entries()].sort((a, b) => a[0].localeCompare(b[0]));
}
function renderChips() {
  const ac = new Map(); for (const r of S.rows) ac.set(r.aba, (ac.get(r.aba) || 0) + 1);
  if (S.f.patio && !S.rows.some((r) => r.patio === S.f.patio)) S.f.patio = '';
  $('#patioLbl').textContent = S.f.patio || 'Todos';
  $('#patioBtn').classList.toggle('on', !!S.f.patio);
  const chip = (attr, v, label, n, on) => `<button class="chip" data-${attr}="${h(v)}" aria-pressed="${on}">${h(label)}${n != null ? `<small>${n}</small>` : ''}</button>`;
  const sts = [['todas', 'Todas'], ['none', 'Sem preencher'], ['ok', 'Executadas'], ['pe', 'Pendentes']];
  $('#statusChips').innerHTML = sts.map(([v, l]) => chip('status', v, l, null, S.f.status === v)).join('');
  $('#abaChips').hidden = ac.size < 2;
  if (S.f.aba && !ac.has(S.f.aba)) S.f.aba = '';
  $('#abaChips').innerHTML = chip('aba', '', 'Todas as abas', null, !S.f.aba) + [...ac.entries()].map(([a, n]) => chip('aba', a, a, n, S.f.aba === a)).join('');
}
$('#statusChips').addEventListener('click', (e) => { const b = e.target.closest('[data-status]'); if (!b) return; S.f.status = b.dataset.status; renderChips(); renderList(); });
$('#abaChips').addEventListener('click', (e) => { const b = e.target.closest('[data-aba]'); if (!b) return; S.f.aba = b.dataset.aba; renderChips(); renderList(); });
let qT;
$('#q').addEventListener('input', () => { $('#qClear').hidden = !$('#q').value; clearTimeout(qT); qT = setTimeout(() => { S.f.q = $('#q').value; renderList(); }, 120); });
$('#qClear').addEventListener('click', () => { $('#q').value = ''; S.f.q = ''; $('#qClear').hidden = true; renderList(); $('#q').focus(); });

// ---------- seletor de pátio (com pesquisa) ----------
function renderPList() {
  const q = norm($('#pQ').value);
  const all = patios().filter(([p]) => !q || norm(p).includes(q));
  const item = (v, label, n, done) => {
    const pct = n ? Math.round((done / n) * 100) : 0;
    return `<button class="pitem${S.f.patio === v ? ' on' : ''}" data-pv="${h(v)}"><b>${h(label)}</b><span class="pbar"><i style="width:${pct}%"></i></span><small>${done}/${n}</small>${S.f.patio === v ? ICON('check') : ''}</button>`;
  };
  const tot = S.rows.length; const done = S.rows.filter((r) => S.fills[r.k]).length;
  $('#pList').innerHTML = (q ? '' : item('', 'Todos os pátios', tot, done)) + all.map(([p, o]) => item(p, p, o.n, o.done)).join('') ||
    '<p class="muted" style="padding:16px">Nenhum pátio com esse nome.</p>';
}
function openPSheet() { $('#pQ').value = ''; renderPList(); $('#pSheet').hidden = false; setTimeout(() => $('#pQ').focus(), 80); }
function closePSheet() { const s = $('#pSheet'); s.classList.add('closing'); setTimeout(() => { s.hidden = true; s.classList.remove('closing'); }, 200); }
$('#patioBtn').addEventListener('click', openPSheet);
$('#pQ').addEventListener('input', renderPList);
$('#pQ').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); $('#pList [data-pv]')?.click(); } });
$('#pSheet').addEventListener('click', (e) => {
  if (e.target.closest('[data-pclose]')) return closePSheet();
  const b = e.target.closest('[data-pv]'); if (!b) return;
  S.f.patio = b.dataset.pv; ls.set('patio', S.f.patio); renderChips(); renderList(); closePSheet();
});

function applyFilters() {
  const q = norm(S.f.q); const terms = q ? q.split(' ') : [];
  return S.rows.filter((r) => {
    if (S.f.patio && r.patio !== S.f.patio) return false;
    if (S.f.aba && r.aba !== S.f.aba) return false;
    if (S.f.status !== 'todas' && statusOf(r.k) !== S.f.status) return false;
    if (terms.length) {
      const f = S.fills[r.k];
      const hay = r._n || (r._n = norm([r.os, r.ativo, r.patio, r.servico, r.desc].join(' ')));
      const extra = f ? norm(f.colab + ' ' + (f.equipe || '')) : '';
      return terms.every((t) => hay.includes(t) || extra.includes(t));
    }
    return true;
  });
}
function hi(text) {
  const t = String(text ?? ''); if (!S.f.q.trim()) return h(t);
  const terms = S.f.q.trim().split(/\s+/).filter((x) => x.length > 1).map((x) => x.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  if (!terms.length) return h(t);
  return h(t).replace(new RegExp('(' + terms.join('|') + ')', 'gi'), '<mark>$1</mark>');
}

const PAGE = 80;
function renderList(keepScroll) {
  S.filtered = applyFilters();
  const c = counts(S.filtered);
  $('#count').textContent = `${S.filtered.length === S.rows.length ? S.rows.length + ' OS' : S.filtered.length + ' de ' + S.rows.length + ' OS'} • ${c.ok} executadas • ${c.pe} pendentes • ${c.none} sem preencher`;
  const want = keepScroll ? Math.max(S.shown, PAGE) : PAGE;
  S.shown = 0; $('#list').innerHTML = '';
  if (!S.filtered.length && S.rows.length) $('#list').innerHTML = '<div class="empty"><b>Nada encontrado</b><span>Ajuste a pesquisa ou os filtros.</span></div>';
  renderMore(want);
  $('#list').classList.toggle('selecting', S.selecting);
}
function itemHtml(r) {
  const f = S.fills[r.k]; const st = statusOf(r.k); const on = S.sel.has(r.k);
  const meta = f ? `<span class="pill ${st}">${f.ok === 'EXECUTADO' ? 'EXECUTADO' : 'PENDENTE'}</span>` +
    (f.dia ? `<span>${brDate(f.dia)}${f.hIni ? ' • ' + f.hIni + (f.hFim ? '–' + f.hFim : '') : ''}</span>` : '') +
    (f.colab ? `<span>${hi(f.colab)}</span>` : '') + (f.equipe ? `<span>${h(f.equipe)}</span>` : '') : `<span>${hi(r.desc)}</span>`;
  return `<button class="it${on ? ' sel' : ''}" data-k="${h(r.k)}">
    <span class="cb${on ? ' on' : ''}">${ICON('check')}</span>
    <span class="st ${st === 'none' ? '' : st}">${st === 'ok' ? ICON('check') : st === 'pe' ? ICON('clock') : ''}</span>
    <span class="it-main"><span class="it-sv">${hi(r.servico || r.desc)}</span><span class="it-meta">${meta}</span></span>
    <span class="it-os">${hi(r.os)}</span></button>`;
}
const gk = (r) => r.patio + '|' + r.ativo;
function renderMore(n = PAGE) {
  const list = $('#list'); const F = S.filtered;
  let stop = Math.min(F.length, S.shown + n);
  if (S.shown >= stop) return;
  while (stop < F.length && gk(F[stop]) === gk(F[stop - 1])) stop++;
  let html = ''; let i = S.shown;
  while (i < stop) { const g = gk(F[i]); const items = []; while (i < stop && gk(F[i]) === g) items.push(F[i++]); html += groupHtml(items[0], g, items); }
  S.shown = stop; list.insertAdjacentHTML('beforeend', html);
}
function groupStats(g) {
  const all = S.filtered.filter((r) => gk(r) === g);
  return { n: all.length, done: all.filter((r) => S.fills[r.k]).length, sel: all.filter((r) => S.sel.has(r.k)).length };
}
function groupHtml(r, g, items) {
  const s = groupStats(g);
  return `<div class="grp" data-g="${h(g)}"><div class="grp-h" data-gh="${h(g)}">
    <span class="cb ${s.sel === s.n ? 'on' : s.sel ? 'part' : ''}">${ICON('check')}</span>
    <span class="tag">${h(r.patio)}</span><b>${hi(r.ativo || '—')}</b>
    <span class="gcount${s.done === s.n ? ' done' : ''}">${s.done}/${s.n}</span></div>${items.map(itemHtml).join('')}</div>`;
}
function updGroupHead(el) {
  const s = groupStats(el.dataset.g); const cb = $('.grp-h .cb', el);
  cb.className = 'cb ' + (s.sel === s.n ? 'on' : s.sel ? 'part' : '');
  const gc = $('.gcount', el); gc.textContent = `${s.done}/${s.n}`; gc.classList.toggle('done', s.done === s.n);
}
new IntersectionObserver((en) => { if (en[0].isIntersecting && S.tab === 'ciclica') renderMore(); }, { rootMargin: '600px' }).observe($('#more'));

// clique / toque longo
let pressT, pressed = false;
$('#list').addEventListener('pointerdown', (e) => {
  const it = e.target.closest('.it'); if (!it || S.selecting) return; pressed = false;
  pressT = setTimeout(() => { pressed = true; navigator.vibrate?.(30); setSelecting(true); toggleSel([it.dataset.k]); }, 520);
});
['pointerup', 'pointerleave', 'pointercancel'].forEach((ev) => $('#list').addEventListener(ev, () => clearTimeout(pressT), { passive: true }));
addEventListener('scroll', () => clearTimeout(pressT), { passive: true });
$('#list').addEventListener('click', (e) => {
  if (pressed) { pressed = false; return; }
  const gh = e.target.closest('.grp-h');
  if (gh) {
    if (!S.selecting) setSelecting(true);
    const ks = S.filtered.filter((r) => gk(r) === gh.dataset.gh).map((r) => r.k);
    toggleSel(ks, !ks.every((k) => S.sel.has(k))); return;
  }
  const it = e.target.closest('.it'); if (!it) return;
  if (S.selecting) toggleSel([it.dataset.k]); else openFill([it.dataset.k]);
});
function toggleSel(ks, force) {
  for (const k of ks) { const on = force ?? !S.sel.has(k); if (on) S.sel.add(k); else S.sel.delete(k); }
  for (const k of ks) {
    const el = $(`.it[data-k="${CSS.escape(k)}"]`); if (!el) continue;
    const on = S.sel.has(k); el.classList.toggle('sel', on); $('.cb', el).classList.toggle('on', on);
  }
  new Set(ks.map((k) => { const r = S.byK.get(k); return r && gk(r); })).forEach((g) => { const el = $(`.grp[data-g="${CSS.escape(g)}"]`); el && updGroupHead(el); });
  updSelbar();
}
function setSelecting(on) {
  S.selecting = on; if (!on) S.sel.clear();
  $('#selToggle').setAttribute('aria-pressed', String(on));
  $('#list').classList.toggle('selecting', on);
  if (!on) { $$('#list .it.sel').forEach((el) => { el.classList.remove('sel'); $('.cb', el).classList.remove('on'); }); $$('#list .grp').forEach(updGroupHead); }
  $('#selbar').hidden = !(on && S.tab === 'ciclica'); updSelbar();
}
function updSelbar() { $('#selN').textContent = S.sel.size; $('#selFill').disabled = !S.sel.size; }
$('#selToggle').addEventListener('click', () => setSelecting(!S.selecting));
$('#selClose').addEventListener('click', () => setSelecting(false));
$('#selAll').addEventListener('click', () => { const ks = S.filtered.map((r) => r.k); if (ks.length && ks.every((k) => S.sel.has(k))) S.sel.clear(); else ks.forEach((k) => S.sel.add(k)); renderList(true); updSelbar(); });
$('#selFill').addEventListener('click', () => S.sel.size && openFill([...S.sel]));

// ---------- preenchimento ----------
let FK = [];
let colabTokens = [];
function fillColabList() {
  const names = new Set(S.users.map((u) => u.nome));
  ls.get('recentColab', []).forEach((n) => names.add(n));
  $('#colabList').innerHTML = [...names].map((n) => `<option value="${h(n)}">`).join('');
}
function renderTokens() {
  $$('#fColab .tok').forEach((t) => t.remove());
  $('#fColabIn').insertAdjacentHTML('beforebegin', colabTokens.map((n, i) => `<span class="tok">${h(n)}<button type="button" data-rm="${i}" aria-label="Remover ${h(n)}">${ICON('x')}</button></span>`).join(''));
  const eq = $('#fEquipe').dataset.value;
  const names = new Set([S.me?.nome, ...ls.get('recentColab', []), ...S.users.filter((u) => eq && u.equipe === eq).map((u) => u.nome)]);
  $('#fColabSugg').innerHTML = [...names].filter((n) => n && !colabTokens.includes(n)).slice(0, 8).map((n) => `<button type="button" class="chip" data-add="${h(n)}">${h(n)}</button>`).join('');
}
function addToken(v) {
  splitColab(v.replace(/;/g, ',')).forEach((n) => { if (!colabTokens.some((t) => t.toLowerCase() === n.toLowerCase())) colabTokens.push(n); });
  $('#fColabIn').value = ''; renderTokens();
}
$('#fColab').addEventListener('click', (e) => { const b = e.target.closest('[data-rm]'); if (b) { colabTokens.splice(+b.dataset.rm, 1); renderTokens(); } else $('#fColabIn').focus(); });
$('#fColabIn').addEventListener('keydown', (e) => {
  if ((e.key === 'Enter' || e.key === ',') && e.target.value.trim()) { e.preventDefault(); addToken(e.target.value); }
  else if (e.key === 'Enter') e.preventDefault();
  else if (e.key === 'Backspace' && !e.target.value && colabTokens.length) { colabTokens.pop(); renderTokens(); }
});
$('#fColabIn').addEventListener('change', (e) => { if (e.target.value.trim()) addToken(e.target.value); });
$('#fColabIn').addEventListener('blur', (e) => { if (e.target.value.trim()) addToken(e.target.value); });
$('#fColabSugg').addEventListener('click', (e) => { const b = e.target.closest('[data-add]'); if (b) addToken(b.dataset.add); });
bindEqChips($('#fEquipe'), renderTokens);
['#fImed', '#fNec'].forEach((s) => bindSeg($(s)));
bindSeg($('#fOk'), () => ($('#fErr').textContent = ''));
$$('.now').forEach((b) => b.addEventListener('click', () => { $('#' + b.dataset.now).value = nowHM(); }));

function openFill(keys) {
  FK = keys.filter((k) => S.byK.has(k)); if (!FK.length) return;
  const single = FK.length === 1; const r0 = S.byK.get(FK[0]); const f0 = single ? S.fills[FK[0]] : null;
  $('#shEyebrow').textContent = single ? `${r0.patio} • OS ${r0.os}` : 'Preenchimento em lote';
  $('#shTitle').textContent = single ? r0.ativo || r0.os : `Preencher ${FK.length} OS de uma vez`;
  if (single) {
    $('#shInfo').innerHTML = `<dl class="info"><dt>OS</dt><dd class="os">${h(r0.os)}</dd><dt>Pátio</dt><dd>${h(r0.patio)}</dd><dt>Ativo</dt><dd>${h(r0.ativo)}</dd>
      <dt>Serviço</dt><dd>${h(r0.servico)}</dd><dt>Descrição</dt><dd>${h(r0.desc)}</dd>${new Set(S.rows.map((r) => r.aba)).size > 1 ? `<dt>Aba</dt><dd>${h(r0.aba)}</dd>` : ''}</dl>` +
      (f0 ? `<p class="last">Preenchido por <b>${h(f0.by || '—')}</b> ${h(ago(f0.at))}</p>` : '');
  } else {
    const filled = FK.filter((k) => S.fills[k]).length;
    $('#shInfo').innerHTML = `<div class="bulk-list">${FK.map((k) => { const r = S.byK.get(k); return `<span><b>${h(r.os)}</b><em>${h(r.ativo)} — ${h(r.servico)}</em></span>`; }).join('')}</div>` +
      (filled ? `<p class="warn">${filled} dessas OS já têm preenchimento e serão substituídas.</p>` : '');
  }
  const last = ls.get('last', {});
  const base = f0 || { ok: 'EXECUTADO', dia: todayISO(), hIni: last.hIni && last.day === todayISO() ? last.hIni : nowHM(), hFim: nowHM(), imediata: 'N', necessidade: 'N',
    colab: (last.colab?.length ? last.colab : [S.me.nome]).join(', '), equipe: last.equipe ?? S.me.equipe ?? '' };
  setSeg($('#fOk'), base.ok); $('#fDia').value = base.dia || todayISO(); $('#fIni').value = base.hIni || ''; $('#fFim').value = base.hFim || '';
  setSeg($('#fImed'), base.imediata || 'N'); setSeg($('#fNec'), base.necessidade || 'N');
  colabTokens = splitColab(base.colab);
  eqChips($('#fEquipe'), base.equipe || ''); renderTokens(); $('#fColabIn').value = '';
  $('#fClear').hidden = !FK.some((k) => S.fills[k]);
  $('#fErr').textContent = '';
  const sh = $('#sheet'); sh.hidden = false; $('.sheet-body', sh).scrollTop = 0;
  history.pushState({ sheet: 1 }, '');
}
function closeSheet(fromPop) {
  const sh = $('#sheet'); if (sh.hidden) return;
  sh.classList.add('closing'); setTimeout(() => { sh.hidden = true; sh.classList.remove('closing'); }, 220);
  if (!fromPop && history.state?.sheet) history.back();
}
addEventListener('popstate', () => closeSheet(true));
$('#sheet').addEventListener('click', (e) => { if (e.target.closest('[data-close]')) closeSheet(); });
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape' || !$('#dlg').hidden) return;
  if (!$('#pSheet').hidden) closePSheet(); else if (!$('#sheet').hidden) closeSheet();
});

$('#fillForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  if ($('#fColabIn').value.trim()) addToken($('#fColabIn').value);
  const fill = {
    ok: segVal($('#fOk')), dia: $('#fDia').value, hIni: $('#fIni').value, hFim: $('#fFim').value,
    imediata: segVal($('#fImed')), necessidade: segVal($('#fNec')), colab: colabTokens.join(', '), equipe: $('#fEquipe').dataset.value || '',
    by: S.me.nome, mat: S.me.matricula, at: Date.now(),
  };
  const miss = [];
  if (!fill.ok) miss.push('OK');
  if (fill.ok === 'EXECUTADO') { if (!fill.dia) miss.push('Dia'); if (!fill.hIni) miss.push('H início'); if (!fill.hFim) miss.push('H fim'); }
  if (!fill.colab) miss.push('Colaboradores');
  if (miss.length) { $('#fErr').textContent = 'Preencha: ' + miss.join(', ') + '.'; return; }
  const btn = $('#fSave'); btn.disabled = true;
  try {
    await S.store.saveFills(S.cycle.id, FK.map((k) => ({ k, g: S.byK.get(k).g, fill })));
    ls.set('last', { colab: colabTokens, equipe: fill.equipe, hIni: fill.hIni, day: todayISO() });
    ls.set('recentColab', [...new Set([...colabTokens, ...ls.get('recentColab', [])])].slice(0, 20));
    fillColabList();
    toast(FK.length > 1 ? `${FK.length} OS preenchidas ✔` : `OS ${S.byK.get(FK[0]).os} salva ✔`);
    if (FK.length > 1) setSelecting(false);
    closeSheet();
  } catch (err) { $('#fErr').textContent = err.message; }
  btn.disabled = false;
});
$('#fClear').addEventListener('click', async () => {
  const n = FK.filter((k) => S.fills[k]).length;
  const ok = await dialog({ title: 'Limpar preenchimento?', html: `<p>Apaga o preenchimento de <b>${n}</b> OS. Elas voltam para “sem preencher”.</p>`, yes: 'Limpar', danger: true });
  if (!ok) return;
  try { await S.store.saveFills(S.cycle.id, FK.filter((k) => S.fills[k]).map((k) => ({ k, g: S.byK.get(k).g, fill: null }))); toast('Preenchimento apagado.'); if (FK.length > 1) setSelecting(false); closeSheet(); }
  catch (err) { $('#fErr').textContent = err.message; }
});

// =====================================================================
// PAINEL — filtros (período, pátio, equipe, colaborador) + visões
// =====================================================================
function dRange() {
  const D = S.d; const t = todayISO();
  if (D.per === 'today') return [t, t];
  if (D.per === '7' || D.per === '30') { const d = new Date(); d.setDate(d.getDate() - (+D.per - 1)); return [isoOf(d), t]; }
  if (D.per === 'custom') return [D.from || '0000', D.to || '9999'];
  return null;
}
const personFilter = () => !!(S.d.eq || S.d.col || S.d.per !== 'all');
function dData() {
  const D = S.d; const rg = dRange();
  const rows = S.rows.filter((r) => !D.patio || r.patio === D.patio);
  const items = []; // {r, f}
  for (const r of rows) {
    const f = S.fills[r.k]; if (!f) continue;
    if (D.eq && f.equipe !== D.eq) continue;
    if (D.col && !splitColab(f.colab).includes(D.col)) continue;
    if (rg && (!f.dia || f.dia < rg[0] || f.dia > rg[1])) continue;
    items.push({ r, f });
  }
  return { rows, items, rg };
}
function fillSelect(sel, opts, v, all) { sel.innerHTML = `<option value="">${all}</option>` + opts.map((o) => `<option ${o === v ? 'selected' : ''}>${h(o)}</option>`).join(''); }
function renderFilters() {
  const D = S.d;
  fillSelect($('#dPatio'), patios().map(([p]) => p), D.patio, 'Todos');
  const eqs = new Set(equipes()); const cols = new Set();
  for (const f of Object.values(S.fills)) { if (f.equipe) eqs.add(f.equipe); splitColab(f.colab).forEach((c) => cols.add(c)); }
  fillSelect($('#dEq'), [...eqs].sort(), D.eq, 'Todas');
  fillSelect($('#dCol'), [...cols].sort((a, b) => a.localeCompare(b)), D.col, 'Todos');
  $('#dPer').value = D.per;
  $('#dFromW').hidden = $('#dToW').hidden = D.per !== 'custom';
  $('#dClear').hidden = !(D.patio || D.eq || D.col || D.per !== 'all');
  const act = [D.per !== 'all' && $('#dPer').selectedOptions[0].text, D.patio, D.eq, D.col].filter(Boolean);
  $('#dFiltN').textContent = act.length || ''; $('#dFiltSum').textContent = act.length ? act.join(' • ') : 'Filtros: período, pátio, equipe, colaborador';
  $$('#dFilters .sel').forEach((l) => l.classList.toggle('on', !!$('select,input', l).value && !['all', ''].includes($('select,input', l).value)));
}
$('#dFilters').addEventListener('change', (e) => {
  const D = S.d; const id = e.target.id;
  if (id === 'dPer') { D.per = e.target.value; if (D.per === 'custom' && !D.from) { const d = new Date(); d.setDate(1); D.from = isoOf(d); D.to = todayISO(); $('#dFrom').value = D.from; $('#dTo').value = D.to; } }
  if (id === 'dFrom') D.from = e.target.value; if (id === 'dTo') D.to = e.target.value;
  if (id === 'dPatio') D.patio = e.target.value; if (id === 'dEq') D.eq = e.target.value; if (id === 'dCol') D.col = e.target.value;
  renderDash();
});
$('#dFiltBtn').addEventListener('click', (e) => { const o = !$('#dFilters').classList.contains('open'); $('#dFilters').classList.toggle('open', o); e.currentTarget.setAttribute('aria-expanded', String(o)); });
$('#dClear').addEventListener('click', () => { Object.assign(S.d, { per: 'all', patio: '', eq: '', col: '' }); renderDash(); });
$('#dNav').addEventListener('click', (e) => { const b = e.target.closest('[data-v]'); if (!b) return; S.d.view = b.dataset.v; renderDash(); b.scrollIntoView({ inline: 'nearest', block: 'nearest' }); });

// tooltip compartilhado dos gráficos
const tip = $('#tip');
$('#dash').addEventListener('pointermove', (e) => {
  const m = e.target.closest('[data-tip]'); if (!m) return (tip.hidden = true);
  tip.innerHTML = m.dataset.tip; tip.hidden = false;
  const r = $('#painel').getBoundingClientRect(); const w = tip.offsetWidth;
  tip.style.left = Math.min(Math.max(8, e.clientX - r.left - w / 2), r.width - w - 8) + 'px';
  tip.style.top = e.clientY - r.top - tip.offsetHeight - 14 + 'px';
});
$('#dash').addEventListener('pointerleave', () => (tip.hidden = true));

const LEG = `<div class="legend-row"><span><i class="k-ok"></i>Executadas</span><span><i class="k-pe"></i>Pendentes</span></div>`;
function hbar(label, ok, pe, max, extra = '', attrs = '') {
  const tot = ok + pe; const w = (v) => (max ? (v / max) * 100 : 0);
  return `<button class="hb" ${attrs} data-tip="<b>${h(label)}</b><br>Executadas: ${ok}<br>Pendentes: ${pe}${extra ? '<br>' + h(extra) : ''}">
    <span class="hb-l">${h(label)}</span><span class="hb-t"><i class="x" style="width:${w(ok)}%"></i><i class="p" style="width:${w(pe)}%"></i></span><span class="hb-v">${tot}${extra ? `<small>${h(extra)}</small>` : ''}</span></button>`;
}
function emptyCard(t, s) { return `<div class="empty"><b>${t}</b><span>${s}</span></div>`; }

function renderDash() {
  if (!S.cycle) { $('#dash').innerHTML = emptyCard('Sem documento ativo', 'Envie a planilha do mês na aba Documento.'); $('#dFilters').hidden = $('#dNav').hidden = true; return; }
  $('#dFilters').hidden = $('#dNav').hidden = false;
  $('#pTitle').textContent = 'Painel • ' + (S.cycle.titulo || '');
  renderFilters();
  $$('#dNav [data-v]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.v === S.d.view)));
  const data = dData(); tip.hidden = true;
  const v = S.d.view;
  $('#dash').innerHTML = v === 'geral' ? viewGeral(data) : v === 'patios' ? viewPatios(data) : v === 'equipes' ? viewEquipes(data)
    : v === 'pessoas' ? viewPessoas(data) : v === 'atencao' ? viewAtencao(data) : viewAtividade(data);
}

function viewGeral({ rows, items, rg }) {
  const ok = items.filter((x) => x.f.ok === 'EXECUTADO').length; const pe = items.length - ok;
  const none = rows.filter((r) => !S.fills[r.k]).length;
  const pct = rows.length ? Math.round((ok / rows.length) * 100) : 0;
  const pf = personFilter();
  const kpi = (cls, n, l, sub = '') => `<div class="kpi ${cls}"><b>${n}</b><span>${l}</span>${sub ? `<small>${sub}</small>` : ''}</div>`;
  // série diária
  let days = [];
  const ds = items.map((x) => x.f.dia).filter(Boolean).sort();
  let a = rg ? rg[0] : ds[0], b = rg ? rg[1] : ds[ds.length - 1];
  if (a === '0000') a = ds[0]; if (b === '9999') b = ds[ds.length - 1] || todayISO();
  if (a && b) {
    const d = new Date(a + 'T12:00'); const end = new Date(b + 'T12:00');
    if ((end - d) / 864e5 > 62) d.setTime(end.getTime() - 61 * 864e5);
    for (; d <= end; d.setDate(d.getDate() + 1)) days.push(isoOf(d));
  }
  const byDay = new Map(days.map((x) => [x, { ok: 0, pe: 0 }]));
  for (const { f } of items) { const o = byDay.get(f.dia); if (o) o[f.ok === 'EXECUTADO' ? 'ok' : 'pe']++; }
  const max = Math.max(1, ...[...byDay.values()].map((o) => o.ok + o.pe));
  const nice = max <= 5 ? 5 : Math.ceil(max / 5) * 5;
  const cols = days.map((d) => { const o = byDay.get(d); const [y, m, dd] = d.split('-');
    return `<div class="col" data-tip="<b>${dd}/${m}/${y}</b><br>Executadas: ${o.ok}<br>Pendentes: ${o.pe}"><div class="col-s"><i class="p" style="height:${(o.pe / nice) * 100}%"></i><i class="x" style="height:${(o.ok / nice) * 100}%"></i></div><span>${+dd}</span></div>`; }).join('');
  const chart = days.length ? `<div class="vchart" style="--n:${days.length}"><div class="grid"><span>${nice}</span><span>${nice / 2}</span><span>0</span></div><div class="cols">${cols}</div></div>`
    : '<p class="muted">Nenhum preenchimento com data no filtro.</p>';
  const imed = items.filter((x) => x.f.imediata === 'S').length; const nec = items.filter((x) => x.f.necessidade === 'S').length;
  return `<div class="kpis">
      ${kpi('k-ok', ok, 'Executadas', pf ? '' : pct + '% do plano')}
      ${kpi('k-pe', pe, 'Pendentes')}
      ${pf ? kpi('', items.length, 'Preenchidas no filtro') : kpi('', none, 'Sem preencher', 'de ' + rows.length + ' OS')}
      <button class="kpi k-alert" data-go="atencao"><b>${imed + nec}</b><span>Pedem atenção</span><small>${imed} imediata • ${nec} necessidade</small></button>
    </div>
    ${pf ? '' : `<div class="card pad"><div class="card-t"><b>Andamento do plano</b><span class="muted small">${ok + pe} de ${rows.length} OS preenchidas</span></div>
      <div class="big-prog" data-tip="Executadas: ${ok}<br>Pendentes: ${pe}<br>Sem preencher: ${none}"><i class="x" style="width:${rows.length ? (ok / rows.length) * 100 : 0}%"></i><i class="p" style="width:${rows.length ? (pe / rows.length) * 100 : 0}%"></i></div>${LEG}</div>`}
    <div class="card pad"><div class="card-t"><b>Preenchimentos por dia</b><span class="muted small">pela data informada (DIA)</span></div>${LEG}${chart}</div>`;
}
function viewPatios({ items }) {
  const by = new Map();
  for (const r of S.rows) if (!S.d.patio || r.patio === S.d.patio) { const o = by.get(r.patio) || { n: 0, ok: 0, pe: 0 }; o.n++; by.set(r.patio, o); }
  for (const { r, f } of items) by.get(r.patio)[f.ok === 'EXECUTADO' ? 'ok' : 'pe']++;
  const list = [...by.entries()].sort((a, b) => (b[1].ok + b[1].pe) / b[1].n - (a[1].ok + a[1].pe) / a[1].n || a[0].localeCompare(b[0]));
  return `<div class="card pad"><div class="card-t"><b>Andamento por pátio</b><span class="muted small">toque para abrir a lista</span></div>${LEG}
    <div class="hbars pct">${list.map(([p, o]) => hbar(p, o.ok, o.pe, o.n, `${Math.round(((o.ok + o.pe) / o.n) * 100)}% de ${o.n}`, `data-p="${h(p)}"`)).join('')}</div></div>`;
}
function viewEquipes({ items }) {
  const by = new Map();
  for (const { f } of items) { const k = f.equipe || 'Sem equipe'; const o = by.get(k) || { ok: 0, pe: 0 }; o[f.ok === 'EXECUTADO' ? 'ok' : 'pe']++; by.set(k, o); }
  if (!by.size) return emptyCard('Nenhum preenchimento', 'Nada encontrado com os filtros escolhidos.');
  const list = [...by.entries()].sort((a, b) => b[1].ok + b[1].pe - (a[1].ok + a[1].pe)); const max = list[0][1].ok + list[0][1].pe;
  return `<div class="card pad"><div class="card-t"><b>OS preenchidas por equipe</b><span class="muted small">toque para filtrar</span></div>${LEG}
    <div class="hbars">${list.map(([e, o]) => hbar(e, o.ok, o.pe, max, '', e !== 'Sem equipe' ? `data-eq="${h(e)}"` : '')).join('')}</div></div>`;
}
function viewPessoas({ items }) {
  const by = new Map();
  for (const { f } of items) for (const c of splitColab(f.colab)) { if (S.d.col && c !== S.d.col) continue; const o = by.get(c) || { ok: 0, pe: 0, last: 0 }; o[f.ok === 'EXECUTADO' ? 'ok' : 'pe']++; o.last = Math.max(o.last, f.at || 0); by.set(c, o); }
  if (!by.size) return emptyCard('Nenhum colaborador', 'Nada encontrado com os filtros escolhidos.');
  const list = [...by.entries()].sort((a, b) => b[1].ok + b[1].pe - (a[1].ok + a[1].pe)); const max = list[0][1].ok + list[0][1].pe;
  return `<div class="card pad"><div class="card-t"><b>OS por colaborador</b><span class="muted small">${by.size} pessoas • toque para filtrar</span></div>${LEG}
    <div class="hbars people">${list.map(([n, o]) => hbar(n, o.ok, o.pe, max, '', `data-col="${h(n)}"`)).join('')}</div></div>`;
}
function osRow(r, f, badge = '') {
  const s = f.ok === 'EXECUTADO' ? 'ok' : 'pe';
  return `<button class="feed-i" data-k="${h(r.k)}"><span class="st ${s}">${ICON(s === 'ok' ? 'check' : 'clock')}</span>
    <span class="fi-tx"><b>${h(r.ativo)} <em>OS ${h(r.os)}</em></b><small>${h(r.servico)}</small><small>${brDate(f.dia)} • ${h(f.colab)}${f.equipe ? ' • ' + h(f.equipe) : ''}</small></span>${badge}</button>`;
}
function viewAtencao({ items }) {
  const list = items.filter((x) => x.f.imediata === 'S' || x.f.necessidade === 'S').sort((a, b) => (b.f.at || 0) - (a.f.at || 0));
  if (!list.length) return emptyCard('Nada pedindo atenção', 'Nenhuma OS marcada com manutenção imediata ou necessidade de manutenção.');
  return `<div class="card pad"><div class="card-t"><b>OS que pedem atenção</b><span class="muted small">${list.length} OS • toque para abrir</span></div>
    <div class="feed">${list.map(({ r, f }) => osRow(r, f, `<span class="flags">${f.imediata === 'S' ? '<span class="flag f1">Imediata</span>' : ''}${f.necessidade === 'S' ? '<span class="flag f2">Necessidade</span>' : ''}</span>`)).join('')}</div></div>`;
}
function viewAtividade({ items }) {
  const list = [...items].sort((a, b) => (b.f.at || 0) - (a.f.at || 0)).slice(0, 60);
  if (!list.length) return emptyCard('Nenhum preenchimento', 'Nada encontrado com os filtros escolhidos.');
  return `<div class="card pad"><div class="card-t"><b>Últimos preenchimentos</b><span class="muted small">${items.length} no filtro</span></div>
    <div class="feed">${list.map(({ r, f }) => osRow(r, f, `<small class="when">${h(ago(f.at))}<br>por ${h(f.by || '')}</small>`)).join('')}</div></div>`;
}
$('#dash').addEventListener('click', (e) => {
  const t = e.target.closest('[data-go],[data-p],[data-eq],[data-col],[data-k]'); if (!t) return;
  if (t.dataset.go) { S.d.view = t.dataset.go; return renderDash(); }
  if (t.dataset.eq) { S.d.eq = t.dataset.eq; S.d.view = 'geral'; return renderDash(); }
  if (t.dataset.col) { S.d.col = t.dataset.col; S.d.view = 'geral'; return renderDash(); }
  if (t.dataset.p) { S.f.patio = t.dataset.p; S.f.status = 'todas'; ls.set('patio', S.f.patio); renderChips(); renderList(); location.hash = 'ciclica'; return; }
  if (t.dataset.k) openFill([t.dataset.k]);
});

// ---------- exportar ----------
function exportName(c) { return (c.fileName || 'CICLICA.xlsx').replace(/\.xlsx$/i, '') + '_PREENCHIDA.xlsx'; }
function renderExport() {
  const has = !!S.cycle; $('#exBtn').disabled = !has;
  if (!has) { $('#exName').textContent = 'Sem documento ativo'; $('#exInfo').textContent = ''; return; }
  const c = counts(S.rows);
  $('#exName').textContent = exportName(S.cycle);
  $('#exInfo').textContent = `${S.cycle.titulo || ''} • ${c.ok + c.pe} de ${c.total} OS preenchidas (${c.ok} executadas, ${c.pe} pendentes)`;
}
$('#exEquipe').checked = ls.get('exEquipe', true);
$('#exEquipe').addEventListener('change', (e) => ls.set('exEquipe', e.target.checked));
async function doExport(cycle, rows, fills, btn) {
  btn && (btn.disabled = true); toast('Gerando planilha…', 8000);
  try {
    const bytes = await S.store.getFile(cycle.id);
    if (!bytes) throw new Error('Arquivo original não encontrado.');
    const out = await exportWorkbook(JSZip, bytes, rows, fills, { equipe: $('#exEquipe').checked });
    download(out.bytes, exportName(cycle), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    toast('Planilha baixada ✔');
  } catch (e) { console.error(e); toast('Erro ao exportar: ' + e.message, 5000); }
  btn && (btn.disabled = false);
}
$('#exBtn').addEventListener('click', (e) => S.cycle && doExport(S.cycle, S.rows, S.fills, e.currentTarget));

// ---------- documento (protegido por senha) ----------
function renderDocLock() {
  const u = S.store?.isUnlocked();
  $('#unlockForm').hidden = !!u; $('#manage').hidden = !u; $('#eqWrap').hidden = !u; $('#lockBtn').hidden = !u;
  renderCycles();
}
$('#unlockForm').addEventListener('submit', async (e) => {
  e.preventDefault(); $('#upErr').textContent = ''; const btn = $('#unlockForm .btn'); btn.disabled = true;
  try { await S.store.unlock($('#upPass').value); $('#upPass').value = ''; toast('Área liberada.'); renderDocLock(); }
  catch (err) { $('#upErr').textContent = err.message; }
  btn.disabled = false;
});
$('#lockBtn').addEventListener('click', () => { S.store.lock(); renderDocLock(); toast('Área bloqueada.'); });

const drop = $('#drop');
['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; f && readUpload(f); });
$('#file').addEventListener('change', (e) => { const f = e.target.files[0]; f && readUpload(f); e.target.value = ''; });

async function readUpload(file) {
  if (!/\.xlsx$/i.test(file.name)) return toast('Envie um arquivo .xlsx');
  $('#preview').innerHTML = '<div class="card pad"><div class="spinner"></div></div>';
  try {
    const buf = new Uint8Array(await file.arrayBuffer());
    S.parsed = await parseWorkbook(JSZip, buf); S.upBytes = buf; S.upName = file.name;
    if (!S.parsed.sheets.some((s) => s.role === 'principal')) throw new Error('Não encontrei nenhuma aba com as colunas OS, OK e DIA.');
    renderPreview();
  } catch (e) { console.error(e); $('#preview').innerHTML = `<div class="card pad"><p class="err">${h(e.message)}</p></div>`; }
}
function renderPreview() {
  const p = S.parsed;
  const rows = p.sheets.map((s) => `<tr><td>${s.role === 'principal' ? `<label class="check"><input type="checkbox" data-aba="${h(s.name)}" checked><span>${h(s.name)}</span></label>` : h(s.name)}</td>
    <td><span class="role ${s.role}">${s.role === 'principal' ? 'preencher' : s.role === 'espelho' ? 'espelho de ' + h(s.mirrorOf) : 'oculta (ignorada)'}</span></td><td>${s.items.length}</td></tr>`).join('');
  const prim = p.sheets.filter((s) => s.role === 'principal');
  const pats = new Map(); prim.forEach((s) => s.items.forEach((i) => pats.set(i.patio || '—', (pats.get(i.patio || '—') || 0) + 1)));
  $('#preview').innerHTML = `<div class="card pad prev">
    <div class="card-t"><b>Conferir antes de publicar</b><span class="muted small">${h(S.upName)}</span></div>
    <label class="fl"><span>Nome do documento</span><input id="pvTitle" value="${h(p.titulo || S.upName.replace(/\.xlsx$/i, ''))}"></label>
    <table><thead><tr><th>Aba</th><th>Uso</th><th>OS</th></tr></thead><tbody>${rows}</tbody></table>
    <p class="muted small">Abas “espelho” repetem as OS de outra aba (ex.: uma por pátio) e recebem os mesmos preenchimentos na exportação.</p>
    <div class="chips">${[...pats.entries()].sort().map(([k, n]) => `<span class="chip">${h(k)}<small>${n}</small></span>`).join('')}</div>
    <label class="check"><input type="checkbox" id="pvActive" checked><span>Tornar este o documento ativo agora</span></label>
    <div class="actions"><button class="btn" id="pvCancel">Cancelar</button><button class="btn primary" id="pvPub">${ICON('upload')}<span>Publicar</span></button></div></div>`;
}
$('#preview').addEventListener('click', async (e) => {
  if (e.target.closest('#pvCancel')) { $('#preview').innerHTML = ''; S.parsed = null; return; }
  const btn = e.target.closest('#pvPub'); if (!btn) return;
  const abas = $$('#preview [data-aba]').filter((c) => c.checked).map((c) => c.dataset.aba);
  if (!abas.length) return toast('Escolha pelo menos uma aba.');
  btn.disabled = true; $('span', btn).textContent = 'Publicando…';
  try {
    const rows = buildRows(S.parsed, abas);
    const gid = new Map(); rows.forEach((r, i) => { if (!gid.has(r.patio)) gid.set(r.patio, 'p' + gid.size); r.g = gid.get(r.patio); r.i = i; });
    const meta = { titulo: $('#pvTitle').value.trim() || S.parsed.titulo || 'Cíclica', ym: S.parsed.ym || '', fileName: S.upName, abas, total: rows.length, patios: [...gid.keys()] };
    const id = await S.store.createCycle(meta, rows, S.upBytes);
    if ($('#pvActive').checked) await S.store.setConfig({ activeCycle: id });
    $('#preview').innerHTML = ''; S.parsed = null; toast('Documento publicado ✔'); renderCycles();
  } catch (err) { console.error(err); toast('Erro: ' + err.message, 6000); btn.disabled = false; $('span', btn).textContent = 'Publicar'; }
});

async function renderCycles() {
  if (!S.me) return;
  try { S.cycles = await S.store.listCycles(); } catch (e) { $('#cycles').innerHTML = `<p class="err">${h(e.message)}</p>`; return; }
  const act = S.config.activeCycle; const u = S.store.isUnlocked();
  $('#cycles').innerHTML = S.cycles.map((c) => `<div class="card pad row-card${c.id === act ? ' active' : ''}" data-id="${h(c.id)}">
    <span class="rc-ic">${ICON('file')}</span>
    <span class="rc-tx"><b>${h(c.titulo)}${c.id === act ? '<span class="badge-a">ATIVO</span>' : ''}</b><small>${h(c.fileName || '')} • ${c.total || 0} OS • enviado ${h(ago(c.createdAt))}${c.createdBy ? ' por ' + h(c.createdBy) : ''}</small></span>
    <span class="rc-act">${u && c.id !== act ? `<button class="btn" data-ca="on">${ICON('star')}Ativar</button>` : ''}
      <button class="btn" data-ca="dl">${ICON('download')}Baixar preenchida</button>
      ${u ? `<button class="btn danger" data-ca="rm" aria-label="Excluir">${ICON('trash')}</button>` : ''}</span></div>`).join('') ||
    emptyCard('Nenhum documento enviado ainda', 'Libere a área com a senha de upload e envie a planilha do mês.');
}
$('#cycles').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-ca]'); if (!b) return; const id = b.closest('[data-id]').dataset.id; const c = S.cycles.find((x) => x.id === id);
  try {
    if (b.dataset.ca === 'on') { await S.store.setConfig({ activeCycle: id }); toast(`${c.titulo} agora é o documento ativo.`); }
    if (b.dataset.ca === 'dl') {
      if (id === S.cycle?.id) return doExport(S.cycle, S.rows, S.fills, b);
      b.disabled = true;
      const rows = await S.store.loadRows(id);
      const fills = await new Promise((res) => { let u = null; u = S.store.watchFills(id, (f, meta) => { if (meta?.cache && navigator.onLine) return; res(f); setTimeout(() => u?.(), 0); }); });
      await doExport(c, rows, fills, b);
    }
    if (b.dataset.ca === 'rm') {
      const ok = await dialog({ title: 'Excluir documento?', html: `<p>Apaga <b>${h(c.titulo)}</b> e <b>todos os preenchimentos</b> dele. Não dá para desfazer.</p><p class="muted small">Dica: baixe a planilha preenchida antes.</p>`, yes: 'Excluir', danger: true });
      if (!ok) return;
      await S.store.deleteCycle(id); if (S.config.activeCycle === id) await S.store.setConfig({ activeCycle: null }); toast('Documento excluído.'); renderCycles();
    }
  } catch (err) { toast(err.message, 5000); if (/senha/i.test(err.message)) { S.store.lock(); renderDocLock(); } }
});

// equipes
function renderEquipes() {
  $('#eqList').innerHTML = equipes().map((e, i) => `<span class="chip">${h(e)}<button type="button" data-eq="${i}" aria-label="Remover ${h(e)}">${ICON('x')}</button></span>`).join('') || '<span class="muted">Nenhuma equipe.</span>';
}
$('#eqList').addEventListener('click', async (e) => { const b = e.target.closest('[data-eq]'); if (!b) return; const eq = [...equipes()]; eq.splice(+b.dataset.eq, 1); try { await S.store.setConfig({ equipes: eq }); } catch (err) { toast(err.message, 4000); } });
$('#eqForm').addEventListener('submit', async (e) => { e.preventDefault(); const v = $('#eqNew').value.trim().toUpperCase(); if (!v) return; try { await S.store.setConfig({ equipes: [...new Set([...equipes(), v])] }); $('#eqNew').value = ''; } catch (err) { toast(err.message, 4000); } });

if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});

boot();
