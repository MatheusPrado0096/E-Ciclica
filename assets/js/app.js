// ===== E-Cíclica • aplicativo =====
import { parseWorkbook, buildRows, exportWorkbook, norm } from './xlsx.js';
import { createLocalStore } from './store-local.js';

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
const todayISO = () => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`; };
const nowHM = () => { const d = new Date(); return `${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const brDate = (iso) => (iso ? iso.split('-').reverse().join('/') : '');
const ago = (t) => {
  if (!t) return ''; const s = (Date.now() - t) / 1000;
  if (s < 60) return 'agora'; if (s < 3600) return `há ${Math.floor(s / 60)} min`; if (s < 86400) return `há ${Math.floor(s / 3600)} h`;
  const d = new Date(t); return `${pad(d.getDate())}/${pad(d.getMonth() + 1)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
};
const initials = (n) => String(n || '?').trim().split(/\s+/).slice(0, 2).map((p) => p[0]).join('').toUpperCase();
const ICON = (id) => `<svg class="ic"><use href="#i-${id}"/></svg>`;

// ---------- estado ----------
const S = {
  store: null, me: null, config: { equipes: cfg.equipes || [] }, cycle: null, cycles: [], rows: [], byK: new Map(), fills: {},
  users: [], f: { q: '', patio: ls.get('patio', ''), status: 'todas', aba: '' }, selecting: false, sel: new Set(),
  shown: 0, filtered: [], tab: 'ciclica', unsubFills: null, unsubConfig: null, parsed: null, upBytes: null, upName: '',
};

// ---------- utilidades de interface ----------
let toastT;
function toast(msg, ms = 2600) { const t = $('#toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('show'), ms); }
function setSeg(el, v) { $$('[role=radio]', el).forEach((b) => b.setAttribute('aria-checked', String(b.dataset.v === v))); el.dataset.value = v ?? ''; }
function segVal(el) { return el.dataset.value || ''; }
function bindSeg(el, onChange) { el.addEventListener('click', (e) => { const b = e.target.closest('[role=radio]'); if (!b) return; setSeg(el, b.dataset.v); onChange?.(b.dataset.v); }); }
function download(bytes, name, type) {
  const url = URL.createObjectURL(new Blob([bytes], { type }));
  const a = Object.assign(document.createElement('a'), { href: url, download: name }); document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}

// diálogo genérico (confirmação / formulário pequeno)
function dialog({ title, html = '', yes = 'OK', no = 'Cancelar', danger = false, onOpen, validate }) {
  return new Promise((resolve) => {
    const d = $('#dlg'); $('#dlgTitle').textContent = title; $('#dlgBody').innerHTML = html;
    const y = $('#dlgYes'); y.textContent = yes; y.classList.toggle('danger', danger); y.style.background = danger ? 'var(--danger)' : ''; y.style.color = danger ? '#fff' : '';
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
$('#themeBtn').addEventListener('click', () => { const t = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'; applyTheme(t); try { localStorage.setItem('eciclica-theme', t); } catch {} });
applyTheme(document.documentElement.dataset.theme);

// ---------- inicialização ----------
async function boot() {
  const useFb = !!cfg.firebase?.apiKey;
  try {
    if (useFb) {
      $('#bootMsg').textContent = 'Conectando…';
      const { createFirebaseStore } = await import('./store-firebase.js');
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
    $('#localBanner').hidden = false;
  }
  S.store.onAuth(async (me, msg) => {
    S.me = me;
    if (me) await enterApp(); else await showAuth(msg);
    $('#boot').classList.add('off'); setTimeout(() => ($('#boot').hidden = true), 450);
  });
}

// ---------- entrada ----------
async function showAuth(msg) {
  $('#app').hidden = true; $('#auth').hidden = false;
  S.unsubFills?.(); S.unsubConfig?.(); S.unsubFills = S.unsubConfig = null;
  S.cycle = null; S.rows = []; S.fills = {}; S.byK = new Map(); setSelecting(false);
  let boot = false;
  try { boot = await S.store.needsBootstrap(); } catch (e) { msg = e.message; }
  $('#loginForm').hidden = boot; $('#bootForm').hidden = !boot;
  $('#loginErr').textContent = msg || '';
  setTimeout(() => (boot ? $('#bName') : $('#lUser')).focus(), 100);
}
const cleanUser = (v) => v.trim().toLowerCase();
$('#loginForm').addEventListener('submit', async (e) => {
  e.preventDefault(); const btn = e.submitter || $('#loginForm .btn'); btn.disabled = true; $('#loginErr').textContent = '';
  try { await S.store.login(cleanUser($('#lUser').value), $('#lPass').value); $('#lPass').value = ''; }
  catch (err) { $('#loginErr').textContent = err.message; }
  btn.disabled = false;
});
$('#bootForm').addEventListener('submit', async (e) => {
  e.preventDefault(); $('#bootErr').textContent = '';
  if ($('#bPass').value !== $('#bPass2').value) return ($('#bootErr').textContent = 'As senhas não conferem.');
  const btn = $('#bootForm .btn'); btn.disabled = true;
  try { await S.store.bootstrap({ name: $('#bName').value.trim(), username: cleanUser($('#bUser').value), password: $('#bPass').value }); }
  catch (err) { $('#bootErr').textContent = err.message; }
  btn.disabled = false;
});

// ---------- app ----------
async function enterApp() {
  $('#auth').hidden = true; $('#app').hidden = false;
  const me = S.me; const admin = me.role === 'admin';
  $('#meInit').textContent = initials(me.name); $('#meName').textContent = me.name;
  $('#meRole').textContent = (admin ? 'Administrador' : 'Funcionário') + ' • @' + me.username + (me.equipe ? ' • ' + me.equipe : '');
  $$('.admin-only').forEach((el) => (el.hidden = !admin));
  S.store.listUsers().then((u) => { S.users = u; fillColabList(); if (admin) renderUsers(); }).catch(() => {});
  S.unsubConfig?.();
  S.unsubConfig = S.store.watchConfig((c) => {
    S.config = { equipes: cfg.equipes || [], ...c };
    if (!S.cycle || S.cycle.id !== c.activeCycle) loadCycle(c.activeCycle);
    if (admin) { renderEquipes(); renderCycles(); }
  });
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
const TABS = ['ciclica', 'painel', 'exportar', 'admin'];
function route() {
  let t = location.hash.slice(1);
  if (!TABS.includes(t) || (t === 'admin' && S.me?.role !== 'admin')) t = 'ciclica';
  S.tab = t;
  $$('.panel').forEach((p) => (p.hidden = p.id !== t));
  $$('.tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === t)));
  moveInd();
  if (t === 'painel') renderDash();
  if (t === 'exportar') renderExport();
  $('#selbar').hidden = !(S.selecting && t === 'ciclica');
}
function moveInd() {
  const b = $(`.tab[data-tab="${S.tab}"]`), ind = $('.tab-ind'); if (!b || b.hidden) return;
  ind.style.width = b.offsetWidth + 'px'; ind.style.transform = `translateX(${b.offsetLeft}px)`;
}
addEventListener('hashchange', route); addEventListener('resize', moveInd);
$('#tabs').addEventListener('click', (e) => { const b = e.target.closest('.tab'); if (b) { location.hash = b.dataset.tab; } });

// menu da conta
$('#meBtn').addEventListener('click', (e) => { e.stopPropagation(); $('#meMenu').hidden = !$('#meMenu').hidden; });
document.addEventListener('click', (e) => { if (!e.target.closest('#meMenu')) $('#meMenu').hidden = true; });
$('#meMenu').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-act]'); if (!b) return; $('#meMenu').hidden = true;
  if (b.dataset.act === 'out') { await S.store.logout(); }
  if (b.dataset.act === 'pw') {
    await dialog({
      title: 'Trocar minha senha', yes: 'Salvar',
      html: `<label class="fl">Senha atual<input type="password" id="pw0" required autocomplete="current-password"></label>
             <label class="fl">Nova senha<input type="password" id="pw1" required minlength="6" autocomplete="new-password"></label>
             <label class="fl">Repita a nova senha<input type="password" id="pw2" required minlength="6" autocomplete="new-password"></label><p class="err" id="pwErr"></p>`,
      validate: async (b) => {
        const [a, n, r] = ['#pw0', '#pw1', '#pw2'].map((s) => $(s, b).value);
        if (n.length < 6) return ($('#pwErr', b).textContent = 'A nova senha precisa ter pelo menos 6 caracteres.'), false;
        if (n !== r) return ($('#pwErr', b).textContent = 'As senhas não conferem.'), false;
        try { await S.store.changePassword(a, n); toast('Senha alterada.'); return true; } catch (err) { $('#pwErr', b).textContent = err.message; return false; }
      },
    });
  }
});

// ---------- cálculo ----------
function statusOf(k) { const f = S.fills[k]; return !f ? 'none' : f.ok === 'EXECUTADO' ? 'ok' : 'pe'; }
function counts(rows) {
  const c = { total: rows.length, ok: 0, pe: 0, none: 0 };
  for (const r of rows) c[statusOf(r.k)]++;
  return c;
}

// ---------- render geral ----------
function renderAll(fillsOnly) {
  const has = !!S.cycle && S.rows.length > 0;
  $('#noCycle').hidden = has; $('#listWrap').hidden = !has;
  if (!S.cycle) {
    $('#cTitle').textContent = 'Sem documento ativo';
    $('#noCycleMsg').textContent = S.me?.role === 'admin' ? 'Envie a planilha do mês na aba Admin → Documento do mês.' : 'O administrador ainda não enviou a planilha deste mês.';
  } else $('#cTitle').textContent = 'Cíclica ' + (S.cycle.titulo || '');
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

function renderChips() {
  const pc = new Map(); const ac = new Map();
  for (const r of S.rows) { pc.set(r.patio, (pc.get(r.patio) || 0) + 1); ac.set(r.aba, (ac.get(r.aba) || 0) + 1); }
  if (S.f.patio && !pc.has(S.f.patio)) S.f.patio = '';
  const chip = (attr, v, label, n, on) => `<button class="chip" data-${attr}="${h(v)}" aria-pressed="${on}">${h(label)}${n != null ? `<small>${n}</small>` : ''}</button>`;
  $('#patioChips').innerHTML = chip('patio', '', 'Todos os pátios', S.rows.length, !S.f.patio) +
    [...pc.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([p, n]) => chip('patio', p, p, n, S.f.patio === p)).join('');
  const sts = [['todas', 'Todas'], ['none', 'Sem preencher'], ['ok', 'Executadas'], ['pe', 'Pendentes']];
  $('#statusChips').innerHTML = sts.map(([v, l]) => chip('status', v, l, null, S.f.status === v)).join('');
  $('#abaChips').hidden = ac.size < 2;
  if (S.f.aba && !ac.has(S.f.aba)) S.f.aba = '';
  $('#abaChips').innerHTML = chip('aba', '', 'Todas as abas', null, !S.f.aba) + [...ac.entries()].map(([a, n]) => chip('aba', a, a, n, S.f.aba === a)).join('');
}
$('#patioChips').addEventListener('click', (e) => { const b = e.target.closest('[data-patio]'); if (!b) return; S.f.patio = b.dataset.patio; ls.set('patio', S.f.patio); renderChips(); renderList(); });
$('#statusChips').addEventListener('click', (e) => { const b = e.target.closest('[data-status]'); if (!b) return; S.f.status = b.dataset.status; renderChips(); renderList(); });
$('#abaChips').addEventListener('click', (e) => { const b = e.target.closest('[data-aba]'); if (!b) return; S.f.aba = b.dataset.aba; renderChips(); renderList(); });
let qT;
$('#q').addEventListener('input', () => { $('#qClear').hidden = !$('#q').value; clearTimeout(qT); qT = setTimeout(() => { S.f.q = $('#q').value; renderList(); }, 120); });
$('#qClear').addEventListener('click', () => { $('#q').value = ''; S.f.q = ''; $('#qClear').hidden = true; renderList(); $('#q').focus(); });

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
  $('#count').textContent = S.filtered.length === S.rows.length ? `${S.rows.length} OS • ${c.ok} executadas • ${c.pe} pendentes • ${c.none} sem preencher`
    : `${S.filtered.length} de ${S.rows.length} OS`;
  const want = keepScroll ? Math.max(S.shown, PAGE) : PAGE;
  S.shown = 0; $('#list').innerHTML = '';
  if (!S.filtered.length && S.rows.length) $('#list').innerHTML = '<div class="empty"><b>Nada encontrado</b><span>Ajuste a pesquisa ou os filtros.</span></div>';
  renderMore(want);
  $('#list').classList.toggle('selecting', S.selecting);
}
function itemHtml(r) {
  const f = S.fills[r.k]; const st = statusOf(r.k); const on = S.sel.has(r.k);
  let meta = '';
  if (f) meta = `<span class="pill ${st}">${f.ok === 'EXECUTADO' ? 'EXECUTADO' : 'PENDENTE'}</span>` +
    (f.dia ? `<span>${brDate(f.dia)}${f.hIni ? ' • ' + f.hIni + (f.hFim ? '–' + f.hFim : '') : ''}</span>` : '') +
    (f.colab ? `<span>${hi(f.colab)}</span>` : '') + (f.equipe ? `<span>${h(f.equipe)}</span>` : '');
  else meta = `<span>${hi(r.desc)}</span>`;
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
  while (stop < F.length && gk(F[stop]) === gk(F[stop - 1])) stop++; // não corta um ativo no meio
  let html = ''; let i = S.shown;
  while (i < stop) { const g = gk(F[i]); const items = []; while (i < stop && gk(F[i]) === g) items.push(F[i++]); html += groupHtml(items[0], g, items); }
  S.shown = stop; list.insertAdjacentHTML('beforeend', html);
}
function groupStats(g) {
  const all = S.filtered.filter((r) => gk(r) === g);
  const done = all.filter((r) => S.fills[r.k]).length; const sel = all.filter((r) => S.sel.has(r.k)).length;
  return { n: all.length, done, sel };
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

// clique / toque longo na lista
let pressT, pressed = false;
$('#list').addEventListener('pointerdown', (e) => {
  const it = e.target.closest('.it'); if (!it || S.selecting) return; pressed = false;
  pressT = setTimeout(() => { pressed = true; navigator.vibrate?.(30); setSelecting(true); toggleSel([it.dataset.k]); }, 520);
});
['pointerup', 'pointerleave', 'pointercancel', 'scroll'].forEach((ev) => $('#list').addEventListener(ev, () => clearTimeout(pressT), { passive: true }));
addEventListener('scroll', () => clearTimeout(pressT), { passive: true });
$('#list').addEventListener('click', (e) => {
  if (pressed) { pressed = false; return; }
  const gh = e.target.closest('.grp-h');
  if (gh) {
    if (!S.selecting) { setSelecting(true); }
    const ks = S.filtered.filter((r) => r.patio + '|' + r.ativo === gh.dataset.gh).map((r) => r.k);
    const allOn = ks.every((k) => S.sel.has(k));
    toggleSel(ks, !allOn); return;
  }
  const it = e.target.closest('.it'); if (!it) return;
  if (S.selecting) toggleSel([it.dataset.k]); else openFill([it.dataset.k]);
});
function toggleSel(ks, force) {
  for (const k of ks) { const on = force ?? !S.sel.has(k); if (on) S.sel.add(k); else S.sel.delete(k); }
  // atualiza sem redesenhar tudo
  for (const k of ks) {
    const el = $(`.it[data-k="${CSS.escape(k)}"]`); if (!el) continue;
    const on = S.sel.has(k); el.classList.toggle('sel', on); $('.cb', el).classList.toggle('on', on);
  }
  new Set(ks.map((k) => { const r = S.byK.get(k); return r && r.patio + '|' + r.ativo; })).forEach((g) => { const el = $(`.grp[data-g="${CSS.escape(g)}"]`); el && updGroupHead(el); });
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
$('#selAll').addEventListener('click', () => { const ks = S.filtered.map((r) => r.k); const all = ks.length && ks.every((k) => S.sel.has(k)); if (all) { S.sel.clear(); renderList(true); updSelbar(); } else { ks.forEach((k) => S.sel.add(k)); renderList(true); updSelbar(); } });
$('#selFill').addEventListener('click', () => S.sel.size && openFill([...S.sel]));

// ---------- preenchimento ----------
let FK = []; // OS abertas na folha
let colabTokens = [];
function fillColabList() {
  const names = new Set(S.users.filter((u) => u.active !== false).map((u) => u.name));
  ls.get('recentColab', []).forEach((n) => names.add(n));
  $('#colabList').innerHTML = [...names].map((n) => `<option value="${h(n)}">`).join('');
}
function renderTokens() {
  $$('#fColab .tok').forEach((t) => t.remove());
  $('#fColabIn').insertAdjacentHTML('beforebegin', colabTokens.map((n, i) => `<span class="tok">${h(n)}<button type="button" data-rm="${i}" aria-label="Remover ${h(n)}">${ICON('x')}</button></span>`).join(''));
  const names = new Set([S.me?.name, ...ls.get('recentColab', []), ...S.users.filter((u) => u.active !== false && u.equipe && u.equipe === segEquipe()).map((u) => u.name)]);
  const sugg = [...names].filter((n) => n && !colabTokens.includes(n)).slice(0, 8);
  $('#fColabSugg').innerHTML = sugg.map((n) => `<button type="button" class="chip" data-add="${h(n)}">${h(n)}</button>`).join('');
}
function addToken(v) {
  v.split(/[,;]+/).map((x) => x.trim()).filter(Boolean).forEach((n) => { if (!colabTokens.some((t) => t.toLowerCase() === n.toLowerCase())) colabTokens.push(n); });
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
function segEquipe() { return $('#fEquipe').dataset.value || ''; }
function renderEquipeChips(v) {
  const eqs = S.config.equipes?.length ? S.config.equipes : cfg.equipes || [];
  $('#fEquipe').innerHTML = eqs.map((e) => `<button type="button" class="chip" role="radio" data-v="${h(e)}" aria-checked="${e === v}" aria-pressed="${e === v}">${h(e)}</button>`).join('');
  $('#fEquipe').dataset.value = v || '';
}
$('#fEquipe').addEventListener('click', (e) => {
  const b = e.target.closest('[data-v]'); if (!b) return;
  const v = b.dataset.v === segEquipe() ? '' : b.dataset.v; renderEquipeChips(v); renderTokens();
});
['#fImed', '#fNec'].forEach((s) => bindSeg($(s)));
bindSeg($('#fOk'), () => $('#fErr').textContent = '');
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
    colab: (last.colab && last.colab.length ? last.colab : [S.me.name]).join(', '), equipe: last.equipe ?? S.me.equipe ?? '' };
  setSeg($('#fOk'), base.ok); $('#fDia').value = base.dia || todayISO(); $('#fIni').value = base.hIni || ''; $('#fFim').value = base.hFim || '';
  setSeg($('#fImed'), base.imediata || 'N'); setSeg($('#fNec'), base.necessidade || 'N');
  colabTokens = String(base.colab || '').split(',').map((x) => x.trim()).filter(Boolean);
  renderEquipeChips(base.equipe || ''); renderTokens(); $('#fColabIn').value = '';
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
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !$('#sheet').hidden && $('#dlg').hidden) closeSheet(); });

$('#fillForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  if ($('#fColabIn').value.trim()) addToken($('#fColabIn').value);
  const fill = {
    ok: segVal($('#fOk')), dia: $('#fDia').value, hIni: $('#fIni').value, hFim: $('#fFim').value,
    imediata: segVal($('#fImed')), necessidade: segVal($('#fNec')), colab: colabTokens.join(', '), equipe: segEquipe(),
    by: S.me.name, uid: S.me.uid, at: Date.now(),
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

// ---------- painel ----------
function renderDash() {
  if (!S.cycle) { $('#dash').innerHTML = '<div class="empty card wide"><b>Sem documento ativo</b></div>'; $('#pTitle').textContent = 'Painel'; $('#pSub').textContent = ''; return; }
  $('#pTitle').textContent = 'Painel • ' + (S.cycle.titulo || ''); $('#pSub').textContent = S.cycle.fileName || '';
  const c = counts(S.rows); const t = c.total || 1; const pct = Math.round((c.ok / t) * 100);
  const byP = new Map();
  for (const r of S.rows) { const o = byP.get(r.patio) || { n: 0, ok: 0, pe: 0 }; o.n++; const s = statusOf(r.k); if (s !== 'none') o[s]++; byP.set(r.patio, o); }
  const bars = [...byP.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([p, o]) =>
    `<button class="bar" data-p="${h(p)}"><b>${h(p)}</b><span class="track"><i class="x" style="width:${(o.ok / o.n) * 100}%"></i><i class="p" style="width:${(o.pe / o.n) * 100}%"></i></span><small>${o.ok + o.pe}/${o.n} • ${Math.round(((o.ok + o.pe) / o.n) * 100)}%</small></button>`).join('');
  const fl = Object.entries(S.fills).filter(([k]) => S.byK.has(k));
  const people = new Map(); for (const [, f] of fl) for (const n of String(f.colab || '').split(',').map((x) => x.trim()).filter(Boolean)) people.set(n, (people.get(n) || 0) + 1);
  const feed = fl.sort((a, b) => (b[1].at || 0) - (a[1].at || 0)).slice(0, 12).map(([k, f]) => { const r = S.byK.get(k); const s = f.ok === 'EXECUTADO' ? 'ok' : 'pe';
    return `<div><span class="st ${s}">${ICON(s === 'ok' ? 'check' : 'clock')}</span><p><b>${h(r.os)}</b> ${h(r.ativo)}<br><span class="muted small">${h(f.by || '')} • ${h(f.ok)}</span></p><small>${h(ago(f.at))}</small></div>`; }).join('');
  const imed = fl.filter(([, f]) => f.imediata === 'S').length; const nec = fl.filter(([, f]) => f.necessidade === 'S').length;
  $('#dash').innerHTML = `
    <div class="card pad"><div class="card-t"><b>Andamento</b><span class="muted small">${c.total} OS</span></div>
      <div class="ring-wrap"><div class="ring" style="--p:${(c.ok / t) * 100};--q:${((c.ok + c.pe) / t) * 100}"><span>${pct}%<small>executado</small></span></div>
      <div class="legend"><span><i style="background:var(--ok)"></i>Executadas <b>${c.ok}</b></span><span><i style="background:var(--pe)"></i>Pendentes <b>${c.pe}</b></span>
      <span><i style="background:color-mix(in srgb,var(--none) 40%,transparent)"></i>Sem preencher <b>${c.none}</b></span>
      <span class="muted small">Manutenção imediata: <b>${imed}</b> • Necessidade de manutenção: <b>${nec}</b></span></div></div></div>
    <div class="card pad"><div class="card-t"><b>Últimos preenchimentos</b></div><div class="feed">${feed || '<p class="muted">Nenhum preenchimento ainda.</p>'}</div></div>
    <div class="card pad wide"><div class="card-t"><b>Por pátio</b><span class="muted small">toque para abrir a lista</span></div><div class="bars">${bars}</div></div>
    <div class="card pad wide"><div class="card-t"><b>Por colaborador</b></div><div class="people">${[...people.entries()].sort((a, b) => b[1] - a[1]).map(([n, k]) => `<span class="person"><i>${h(initials(n))}</i>${h(n)} <small>${k}</small></span>`).join('') || '<p class="muted">—</p>'}</div></div>`;
}
$('#dash').addEventListener('click', (e) => { const b = e.target.closest('[data-p]'); if (!b) return; S.f.patio = b.dataset.p; S.f.status = 'todas'; ls.set('patio', S.f.patio); renderChips(); renderList(); location.hash = 'ciclica'; });

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

// ---------- admin: documento ----------
$('#adminChips').addEventListener('click', (e) => {
  const b = e.target.closest('[data-sub]'); if (!b) return;
  $$('#adminChips .chip').forEach((c) => c.setAttribute('aria-pressed', String(c === b)));
  $$('[data-subp]').forEach((p) => (p.hidden = p.dataset.subp !== b.dataset.sub));
  if (b.dataset.sub === 'users') S.store.listUsers().then((u) => { S.users = u; renderUsers(); });
});
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
  const cols = prim[0].header.labels;
  $('#preview').innerHTML = `<div class="card pad prev">
    <div class="card-t"><b>Conferir antes de publicar</b><span class="muted small">${h(S.upName)}</span></div>
    <label class="fl">Nome do documento<input id="pvTitle" value="${h(p.titulo || S.upName.replace(/\.xlsx$/i, ''))}"></label>
    <table><thead><tr><th>Aba</th><th>Uso</th><th>OS</th></tr></thead><tbody>${rows}</tbody></table>
    <p class="muted small">Abas “espelho” repetem as OS de outra aba (ex.: uma por pátio) e recebem os mesmos preenchimentos na exportação.</p>
    <div class="chips">${[...pats.entries()].sort().map(([k, n]) => `<span class="chip">${h(k)}<small>${n}</small></span>`).join('')}</div>
    <p class="muted small">Colunas reconhecidas: ${Object.values(cols).map(h).join(' • ')}</p>
    <label class="check"><input type="checkbox" id="pvActive" checked><span>Tornar este o documento ativo agora</span></label>
    <div class="actions"><button class="btn" id="pvCancel">Cancelar</button><button class="btn primary" id="pvPub">${ICON('upload')}<span>Publicar</span></button></div></div>`;
}
$('#preview').addEventListener('click', async (e) => {
  if (e.target.closest('#pvCancel')) { $('#preview').innerHTML = ''; S.parsed = null; return; }
  if (!e.target.closest('#pvPub')) return;
  const btn = e.target.closest('#pvPub');
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
  if (S.me?.role !== 'admin') return;
  try { S.cycles = await S.store.listCycles(); } catch (e) { $('#cycles').innerHTML = `<p class="err">${h(e.message)}</p>`; return; }
  const act = S.config.activeCycle;
  $('#cycles').innerHTML = S.cycles.map((c) => `<div class="card pad row-card${c.id === act ? ' active' : ''}" data-id="${h(c.id)}">
    <span class="rc-ic">${ICON('file')}</span>
    <span class="rc-tx"><b>${h(c.titulo)}${c.id === act ? '<span class="badge-a">ATIVO</span>' : ''}</b><small>${h(c.fileName || '')} • ${c.total || 0} OS • enviado ${h(ago(c.createdAt))}${c.createdBy ? ' por ' + h(c.createdBy) : ''}</small></span>
    <span class="rc-act">${c.id !== act ? `<button class="btn" data-ca="on">${ICON('star')}Ativar</button>` : ''}
      <button class="btn" data-ca="dl">${ICON('download')}Baixar preenchida</button>
      <button class="btn danger" data-ca="rm" aria-label="Excluir">${ICON('trash')}</button></span></div>`).join('') || '<p class="muted">Nenhum documento enviado ainda.</p>';
}
$('#cycles').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-ca]'); if (!b) return; const id = b.closest('[data-id]').dataset.id; const c = S.cycles.find((x) => x.id === id);
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
    try { await S.store.deleteCycle(id); if (S.config.activeCycle === id) await S.store.setConfig({ activeCycle: null }); toast('Documento excluído.'); renderCycles(); } catch (err) { toast(err.message, 5000); }
  }
});

// ---------- admin: usuários ----------
bindSeg($('#uRole')); setSeg($('#uRole'), 'func');
function equipeOptions(sel) { return ['<option value="">— sem equipe —</option>', ...(S.config.equipes || []).map((e) => `<option ${e === sel ? 'selected' : ''}>${h(e)}</option>`)].join(''); }
function renderUsers() {
  $('#uEquipe').innerHTML = equipeOptions($('#uEquipe').value);
  $('#users').innerHTML = S.users.sort((a, b) => (b.role === 'admin') - (a.role === 'admin') || a.name.localeCompare(b.name)).map((u) => `<div class="card pad row-card${u.active === false ? ' off' : ''}" data-uid="${h(u.uid)}">
    <span class="rc-ic" style="background:var(--yellow);color:#14202e;font-weight:800">${h(initials(u.name))}</span>
    <span class="rc-tx"><b>${h(u.name)}</b><small>@${h(u.username)} • ${u.role === 'admin' ? 'Administrador' : 'Funcionário'}${u.equipe ? ' • ' + h(u.equipe) : ''}${u.active === false ? ' • DESATIVADO' : ''}</small></span>
    <span class="rc-act"><button class="btn" data-ua="edit">${ICON('edit')}Editar</button><button class="btn" data-ua="pw">${ICON('key')}Nova senha</button>
    ${u.uid !== S.me.uid ? `<button class="btn" data-ua="toggle">${u.active === false ? 'Ativar' : 'Desativar'}</button><button class="btn danger" data-ua="rm" aria-label="Remover">${ICON('trash')}</button>` : ''}</span></div>`).join('');
}
$('#userForm').addEventListener('submit', async (e) => {
  e.preventDefault(); $('#userErr').textContent = ''; const btn = $('#userForm .btn.primary'); btn.disabled = true;
  try {
    await S.store.createUser({ name: $('#uName').value.trim(), username: cleanUser($('#uUser').value), password: $('#uPass').value, role: segVal($('#uRole')) || 'func', equipe: $('#uEquipe').value });
    toast(`Usuário ${cleanUser($('#uUser').value)} criado ✔`); $('#userForm').reset(); setSeg($('#uRole'), 'func');
    S.users = await S.store.listUsers(); renderUsers(); fillColabList();
  } catch (err) { $('#userErr').textContent = err.message; }
  btn.disabled = false;
});
$('#users').addEventListener('click', async (e) => {
  const b = e.target.closest('[data-ua]'); if (!b) return; const uid = b.closest('[data-uid]').dataset.uid; const u = S.users.find((x) => x.uid === uid);
  try {
    if (b.dataset.ua === 'edit') {
      const r = await dialog({
        title: 'Editar ' + u.name, yes: 'Salvar',
        html: `<label class="fl">Nome<input id="eName" value="${h(u.name)}" required></label><label class="fl">Equipe<select id="eEq">${equipeOptions(u.equipe)}</select></label>
          ${uid !== S.me.uid ? `<label class="fl">Perfil<select id="eRole"><option value="func" ${u.role !== 'admin' ? 'selected' : ''}>Funcionário</option><option value="admin" ${u.role === 'admin' ? 'selected' : ''}>Administrador</option></select></label>` : ''}`,
        validate: (bd) => ({ name: $('#eName', bd).value.trim() || u.name, equipe: $('#eEq', bd).value, ...($('#eRole', bd) ? { role: $('#eRole', bd).value } : {}) }),
      });
      if (r) { await S.store.updateUser(uid, r); toast('Usuário atualizado.'); }
    }
    if (b.dataset.ua === 'pw') {
      const r = await dialog({ title: 'Nova senha para ' + u.name, yes: 'Definir', html: `<label class="fl">Nova senha<input id="nPw" minlength="6" required autocomplete="new-password"></label><p class="muted small">Passe a nova senha para o funcionário. Ele pode trocar depois em “Trocar minha senha”.</p>`,
        validate: (bd) => { const v = $('#nPw', bd).value; return v.length >= 6 ? v : false; } });
      if (r) { await S.store.resetPassword(uid, r); toast('Senha redefinida.'); }
    }
    if (b.dataset.ua === 'toggle') { await S.store.updateUser(uid, { active: u.active === false }); toast(u.active === false ? 'Acesso ativado.' : 'Acesso desativado.'); }
    if (b.dataset.ua === 'rm') {
      const ok = await dialog({ title: 'Remover usuário?', html: `<p>Remove o acesso de <b>${h(u.name)}</b> (@${h(u.username)}). Os preenchimentos que ele fez continuam.</p>`, yes: 'Remover', danger: true });
      if (ok) { await S.store.removeUser(uid); toast('Usuário removido.'); }
    }
  } catch (err) { toast(err.message, 5000); }
  S.users = await S.store.listUsers(); renderUsers(); fillColabList();
});

// ---------- admin: equipes ----------
function renderEquipes() {
  $('#eqList').innerHTML = (S.config.equipes || []).map((e, i) => `<span class="chip">${h(e)}<button type="button" data-eq="${i}" aria-label="Remover ${h(e)}">${ICON('x')}</button></span>`).join('') || '<span class="muted">Nenhuma equipe.</span>';
  if (!$('#uEquipe').matches(':focus')) $('#uEquipe').innerHTML = equipeOptions($('#uEquipe').value);
}
$('#eqList').addEventListener('click', async (e) => { const b = e.target.closest('[data-eq]'); if (!b) return; const eq = [...(S.config.equipes || [])]; eq.splice(+b.dataset.eq, 1); await S.store.setConfig({ equipes: eq }); });
$('#eqForm').addEventListener('submit', async (e) => { e.preventDefault(); const v = $('#eqNew').value.trim().toUpperCase(); if (!v) return; const eq = [...new Set([...(S.config.equipes || []), v])]; await S.store.setConfig({ equipes: eq }); $('#eqNew').value = ''; });

// ---------- service worker (funciona sem internet depois do 1º acesso) ----------
if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});

boot();
