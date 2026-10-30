// ===== E-Cíclica • aplicativo =====
import { parseWorkbook, buildRows, exportWorkbook, norm } from './xlsx.js?v=6';
import { createLocalStore } from './store-local.js?v=6';

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
      const { createFirebaseStore } = await import('./store-firebase.js?v=6');
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
const TABS = ['ciclica', 'painel', 'relatorio', 'exportar', 'documento'];
function route() {
  let t = location.hash.slice(1);
  if (!TABS.includes(t)) t = 'ciclica';
  S.tab = t;
  $$('.panel').forEach((p) => (p.hidden = p.id !== t));
  $$('.tab').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.tab === t)));
  moveInd();
  if (t === 'painel') renderDash();
  if (t === 'exportar') renderExport();
  if (t === 'relatorio') renderReport();
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
  if (S.tab === 'relatorio') renderReport();
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
  $('#statusChips').innerHTML = sts.map(([v, l]) => chip('status', v, l, null, S.f.status === v)).join('') +
    `<button class="chip chip-seg" data-segf aria-pressed="${!!S.f.seg}">${ICON('shield')}Só segurança</button>`;
  $('#abaChips').hidden = ac.size < 2;
  if (S.f.aba && !ac.has(S.f.aba)) S.f.aba = '';
  $('#abaChips').innerHTML = chip('aba', '', 'Todas as abas', null, !S.f.aba) + [...ac.entries()].map(([a, n]) => chip('aba', a, a, n, S.f.aba === a)).join('');
}
$('#statusChips').addEventListener('click', (e) => { if (e.target.closest('[data-segf]')) { S.f.seg = !S.f.seg; renderChips(); renderList(); return; } const b = e.target.closest('[data-status]'); if (!b) return; S.f.status = b.dataset.status; renderChips(); renderList(); });
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
    if (S.f.seg && !isSeg(r)) return false;
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
    <span class="it-os">${hi(r.os)}${isSeg(r) ? '<span class="segb">SEG</span>' : ''}</span></button>`;
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

// ---------- preenchimento: colaboradores com matrícula ----------
let FK = [];
let colabTokens = []; // [{n: nome, m: matrícula}]
const keyN = (n) => norm(n);
function matOf(nome) { const u = S.users.find((x) => keyN(x.nome) === keyN(nome)); return u ? u.matricula : ''; }
function nameOf(mat) { const u = S.users.find((x) => String(x.matricula) === String(mat)); return u ? u.nome : ''; }
function recent() { return ls.get('recentColab2', []).filter((x) => x && x.n); }
// tokens de um preenchimento (aceita o formato antigo, só com nomes)
function tokensOf(f) {
  if (!f) return [];
  if (Array.isArray(f.colabs)) return f.colabs.map((c) => ({ n: c.n, m: c.m || matOf(c.n) }));
  const ms = splitColab(f.mats);
  return splitColab(f.colab).map((n, i) => ({ n, m: (ms[i] && ms[i] !== '-' ? ms[i] : '') || matOf(n) }));
}
function fillColabList() {
  const all = new Map();
  S.users.forEach((u) => all.set(keyN(u.nome), { n: u.nome, m: u.matricula }));
  recent().forEach((c) => all.has(keyN(c.n)) || all.set(keyN(c.n), c));
  $('#colabList').innerHTML = [...all.values()].map((c) => `<option value="${h(c.n)}">${c.m ? 'Matrícula ' + h(c.m) : ''}</option>`).join('');
}
function renderTokens() {
  $('#fColab').innerHTML = colabTokens.map((c, i) => `<span class="tok">${h(c.n)}${c.m ? `<em>${h(c.m)}</em>` : '<em class="nomat">sem matrícula</em>'}<button type="button" data-rm="${i}" aria-label="Remover ${h(c.n)}">${ICON('x')}</button></span>`).join('')
    || '<span class="muted small" style="padding:6px">Nenhum colaborador ainda.</span>';
  const eq = $('#fEquipe').dataset.value; const have = new Set(colabTokens.map((c) => keyN(c.n)));
  const sugg = new Map();
  const add = (c) => c?.n && !have.has(keyN(c.n)) && !sugg.has(keyN(c.n)) && sugg.set(keyN(c.n), c);
  add({ n: S.me?.nome, m: S.me?.matricula });
  recent().forEach(add);
  S.users.filter((u) => eq && u.equipe === eq).forEach((u) => add({ n: u.nome, m: u.matricula }));
  $('#fColabSugg').innerHTML = [...sugg.values()].slice(0, 8).map((c) => `<button type="button" class="chip" data-add="${h(c.n)}" data-m="${h(c.m || '')}">${h(c.n)}</button>`).join('');
}
function addColab(n, m) {
  n = String(n || '').trim().replace(/\s+/g, ' '); m = cleanMat(m);
  if (!n && m) n = nameOf(m);
  if (!n) return false;
  if (!m) m = matOf(n);
  const i = colabTokens.findIndex((t) => keyN(t.n) === keyN(n));
  if (i >= 0) { if (m) colabTokens[i].m = m; } else colabTokens.push({ n, m });
  $('#fColabIn').value = ''; $('#fMatIn').value = ''; renderTokens();
  return true;
}
function addFromInputs() { return addColab($('#fColabIn').value, $('#fMatIn').value); }
$('#fColab').addEventListener('click', (e) => { const b = e.target.closest('[data-rm]'); if (b) { colabTokens.splice(+b.dataset.rm, 1); renderTokens(); } });
// nome conhecido -> preenche a matrícula; matrícula conhecida -> preenche o nome
$('#fColabIn').addEventListener('input', (e) => { const m = matOf(e.target.value); if (m) $('#fMatIn').value = m; });
$('#fMatIn').addEventListener('input', (e) => { const n = nameOf(cleanMat(e.target.value)); if (n && !$('#fColabIn').value.trim()) $('#fColabIn').value = n; });
$('#fColabIn').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); if ($('#fMatIn').value || matOf(e.target.value)) addFromInputs(); else $('#fMatIn').focus(); } });
$('#fMatIn').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addFromInputs(); $('#fColabIn').focus(); } });
$('#fColabAdd').addEventListener('click', () => { if (!addFromInputs()) $('#fColabIn').focus(); });
$('#fColabSugg').addEventListener('click', (e) => { const b = e.target.closest('[data-add]'); if (b) addColab(b.dataset.add, b.dataset.m); });
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
    colabs: last.colabs?.length ? last.colabs : [{ n: S.me.nome, m: S.me.matricula }], equipe: last.equipe ?? S.me.equipe ?? '' };
  setSeg($('#fOk'), base.ok); $('#fDia').value = base.dia || todayISO(); $('#fIni').value = base.hIni || ''; $('#fFim').value = base.hFim || '';
  setSeg($('#fImed'), base.imediata || 'N'); setSeg($('#fNec'), base.necessidade || 'N');
  colabTokens = tokensOf(base).map((c) => ({ ...c }));
  eqChips($('#fEquipe'), base.equipe || ''); renderTokens(); $('#fColabIn').value = ''; $('#fMatIn').value = '';
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
  if ($('#fColabIn').value.trim() || $('#fMatIn').value.trim()) addFromInputs();
  const fill = {
    ok: segVal($('#fOk')), dia: $('#fDia').value, hIni: $('#fIni').value, hFim: $('#fFim').value,
    imediata: segVal($('#fImed')), necessidade: segVal($('#fNec')), colab: colabTokens.map((c) => c.n).join(', '), mats: colabTokens.map((c) => c.m || '-').join(', '), colabs: colabTokens.map((c) => ({ n: c.n, m: c.m || '' })), equipe: $('#fEquipe').dataset.value || '',
    by: S.me.nome, mat: S.me.matricula, at: Date.now(),
  };
  const miss = [];
  if (!fill.ok) miss.push('OK');
  if (fill.ok === 'EXECUTADO') { if (!fill.dia) miss.push('Dia'); if (!fill.hIni) miss.push('H início'); if (!fill.hFim) miss.push('H fim'); }
  if (!fill.colab) miss.push('Colaboradores');
  const semMat = colabTokens.filter((c) => !c.m).map((c) => c.n);
  if (!miss.length && semMat.length && !(await dialog({ title: 'Colaborador sem matrícula', html: `<p>${semMat.map(h).join(', ')} ${semMat.length > 1 ? 'estão' : 'está'} sem matrícula. Salvar assim mesmo?</p><p class="muted small">Para incluir, remova e adicione de novo com a matrícula.</p>`, yes: 'Salvar assim', no: 'Voltar' }))) return;
  if (miss.length) { $('#fErr').textContent = 'Preencha: ' + miss.join(', ') + '.'; return; }
  const btn = $('#fSave'); btn.disabled = true;
  try {
    await S.store.saveFills(S.cycle.id, FK.map((k) => ({ k, g: S.byK.get(k).g, fill })));
    ls.set('last', { colabs: colabTokens, equipe: fill.equipe, hIni: fill.hIni, day: todayISO() });
    const rc = new Map(); [...colabTokens, ...recent()].forEach((c) => rc.has(keyN(c.n)) || rc.set(keyN(c.n), c)); ls.set('recentColab2', [...rc.values()].slice(0, 20));
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
    if (D.col && !tokensOf(f).some((c) => c.n === D.col)) continue;
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
  for (const f of Object.values(S.fills)) { if (f.equipe) eqs.add(f.equipe); tokensOf(f).forEach((c) => cols.add(c.n)); }
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

// ---------- gráficos de pizza ----------
const isSeg = (r) => r.seg ?? (r.seg = /(^|_)SEG(_|$)/.test(norm(r.desc).replace(/\s+/g, '_')));
const pctTxt = (n, t) => (t ? Math.round((n / t) * 100) : 0) + '%';
// parts: [{v, c (cor css), l (rótulo)}]; hole: 0..1
function pie(parts, { size = 180, hole = 0.62, center = '', sub = '' } = {}) {
  const tot = parts.reduce((a, p) => a + p.v, 0); const r = 25; const sw = 50 * (1 - hole) / 1; const rr = 50 - sw / 2;
  let off = 0; const live = parts.filter((p) => p.v > 0); const gap = live.length > 1 ? 0.6 : 0;
  const segs = tot ? live.map((p) => {
    const len = (p.v / tot) * 100; const d = Math.max(0, len - gap);
    const c = `<circle cx="50" cy="50" r="${rr}" pathLength="100" fill="none" stroke="${p.c}" stroke-width="${sw}" stroke-dasharray="${d} ${100 - d}" stroke-dashoffset="${-off}" data-tip="<b>${h(p.l)}</b><br>${p.v} OS • ${pctTxt(p.v, tot)}"/>`;
    off += len; return c;
  }).join('') : `<circle cx="50" cy="50" r="${rr}" fill="none" stroke="var(--c-none)" stroke-width="${sw}"/>`;
  void r;
  return `<div class="pie" style="width:${size}px;height:${size}px"><svg viewBox="0 0 100 100" transform="rotate(-90)">${segs}</svg>${center !== '' ? `<div class="pie-c"><b>${center}</b>${sub ? `<small>${sub}</small>` : ''}</div>` : ''}</div>`;
}
function legend(parts, tot) {
  return `<ul class="pleg">${parts.map((p) => `<li${p.attr ? ' ' + p.attr : ''}><i style="background:${p.c}"></i><span>${h(p.l)}</span><b>${p.v}</b><small>${pctTxt(p.v, tot)}</small></li>`).join('')}</ul>`;
}
const CAT = ['var(--cat1)', 'var(--cat2)', 'var(--cat3)', 'var(--cat4)', 'var(--cat5)', 'var(--cat6)'];
function emptyCard(t, s) { return `<div class="empty"><b>${t}</b><span>${s}</span></div>`; }

function renderDash() {
  if (!S.cycle) { $('#dash').innerHTML = emptyCard('Sem documento ativo', 'Envie a planilha do mês na aba Documento.'); $('#dFilters').hidden = $('#dNav').hidden = true; $('#dFiltBtn').hidden = true; return; }
  $('#dFilters').hidden = $('#dNav').hidden = false; $('#dFiltBtn').hidden = false;
  $('#pTitle').textContent = 'Painel • ' + (S.cycle.titulo || '');
  renderFilters();
  $$('#dNav [data-v]').forEach((b) => b.setAttribute('aria-selected', String(b.dataset.v === S.d.view)));
  const data = dData(); tip.hidden = true;
  const v = S.d.view;
  $('#dash').innerHTML = v === 'geral' ? viewGeral(data) : v === 'patios' ? viewPatios(data) : v === 'equipes' ? viewEquipes(data)
    : v === 'pessoas' ? viewPessoas(data) : v === 'atencao' ? viewAtencao(data) : viewAtividade(data);
}
// partes "status" de um conjunto de linhas, respeitando os filtros de pessoa/período
function statusParts(rows, items) {
  const set = new Set(rows.map((r) => r.k));
  const its = items.filter((x) => set.has(x.r.k));
  const ok = its.filter((x) => x.f.ok === 'EXECUTADO').length; const pe = its.length - ok;
  const pf = personFilter();
  const rest = pf ? rows.length - its.length : rows.filter((r) => !S.fills[r.k]).length;
  return { ok, pe, rest, total: rows.length, parts: [
    { v: ok, c: 'var(--c-ok)', l: 'Executadas' }, { v: pe, c: 'var(--c-pe)', l: 'Pendentes' },
    { v: rest, c: 'var(--c-none)', l: pf ? 'Demais OS (fora do filtro)' : 'Sem preencher' }] };
}
function viewGeral({ rows, items }) {
  const all = statusParts(rows, items); const seg = statusParts(rows.filter(isSeg), items);
  const imed = items.filter((x) => x.f.imediata === 'S').length; const nec = items.filter((x) => x.f.necessidade === 'S').length;
  const card = (title, icon, st, note) => `<div class="card pad pie-card"><div class="card-t"><b>${ICON(icon)}${title}</b><span class="muted small">${st.total} OS</span></div>
    <div class="pie-row">${pie(st.parts, { center: pctTxt(st.ok, st.total), sub: 'realizadas' })}${legend(st.parts, st.total)}</div>${note ? `<p class="muted small">${note}</p>` : ''}</div>`;
  return `<div class="pie-grid">
    ${card('Todas as OS', 'list', all)}
    ${card('OS de segurança', 'shield', seg, 'OS com “SEG” na descrição (ex.: EE_INS_<b>SEG</b>_S_I…).')}
  </div>
  <button class="card pad alert-row" data-go="atencao"><span class="ar-ic">${ICON('alert')}</span><span class="ar-tx"><b>${imed + nec} OS pedem atenção</b><small>${imed} com manutenção imediata • ${nec} com necessidade de manutenção</small></span>${ICON('right')}</button>`;
}
function viewPatios({ rows, items }) {
  const by = new Map(); for (const r of rows) (by.get(r.patio) || by.set(r.patio, []).get(r.patio)).push(r);
  const list = [...by.entries()].map(([p, rs]) => [p, statusParts(rs, items), statusParts(rs.filter(isSeg), items)])
    .sort((a, b) => b[1].ok / b[1].total - a[1].ok / a[1].total || a[0].localeCompare(b[0]));
  return `<div class="card pad"><div class="card-t"><b>OS realizadas por pátio</b><span class="muted small">toque para abrir a lista</span></div>
    ${legend(list.length ? [{ v: list.reduce((a, x) => a + x[1].ok, 0), c: 'var(--c-ok)', l: 'Executadas' }, { v: list.reduce((a, x) => a + x[1].pe, 0), c: 'var(--c-pe)', l: 'Pendentes' }, { v: list.reduce((a, x) => a + x[1].rest, 0), c: 'var(--c-none)', l: personFilter() ? 'Demais OS' : 'Sem preencher' }] : [], rows.length).replace('class="pleg"', 'class="pleg inline"')}
    <div class="mini-grid">${list.map(([p, st, sg]) => `<button class="mini" data-p="${h(p)}">${pie(st.parts, { size: 92, hole: 0.6, center: pctTxt(st.ok, st.total) })}
      <b>${h(p)}</b><small>${st.ok}/${st.total} realizadas</small><small class="segs">${ICON('shield')}${sg.total ? `${sg.ok}/${sg.total} seg.` : 'sem OS seg.'}</small></button>`).join('')}</div></div>`;
}
function shareCard(title, map, filterAttr) {
  const list = [...map.entries()].sort((a, b) => b[1] - a[1]);
  if (!list.length) return emptyCard('Nenhum preenchimento', 'Nada encontrado com os filtros escolhidos.');
  const top = list.slice(0, 6); const other = list.slice(6).reduce((a, x) => a + x[1], 0);
  const parts = top.map(([l, v], i) => ({ v, c: CAT[i], l, attr: filterAttr && l !== 'Sem equipe' ? `data-${filterAttr}="${h(l)}"` : '' }));
  if (other) parts.push({ v: other, c: 'var(--cat-other)', l: `Outros (${list.length - 6})` });
  const tot = parts.reduce((a, p) => a + p.v, 0);
  return `<div class="card pad pie-card"><div class="card-t"><b>${title}</b><span class="muted small">toque no nome para filtrar</span></div>
    <div class="pie-row">${pie(parts, { size: 200, hole: 0.5, center: tot, sub: 'OS realizadas' })}${legend(parts, tot)}</div></div>`;
}
function viewEquipes({ items }) {
  const m = new Map(); for (const { f } of items) if (f.ok === 'EXECUTADO') { const k = f.equipe || 'Sem equipe'; m.set(k, (m.get(k) || 0) + 1); }
  return shareCard('OS realizadas por equipe', m, 'eq');
}
function viewPessoas({ items }) {
  const m = new Map(); for (const { f } of items) if (f.ok === 'EXECUTADO') for (const c of tokensOf(f)) { if (S.d.col && c.n !== S.d.col) continue; m.set(c.n, (m.get(c.n) || 0) + 1); }
  return shareCard('OS realizadas por colaborador', m, 'col');
}
function osRow(r, f, badge = '') {
  const s = f.ok === 'EXECUTADO' ? 'ok' : 'pe';
  return `<button class="feed-i" data-k="${h(r.k)}"><span class="st ${s}">${ICON(s === 'ok' ? 'check' : 'clock')}</span>
    <span class="fi-tx"><b>${h(r.ativo)} <em>OS ${h(r.os)}</em>${isSeg(r) ? ' <span class="segb">SEG</span>' : ''}</b><small>${h(r.servico)}</small><small>${brDate(f.dia)} • ${h(f.colab)}${f.equipe ? ' • ' + h(f.equipe) : ''}</small></span>${badge}</button>`;
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

// =====================================================================
// RELATÓRIO DIÁRIO
// =====================================================================
const R = { dia: todayISO(), quem: 'me' };
function isMine(f) {
  const me = S.me; if (!me) return false;
  return tokensOf(f).some((c) => (c.m && String(c.m) === String(me.matricula)) || keyN(c.n) === keyN(me.nome));
}
function repItems() {
  const out = [];
  for (const r of S.rows) {
    const f = S.fills[r.k]; if (!f || f.dia !== R.dia) continue;
    if (R.quem === 'me' && !isMine(f)) continue;
    if (R.quem.startsWith('eq:') && f.equipe !== R.quem.slice(3)) continue;
    if (R.quem.startsWith('col:') && !tokensOf(f).some((c) => c.n === R.quem.slice(4))) continue;
    out.push({ r, f });
  }
  return out.sort((a, b) => a.r.patio.localeCompare(b.r.patio) || (a.f.hIni || '').localeCompare(b.f.hIni || ''));
}
function quemLabel() {
  if (R.quem === 'me') return `${S.me.nome} (matrícula ${S.me.matricula})`;
  if (R.quem === 'all') return 'Todos';
  if (R.quem.startsWith('eq:')) return 'Equipe ' + R.quem.slice(3);
  const n = R.quem.slice(4); const m = matOf(n); return n + (m ? ` (matrícula ${m})` : '');
}
function repSummary(items) {
  const ok = items.filter((x) => x.f.ok === 'EXECUTADO'); const pe = items.length - ok.length;
  const segOk = ok.filter((x) => isSeg(x.r)).length; const segTot = items.filter((x) => isSeg(x.r)).length;
  const ini = items.map((x) => x.f.hIni).filter(Boolean).sort()[0] || ''; const fim = items.map((x) => x.f.hFim).filter(Boolean).sort().pop() || '';
  const byP = new Map(); for (const x of ok) byP.set(x.r.patio, (byP.get(x.r.patio) || 0) + 1);
  const imed = items.filter((x) => x.f.imediata === 'S').length; const nec = items.filter((x) => x.f.necessidade === 'S').length;
  return { ok: ok.length, pe, segOk, segTot, ini, fim, byP, imed, nec };
}
function renderReport() {
  if (!S.me) return;
  if (!$('#rDia').value) $('#rDia').value = R.dia;
  const cols = new Set(); const eqs = new Set();
  for (const f of Object.values(S.fills)) if (f.dia === R.dia) { tokensOf(f).forEach((c) => cols.add(c.n)); if (f.equipe) eqs.add(f.equipe); }
  if (S.me.equipe) eqs.add(S.me.equipe);
  $('#rQuem').innerHTML = `<option value="me">Eu — ${h(S.me.nome)}</option><option value="all">Todos</option>` +
    [...eqs].sort().map((e) => `<option value="eq:${h(e)}">Equipe ${h(e)}</option>`).join('') +
    [...cols].filter((c) => keyN(c) !== keyN(S.me.nome)).sort((a, b) => a.localeCompare(b)).map((c) => `<option value="col:${h(c)}">${h(c)}</option>`).join('');
  if (![...$('#rQuem').options].some((o) => o.value === R.quem)) R.quem = 'me';
  $('#rQuem').value = R.quem;
  if (!S.cycle) { $('#rep').innerHTML = emptyCard('Sem documento ativo', 'Envie a planilha do mês na aba Documento.'); $('#repAct').hidden = true; return; }
  const items = repItems(); const sm = repSummary(items);
  $('#repAct').hidden = !items.length;
  if (!items.length) { $('#rep').innerHTML = emptyCard('Nada preenchido neste dia', `Não há OS com DIA ${brDate(R.dia)} para ${h(quemLabel())}.`); return; }
  const tile = (n, l, cls = '') => `<div class="rt ${cls}"><b>${n}</b><span>${l}</span></div>`;
  const groups = new Map(); for (const x of items) (groups.get(x.r.patio) || groups.set(x.r.patio, []).get(x.r.patio)).push(x);
  $('#rep').innerHTML = `<article class="card pad report" id="repDoc">
    <header class="rep-h"><img src="assets/img/mrs-mark.png" alt="MRS" width="70"><div><span class="eyebrow">Relatório diário • Cíclica ${h(S.cycle.titulo || '')}</span>
      <h3>${brDate(R.dia)}</h3><p>${h(quemLabel())}${sm.ini ? ` • ${sm.ini} às ${sm.fim}` : ''}</p></div></header>
    <div class="rtiles">${tile(sm.ok, 'OS realizadas', 'ok')}${tile(sm.segOk + (sm.segTot > sm.segOk ? `<small>/${sm.segTot}</small>` : ''), 'de segurança', 'seg')}${tile(sm.pe, 'pendentes', 'pe')}${tile(sm.imed + sm.nec, 'pedem atenção', 'al')}</div>
    ${[...groups.entries()].map(([p, xs]) => `<section class="rgrp"><h4>${h(p)} <small>${xs.filter((x) => x.f.ok === 'EXECUTADO').length} realizadas de ${xs.length}</small></h4>
      <table class="rtab"><thead><tr><th>OS</th><th>Ativo / serviço</th><th>Status</th><th>Horário</th></tr></thead><tbody>
      ${xs.map(({ r, f }) => `<tr><td class="mono">${h(r.os)}${isSeg(r) ? '<br><span class="segb">SEG</span>' : ''}</td><td><b>${h(r.ativo)}</b><br><span>${h(r.servico)}</span>${f.imediata === 'S' || f.necessidade === 'S' ? `<br><span class="flag ${f.imediata === 'S' ? 'f1' : 'f2'}">${f.imediata === 'S' ? 'Manutenção imediata' : 'Necessidade de manutenção'}</span>` : ''}${R.quem !== 'me' ? `<br><small>${h(f.colab)}</small>` : ''}</td>
        <td><span class="pill ${f.ok === 'EXECUTADO' ? 'ok' : 'pe'}">${f.ok === 'EXECUTADO' ? 'Executado' : 'Pendente'}</span></td><td class="mono">${h(f.hIni || '')}${f.hFim ? '–' + h(f.hFim) : ''}</td></tr>`).join('')}
      </tbody></table></section>`).join('')}
    <footer class="rep-f">Gerado pelo E-Cíclica em ${new Date().toLocaleString('pt-BR')}</footer></article>`;
}
$('#rDia').addEventListener('change', (e) => { R.dia = e.target.value || todayISO(); renderReport(); });
$('#rQuem').addEventListener('change', (e) => { R.quem = e.target.value; renderReport(); });
function repFileBase() { return `Relatorio_${R.dia}_${(R.quem === 'me' ? S.me.nome : quemLabel()).replace(/\s*\(.*\)/, '').replace(/[^\wÀ-ú-]+/g, '_')}`; }
$('#rShare').addEventListener('click', async () => {
  const items = repItems(); const sm = repSummary(items);
  const txt = `*Relatório diário – ${brDate(R.dia)}*\n${quemLabel()}\nCíclica ${S.cycle.titulo || ''}\n\n✅ OS realizadas: ${sm.ok}\n🛡️ De segurança: ${sm.segOk}\n⏳ Pendentes: ${sm.pe}\n⚠️ Pedem atenção: ${sm.imed + sm.nec}` +
    (sm.ini ? `\n🕒 ${sm.ini} às ${sm.fim}` : '') + (sm.byP.size ? `\n\nPor pátio: ${[...sm.byP.entries()].map(([p, n]) => `${p} (${n})`).join(', ')}` : '');
  try { if (navigator.share) { await navigator.share({ text: txt }); return; } } catch (e) { if (e.name === 'AbortError') return; }
  try { await navigator.clipboard.writeText(txt); toast('Resumo copiado ✔'); } catch {}
  window.open('https://wa.me/?text=' + encodeURIComponent(txt), '_blank', 'noopener');
});
$('#rPdf').addEventListener('click', () => {
  const t = document.title; const th = document.documentElement.dataset.theme; document.title = repFileBase();
  document.documentElement.dataset.theme = 'light'; document.body.classList.add('print-rep');
  setTimeout(() => { window.print(); document.body.classList.remove('print-rep'); document.documentElement.dataset.theme = th; document.title = t; }, 50);
});
$('#rXlsx').addEventListener('click', async () => {
  const items = repItems(); const sm = repSummary(items);
  const rows = [[`Relatório diário — ${brDate(R.dia)}`], [quemLabel()], [`Cíclica ${S.cycle.titulo || ''}`], [],
    ['OS realizadas', sm.ok], ['De segurança', sm.segOk], ['Pendentes', sm.pe], ['Pedem atenção', sm.imed + sm.nec], ['Horário', sm.ini ? `${sm.ini} às ${sm.fim}` : ''], [],
    ['PÁTIO', 'ATIVO', 'OS', 'SERVIÇO', 'DESCRIÇÃO', 'SEGURANÇA', 'OK', 'H INÍCIO', 'H FIM', 'MANUT. IMEDIATA', 'NECESSIDADE', 'COLABORADORES', 'MATRÍCULAS', 'EQUIPE']];
  const headRow = rows.length - 1;
  for (const { r, f } of items) { const t = tokensOf(f); rows.push([r.patio, r.ativo, r.os, r.servico, r.desc, isSeg(r) ? 'SIM' : '', f.ok, f.hIni || '', f.hFim || '', f.imediata || '', f.necessidade || '', t.map((c) => c.n).join(', '), t.map((c) => c.m || '-').join(', '), f.equipe || '']); }
  const { simpleXlsx } = await import('./xlsx.js?v=6');
  const bytes = await simpleXlsx(JSZip, [{ name: 'Relatório ' + brDate(R.dia).replace(/\//g, '-'), rows, head: [headRow], title: [0], widths: [10, 26, 12, 50, 36, 11, 12, 9, 9, 10, 11, 30, 22, 9] }]);
  download(bytes, repFileBase() + '.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  toast('Relatório baixado ✔');
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
$('#exMats').checked = ls.get('exMats', true);
$('#exMats').addEventListener('change', (e) => ls.set('exMats', e.target.checked));
async function doExport(cycle, rows, fills, btn) {
  btn && (btn.disabled = true); toast('Gerando planilha…', 8000);
  try {
    const bytes = await S.store.getFile(cycle.id);
    if (!bytes) throw new Error('Arquivo original não encontrado.');
    const withMats = {};
    for (const [k, f] of Object.entries(fills)) { const t = tokensOf(f); withMats[k] = { ...f, colab: t.map((c) => c.n).join(', '), mats: t.map((c) => c.m || '-').join(', ') }; }
    const out = await exportWorkbook(JSZip, bytes, rows, withMats, { equipe: $('#exEquipe').checked, mats: $('#exMats').checked });
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
