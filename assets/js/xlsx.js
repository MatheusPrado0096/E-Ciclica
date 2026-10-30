// ===== E-Cíclica • leitura e preenchimento da planilha (.xlsx) =====
// Trabalha direto no XML do arquivo, então o documento exportado mantém
// formatação, filtros, abas, fórmulas e larguras exatamente como o original.

const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ---------- utilidades ----------
const ENT = { '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"', '&apos;': "'" };
export function unesc(s) {
  return s.replace(/&(amp|lt|gt|quot|apos);|&#(x?[0-9a-fA-F]+);/g, (m, n, c) =>
    n ? ENT[m] : String.fromCodePoint(c[0] === 'x' ? parseInt(c.slice(1), 16) : parseInt(c, 10)));
}
export function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    // remove caracteres de controle que o Excel não aceita
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}
export function norm(s) {
  return String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[°º]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}
export const colToNum = (c) => [...c].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
export const numToCol = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };

function attr(tag, name) {
  const m = tag.match(new RegExp('\\s' + name + '="([^"]*)"'));
  return m ? unesc(m[1]) : null;
}
function resolvePath(base, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/'); parts.pop();
  for (const p of target.split('/')) { if (p === '..') parts.pop(); else if (p !== '.') parts.push(p); }
  return parts.join('/');
}

// ---------- shared strings ----------
function parseSharedStrings(xml) {
  if (!xml) return [];
  const out = [];
  const re = /<si>([\s\S]*?)<\/si>|<si\/>/g; let m;
  while ((m = re.exec(xml))) {
    const inner = (m[1] || '').replace(/<rPh[\s\S]*?<\/rPh>/g, '');
    let t = ''; const tr = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g; let k;
    while ((k = tr.exec(inner))) t += k[1];
    out.push(unesc(t));
  }
  return out;
}

// ---------- células e linhas ----------
const ROW_RE = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g;
const CELL_RE = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;

function cellValue(attrs, inner, ss) {
  if (inner == null) return null;
  const t = attr(' ' + attrs, 't');
  if (t === 'inlineStr') {
    let s = ''; const tr = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g; let k;
    while ((k = tr.exec(inner))) s += k[1];
    return unesc(s);
  }
  const v = inner.match(/<v>([\s\S]*?)<\/v>/);
  if (!v) return null;
  if (t === 's') return ss[+v[1]] ?? '';
  if (t === 'str' || t === 'e') return unesc(v[1]);
  if (t === 'b') return v[1] === '1';
  const n = Number(v[1]);
  return Number.isFinite(n) ? n : unesc(v[1]);
}

function parseRows(xml, ss) {
  const rows = new Map();
  ROW_RE.lastIndex = 0; let m;
  while ((m = ROW_RE.exec(xml))) {
    const r = +attr(' ' + m[1], 'r');
    const cells = {}; let hasFormulaA = false;
    if (m[2]) {
      CELL_RE.lastIndex = 0; let c;
      while ((c = CELL_RE.exec(m[2]))) {
        const ref = attr(' ' + c[1], 'r'); if (!ref) continue;
        const col = ref.replace(/\d+/g, '');
        const v = cellValue(c[1], c[2], ss);
        if (v !== null && v !== '') cells[col] = v;
        if (col === 'A' && c[2] && /<f[\s>]/.test(c[2])) hasFormulaA = true;
      }
    }
    rows.set(r, { cells, hasFormulaA });
  }
  return rows;
}

// ---------- reconhecimento das colunas ----------
export const FIELDS = {
  patio: (h) => h === 'PATIO',
  ativo: (h) => h === 'ATIVO',
  os: (h) => /^OS\b/.test(h) || h === 'ORDEM' || h === 'ORDEM DE SERVICO',
  servico: (h) => h.startsWith('SERVICO'),
  desc: (h) => h.startsWith('DESCRICAO'),
  ok: (h) => h === 'OK' || h === 'STATUS',
  dia: (h) => h === 'DIA' || h === 'DATA',
  hIni: (h) => /^H(ORA)?\.? ?INICIO/.test(h),
  hFim: (h) => /^H(ORA)?\.? ?FIM/.test(h),
  imediata: (h) => h.includes('IMEDIATA'),
  necessidade: (h) => h.includes('NECESSIDADE'),
  colab: (h) => h.startsWith('COLABORADOR'),
  equipe: (h) => h === 'EQUIPE',
};
const REQUIRED = ['os', 'ok', 'dia'];

function detectHeader(rows) {
  for (let r = 1; r <= 6; r++) {
    const row = rows.get(r); if (!row) continue;
    const map = {}; const labels = {};
    for (const [col, v] of Object.entries(row.cells)) {
      const h = norm(v);
      for (const [f, test] of Object.entries(FIELDS)) {
        if (!map[f] && test(h)) { map[f] = col; labels[f] = String(v).trim(); break; }
      }
    }
    if (REQUIRED.every((f) => map[f])) return { row: r, map, labels };
  }
  return null;
}

async function readWorkbookParts(zip) {
  const wb = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const relMap = {};
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/?>/g)) relMap[attr(' ' + m[1], 'Id')] = attr(' ' + m[1], 'Target');
  const sheets = [];
  for (const m of wb.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const a = ' ' + m[1];
    const rid = attr(a, 'r:id') || (a.match(/\s[\w]+:id="([^"]*)"/) || [])[1];
    sheets.push({ name: attr(a, 'name'), state: attr(a, 'state') || 'visible', path: resolvePath('xl/workbook.xml', relMap[rid]) });
  }
  return { wb, sheets };
}

// ---------- leitura ----------
/**
 * Lê o arquivo e devolve as abas reconhecidas e as OS a preencher.
 * Abas "espelho" (ex.: uma aba por pátio que repete as OS da aba principal)
 * são detectadas e recebem os mesmos preenchimentos na exportação.
 */
export async function parseWorkbook(JSZip, data) {
  const zip = await JSZip.loadAsync(data);
  const ssFile = zip.file('xl/sharedStrings.xml');
  const ss = parseSharedStrings(ssFile ? await ssFile.async('string') : '');
  const { sheets } = await readWorkbookParts(zip);

  const found = [];
  for (const sh of sheets) {
    const f = zip.file(sh.path); if (!f) continue;
    const rows = parseRows(await f.async('string'), ss);
    const hdr = detectHeader(rows);
    if (!hdr) continue;
    const items = []; let formulas = 0;
    for (const [r, row] of rows) {
      if (r <= hdr.row) continue;
      const os = row.cells[hdr.map.os];
      if (os == null || String(os).trim() === '') continue;
      if (row.hasFormulaA) formulas++;
      const g = (k) => (hdr.map[k] && row.cells[hdr.map[k]] != null ? String(row.cells[hdr.map[k]]).trim() : '');
      items.push({ r, os: String(os).trim(), patio: g('patio'), ativo: g('ativo'), servico: g('servico'), desc: g('desc') });
    }
    if (!items.length) continue;
    found.push({ name: sh.name, path: sh.path, hidden: sh.state !== 'visible', header: hdr, items, formulas });
  }

  // aba principal x espelho: espelho = todas as OS já existem numa aba maior
  const bySize = [...found].sort((a, b) => b.items.length - a.items.length);
  const primaries = [];
  for (const s of bySize) {
    const set = new Set(s.items.map((i) => i.os));
    const parent = primaries.find((p) => p.items.length >= s.items.length && [...set].every((os) => p.osSet.has(os)));
    if (s.hidden) s.role = 'oculta';
    else if (parent) { s.role = 'espelho'; s.mirrorOf = parent.name; }
    else { s.role = 'principal'; s.osSet = set; primaries.push(s); }
  }
  for (const s of found) delete s.osSet;

  // título do mês: tenta tirar do cabeçalho "SERVIÇO OUTUBRO/2026 FPI"
  let titulo = '';
  const p0 = found.find((s) => s.role === 'principal');
  const lbl = p0?.header.labels.servico || p0?.header.labels.desc || '';
  const mm = norm(lbl).match(/(JANEIRO|FEVEREIRO|MARCO|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)\s*\/?\s*(\d{4})/);
  const m2 = norm(lbl).match(/(\d{1,2})\/(\d{4})/);
  const MESES = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  let ym = '';
  if (mm) {
    const i = ['JANEIRO', 'FEVEREIRO', 'MARCO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'].indexOf(mm[1]);
    titulo = MESES[i] + '/' + mm[2]; ym = mm[2] + '-' + String(i + 1).padStart(2, '0');
  } else if (m2) { titulo = MESES[+m2[1] - 1] + '/' + m2[2]; ym = m2[2] + '-' + m2[1].padStart(2, '0'); }

  return { sheets: found, titulo, ym };
}

/** Monta a lista de OS (com chave única) a partir das abas principais escolhidas. */
export function buildRows(parsed, chosen) {
  const use = parsed.sheets.filter((s) => chosen ? chosen.includes(s.name) : s.role === 'principal');
  const rows = []; const seen = new Map();
  for (const s of use) {
    for (const it of s.items) {
      const n = (seen.get(it.os) || 0) + 1; seen.set(it.os, n);
      const k = n === 1 ? it.os : it.os + '#' + n;
      rows.push({ k, aba: s.name, r: it.r, patio: it.patio || '—', ativo: it.ativo, os: it.os, servico: it.servico, desc: it.desc });
    }
  }
  return rows;
}

// ---------- exportação ----------
const EPOCH = Date.UTC(1899, 11, 30);
export function dateSerial(iso) { // "2026-10-30" -> 45960
  const [y, m, d] = iso.split('-').map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - EPOCH) / 86400000);
}
export function timeSerial(hhmm) { // "08:46" -> fração do dia
  const [h, m, s] = hhmm.split(':').map(Number);
  return ((h * 60 + m) * 60 + (s || 0)) / 86400;
}

class StylePool {
  constructor(xml) { this.xml = xml; this.cache = new Map(); this.fmtIds = {}; this.added = []; }
  init() {
    const xfs = this.xml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/);
    this.xfList = xfs ? [...xfs[1].matchAll(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)].map((m) => m[0]) : [];
    const ids = [...this.xml.matchAll(/<numFmt\b[^>]*numFmtId="(\d+)"/g)].map((m) => +m[1]);
    this.nextFmt = Math.max(163, ...ids) + 1;
  }
  fmt(code) {
    if (this.fmtIds[code]) return this.fmtIds[code];
    const exist = [...this.xml.matchAll(/<numFmt\b[^>]*?numFmtId="(\d+)"[^>]*?formatCode="([^"]*)"/g)].find((m) => unesc(m[2]) === code);
    if (exist) return (this.fmtIds[code] = +exist[1]);
    const id = this.nextFmt++; this.fmtIds[code] = id;
    const tag = `<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`;
    if (/<numFmts\b/.test(this.xml)) this.xml = this.xml.replace(/<numFmts\b([^>]*)>/, (m) => m + tag).replace(/(<numFmts\b[^>]*count=")(\d+)"/, (m, a, n) => a + (+n + 1) + '"');
    else this.xml = this.xml.replace(/(<styleSheet\b[^>]*>)/, `$1<numFmts count="1">${tag}</numFmts>`);
    return id;
  }
  /** estilo s com outro formato numérico */
  withFormat(s, code) {
    const key = (s ?? '0') + '|' + code;
    if (this.cache.has(key)) return this.cache.get(key);
    const id = this.fmt(code);
    let xf = this.xfList[+(s ?? 0)] || '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    xf = xf.replace(/\snumFmtId="\d+"/, '').replace(/\sapplyNumberFormat="\d"/, '').replace(/^<xf\b/, `<xf numFmtId="${id}" applyNumberFormat="1"`);
    const idx = this.xfList.length + this.added.length;
    this.added.push(xf); this.cache.set(key, idx);
    return idx;
  }
  out() {
    if (!this.added.length) return this.xml;
    return this.xml.replace(/<\/cellXfs>/, this.added.join('') + '</cellXfs>')
      .replace(/(<cellXfs\b[^>]*count=")(\d+)"/, (m, a, n) => a + (+n + this.added.length) + '"');
  }
}

function cellXml(ref, s, val, styles, fmt) {
  const sa = s != null ? ` s="${s}"` : '';
  if (val == null || val === '') return `<c r="${ref}"${sa}/>`;
  if (typeof val === 'number') {
    const s2 = fmt ? styles.withFormat(s, fmt) : s;
    return `<c r="${ref}"${s2 != null ? ` s="${s2}"` : ''}><v>${val}</v></c>`;
  }
  return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${esc(val)}</t></is></c>`;
}

function patchRow(rowXml, rowNum, values, styles, styleHint) {
  // values: { COL: {v, fmt} }
  const open = rowXml.match(/^<row\b[^>]*?(\/?)>/);
  let head = open[0].replace(/\s+spans="[^"]*"/, '');
  let inner = '';
  if (open[1] === '/') head = head.replace(/\s*\/>$/, '>');
  else inner = rowXml.slice(open[0].length, rowXml.length - '</row>'.length);
  const cells = []; CELL_RE.lastIndex = 0; let c;
  while ((c = CELL_RE.exec(inner))) {
    const ref = attr(' ' + c[1], 'r');
    cells.push({ col: ref.replace(/\d+/g, ''), xml: c[0], s: attr(' ' + c[1], 's') });
  }
  for (const [col, { v, fmt }] of Object.entries(values)) {
    const ref = col + rowNum;
    const i = cells.findIndex((x) => x.col === col);
    const s = i >= 0 ? cells[i].s : (styleHint[col] ?? null);
    const xml = cellXml(ref, s, v, styles, fmt);
    if (i >= 0) cells[i].xml = xml;
    else cells.push({ col, xml, s });
  }
  cells.sort((a, b) => colToNum(a.col) - colToNum(b.col));
  return head + cells.map((x) => x.xml).join('') + '</row>';
}

function patchSheet(xml, header, byRow, styles, opts) {
  const map = { ...header.map };
  const writeEquipe = opts.equipe && !map.equipe;
  let lastCol = 0;
  for (const c of Object.values(map)) lastCol = Math.max(lastCol, colToNum(c));
  if (writeEquipe) map.equipe = numToCol(lastCol + 1);
  // estilo de referência: célula de colaboradores (ou a última) da mesma linha
  const refCol = header.map.colab || numToCol(lastCol);

  const todo = new Map(byRow);
  if (writeEquipe) todo.set(header.row, '__header__');
  if (!todo.size) return xml;

  ROW_RE.lastIndex = 0;
  let out = ''; let last = 0; let m;
  while ((m = ROW_RE.exec(xml))) {
    const r = +attr(' ' + m[1], 'r');
    if (!todo.has(r)) continue;
    const fill = todo.get(r); todo.delete(r);
    const rowXml = m[0];
    const sHint = {};
    const refCell = rowXml.match(new RegExp(`<c\\b[^>]*\\sr="${refCol}${r}"[^>]*`));
    if (refCell) sHint[map.equipe] = attr(refCell[0], 's');
    let values;
    if (fill === '__header__') values = { [map.equipe]: { v: 'EQUIPE' } };
    else {
      values = {};
      const put = (f, v, fmt) => { if (map[f]) values[map[f]] = { v, fmt }; };
      put('ok', fill.ok || '');
      put('dia', fill.dia ? dateSerial(fill.dia) : '', 'dd/mm/yyyy');
      put('hIni', fill.hIni ? timeSerial(fill.hIni) : '', 'hh:mm');
      put('hFim', fill.hFim ? timeSerial(fill.hFim) : '', 'hh:mm');
      put('imediata', fill.imediata || '');
      put('necessidade', fill.necessidade || '');
      put('colab', fill.colab || '');
      if (opts.equipe || header.map.equipe) put('equipe', fill.equipe || '');
    }
    out += xml.slice(last, m.index) + patchRow(rowXml, r, values, styles, sHint);
    last = m.index + m[0].length;
  }
  out += xml.slice(last);
  if (writeEquipe) {
    const eqNum = colToNum(map.equipe);
    out = out.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/, (all, c1, r1, c2, r2) =>
      colToNum(c2) < eqNum ? `<dimension ref="${c1}${r1}:${map.equipe}${r2}"` : all);
    // largura da nova coluna
    if (/<cols>/.test(out) && !new RegExp(`<col\\b[^>]*min="${eqNum}"`).test(out))
      out = out.replace('</cols>', `<col min="${eqNum}" max="${eqNum}" width="16" customWidth="1"/></cols>`);
  }
  return out;
}

/**
 * Gera o .xlsx preenchido.
 * @param fillsByKey { [k]: {ok, dia, hIni, hFim, imediata, necessidade, colab, equipe} }
 * @param rows linhas do ciclo (k, os, aba, r)
 * @param opts { equipe: bool }  (equipe = acrescenta a coluna EQUIPE)
 */
export async function exportWorkbook(JSZip, data, rows, fillsByKey, opts = {}) {
  const zip = await JSZip.loadAsync(data);
  const parsed = await parseWorkbook(JSZip, data);
  const stylesPath = 'xl/styles.xml';
  const styles = new StylePool(await zip.file(stylesPath).async('string')); styles.init();

  // OS -> preenchimento (para abas espelho)
  const byOs = new Map(); const byAbaRow = new Map();
  for (const row of rows) {
    const f = fillsByKey[row.k]; if (!f) continue;
    byAbaRow.set(row.aba + '|' + row.r, f);
    if (!byOs.has(row.os)) byOs.set(row.os, f);
  }

  let total = 0;
  for (const sh of parsed.sheets) {
    if (sh.hidden) continue;
    const byRow = new Map();
    for (const it of sh.items) {
      const f = byAbaRow.get(sh.name + '|' + it.r) || (sh.role !== 'principal' ? byOs.get(it.os) : null);
      if (f) byRow.set(it.r, f);
    }
    if (!byRow.size && !opts.equipe) continue;
    const xml = await zip.file(sh.path).async('string');
    zip.file(sh.path, patchSheet(xml, sh.header, byRow, styles, opts));
    total += byRow.size;
  }
  zip.file(stylesPath, styles.out());

  // pede ao Excel para recalcular tudo ao abrir (abas com fórmulas)
  let wb = await zip.file('xl/workbook.xml').async('string');
  if (/<calcPr\b/.test(wb)) wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (m, a, sl) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${sl}>`);
  zip.file('xl/workbook.xml', wb);

  const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  return { bytes: out, cells: total };
}
