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
  mats: (h) => h.startsWith('MATRICULA'),
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
  let lastCol = 0;
  for (const c of Object.values(map)) lastCol = Math.max(lastCol, colToNum(c));
  // colunas extras acrescentadas no fim (só se a planilha ainda não tiver)
  const extras = [];
  if (opts.mats && !map.mats) extras.push(['mats', 'MATRÍCULAS']);
  if (opts.equipe && !map.equipe) extras.push(['equipe', 'EQUIPE']);
  for (const [f] of extras) map[f] = numToCol(++lastCol);
  // estilo de referência: célula de colaboradores (ou a última) da mesma linha
  const refCol = header.map.colab || numToCol(colToNum(Object.values(header.map).sort((x, y) => colToNum(y) - colToNum(x))[0]));

  const todo = new Map(byRow);
  if (extras.length) todo.set(header.row, '__header__');
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
    if (refCell) for (const [f] of extras) sHint[map[f]] = attr(refCell[0], 's');
    let values = {};
    if (fill === '__header__') for (const [f, label] of extras) values[map[f]] = { v: label };
    else {
      const put = (f, v, fmt) => { if (map[f]) values[map[f]] = { v, fmt }; };
      put('ok', fill.ok || '');
      put('dia', fill.dia ? dateSerial(fill.dia) : '', 'dd/mm/yyyy');
      put('hIni', fill.hIni ? timeSerial(fill.hIni) : '', 'hh:mm');
      put('hFim', fill.hFim ? timeSerial(fill.hFim) : '', 'hh:mm');
      put('imediata', fill.imediata || '');
      put('necessidade', fill.necessidade || '');
      put('colab', fill.colab || '');
      if (opts.mats || header.map.mats) put('mats', fill.mats || '');
      if (opts.equipe || header.map.equipe) put('equipe', fill.equipe || '');
    }
    out += xml.slice(last, m.index) + patchRow(rowXml, r, values, styles, sHint);
    last = m.index + m[0].length;
  }
  out += xml.slice(last);
  if (extras.length) {
    const lastL = numToCol(lastCol);
    out = out.replace(/<dimension ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/, (all, c1, r1, c2, r2) =>
      colToNum(c2) < lastCol ? `<dimension ref="${c1}${r1}:${lastL}${r2}"` : all);
    for (const [f] of extras) {
      const n = colToNum(map[f]); const w = f === 'mats' ? 24 : 14;
      if (/<cols>/.test(out) && !new RegExp(`<col\\b[^>]*min="${n}"`).test(out))
        out = out.replace('</cols>', `<col min="${n}" max="${n}" width="${w}" customWidth="1"/></cols>`);
    }
  }
  return out;
}

/**
 * Gera o .xlsx preenchido.
 * @param fillsByKey { [k]: {ok, dia, hIni, hFim, imediata, necessidade, colab, equipe} }
 * @param rows linhas do ciclo (k, os, aba, r)
 * @param opts { equipe: bool, mats: bool }  (acrescenta as colunas EQUIPE / MATRÍCULAS)
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
    if (!byRow.size && !opts.equipe && !opts.mats) continue;
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

// ---------- planilha simples (relatório) ----------
/** sheets: [{ name, rows: [[...valores]], widths: [n], head: [índices de linhas em negrito] }] */
export async function simpleXlsx(JSZip, sheets) {
  const zip = new JSZip();
  const ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'];
  sheets.forEach((_, i) => ct.push(`<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`));
  zip.file('[Content_Types].xml', ct.join('') + '</Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${NS_REL}"><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name.slice(0, 31))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets></workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  // 0 normal, 1 negrito com fundo, 2 título grande, 3 quebra de linha
  zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><b/><sz val="14"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF00335E"/></patternFill></fill></fills><borders count="2"><border/><border><bottom style="thin"><color rgb="FFCCCCCC"/></bottom></border></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="4"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"><alignment vertical="center" wrapText="1"/></xf><xf fontId="2" applyFont="1"/><xf borderId="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
  sheets.forEach((s, i) => {
    const cols = s.widths ? `<cols>${s.widths.map((w, j) => `<col min="${j + 1}" max="${j + 1}" width="${w}" customWidth="1"/>`).join('')}</cols>` : '';
    const rows = s.rows.map((row, r) => `<row r="${r + 1}">${row.map((v, c) => {
      const ref = numToCol(c + 1) + (r + 1);
      const st = (s.head || []).includes(r) ? 1 : (s.title || []).includes(r) ? 2 : r > 0 && v != null && v !== '' ? 3 : 0;
      if (v == null || v === '') return '';
      if (typeof v === 'number') return `<c r="${ref}" s="${st}"><v>${v}</v></c>`;
      return `<c r="${ref}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
    }).join('')}</row>`).join('');
    zip.file(`xl/worksheets/sheet${i + 1}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols}<sheetData>${rows}</sheetData></worksheet>`);
  });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
