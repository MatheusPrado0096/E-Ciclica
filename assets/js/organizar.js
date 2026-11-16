// ===== E-Cíclica • organização por pátio e abrigo (grupo) =====
// Regras (prompt de organização da cíclica):
//  - grupo/abrigo válido: 05, 10, 15, 20, 25, 30, 35, 40 ou 45
//  - comum: 2 dígitos logo após o pátio ("FCA 05 - 1T CDV", "FIP 05-1T CDV", "FRE 15 TE 05" → 15)
//  - "ABR NN" ("FCA TE07 - ABR 10" → 10)
//  - MCH "MCH N GG" → último número; "MCH GG-N" → número antes do hífen
//  - LPM → último número válido do nome
//  - demais (LDE010, TF005, MED01B, NO BREAK…) → sem grupo, no fim do pátio

export const GRUPOS = ['05', '10', '15', '20', '25', '30', '35', '40', '45'];
// cores claras por grupo (planilha e site)
export const GRUPO_COR = {
  '05': 'BDD7EE', // azul-claro
  '10': 'C6EFCE', // verde-claro
  '15': 'FFF2CC', // amarelo-claro
  '20': 'FCE4D6', // laranja-claro
  '25': 'E4DFEC', // lilás-claro
  '30': 'DDEBF7', // azul muito claro
  '35': 'E7E6E6', // cinza-claro
  '40': 'D6DCE4', // azul acinzentado
  '45': 'FFE1E1', // vermelho muito claro
};
const valido = (n) => { if (n == null) return null; const g = String(+n).padStart(2, '0'); return GRUPOS.includes(g) ? g : null; };
const up = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();

export function grupoOf(patio, ativo) {
  const P = up(patio); let a = up(ativo);
  if (P && a.startsWith(P)) a = a.slice(P.length).trim(); else return null; // pátio fora do padrão (ex.: MF-)
  let m;
  if ((m = a.match(/^MCH\s+(\d+)\s*-\s*\w+/))) return valido(m[1]);           // MCH GG-N
  if ((m = a.match(/^MCH\s+\d+\s+(\d+)\b/))) return valido(m[1]);             // MCH N GG
  if (/^LPM\b/.test(a)) {                                                      // LPM → último número válido
    const nums = a.match(/\d+/g) || [];
    for (let i = nums.length - 1; i >= 0; i--) { const g = valido(nums[i]); if (g && nums[i].length <= 2) return g; }
    return null;
  }
  if ((m = a.match(/\bABR\s*(\d{2})\b/))) return valido(m[1]);                // ... ABR 10
  if ((m = a.match(/^(\d{2})(?!\d)/))) return valido(m[1]);                   // 05 - 1T / 05-1T / 15 TE 05
  return null;
}

const coll = new Intl.Collator('pt-BR', { numeric: true, sensitivity: 'base' });
/** chave de ordenação: pátio → grupo (05…45, sem grupo no fim) → ativo → OS */
export function compareOrg(a, b) {
  const p = coll.compare(a.patio || '', b.patio || ''); if (p) return p;
  const ga = a.grupo ? GRUPOS.indexOf(a.grupo) : 99, gb = b.grupo ? GRUPOS.indexOf(b.grupo) : 99;
  if (ga !== gb) return ga - gb;
  const at = coll.compare(a.ativo || '', b.ativo || ''); if (at) return at;
  return coll.compare(String(a.os || ''), String(b.os || ''));
}
export function organizar(rows) {
  for (const r of rows) if (r.grupo === undefined) r.grupo = grupoOf(r.patio, r.ativo);
  return [...rows].sort(compareOrg);
}
/** pátio no padrão de 3 letras (ex.: FCA). "MF-" e afins são fora do padrão. */
export const patioPadrao = (p) => /^[A-Z]{3}$/.test(up(p));
