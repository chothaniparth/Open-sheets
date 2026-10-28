// A1-notation helpers and relative-reference shifting for formulas.

export const MAX_ROWS = 1_048_576
export const MAX_COLS = 16_384

export interface CellAddr {
  r: number
  c: number
}
export interface Range {
  r0: number
  c0: number
  r1: number
  c1: number
}

export function colName(c: number): string {
  let s = ''
  let n = c + 1
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

export function colIndex(name: string): number {
  let n = 0
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

export const cellName = (r: number, c: number) => `${colName(c)}${r + 1}`

export function rangeName(rg: Range): string {
  const a = cellName(rg.r0, rg.c0)
  return rg.r0 === rg.r1 && rg.c0 === rg.c1 ? a : `${a}:${cellName(rg.r1, rg.c1)}`
}

export function normRange(a: CellAddr, b: CellAddr): Range {
  return { r0: Math.min(a.r, b.r), c0: Math.min(a.c, b.c), r1: Math.max(a.r, b.r), c1: Math.max(a.c, b.c) }
}

export function parseCell(s: string): CellAddr | null {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/.exec(s.trim())
  if (!m) return null
  const r = Number(m[2]) - 1
  const c = colIndex(m[1])
  if (r < 0 || r >= MAX_ROWS || c < 0 || c >= MAX_COLS) return null
  return { r, c }
}

/** Split an optional sheet prefix: "'My Sheet'!A1:B2" → { sheet: "My Sheet", ref: "A1:B2" }. */
export function splitSheet(ref: string): { sheet?: string; ref: string } {
  const i = ref.lastIndexOf('!')
  if (i < 0) return { ref: ref.trim() }
  let sheet = ref.slice(0, i).trim()
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'")
  return { sheet, ref: ref.slice(i + 1).trim() }
}

/** Parse "A1", "A1:C9", "B:D" (whole columns) or "3:5" (whole rows). */
export function parseRange(s: string, used?: { rows: number; cols: number }): Range | null {
  const ref = s.trim().toUpperCase()
  const parts = ref.split(':')
  if (parts.length === 1) {
    const a = parseCell(parts[0])
    return a ? { r0: a.r, c0: a.c, r1: a.r, c1: a.c } : null
  }
  if (parts.length !== 2) return null
  const [p, q] = parts
  const a = parseCell(p)
  const b = parseCell(q)
  if (a && b) return normRange(a, b)
  const maxR = Math.max((used?.rows ?? 1) - 1, 0)
  const maxC = Math.max((used?.cols ?? 1) - 1, 0)
  if (/^\$?[A-Z]{1,3}$/.test(p) && /^\$?[A-Z]{1,3}$/.test(q)) {
    const c0 = colIndex(p.replace('$', ''))
    const c1 = colIndex(q.replace('$', ''))
    return { r0: 0, r1: maxR, c0: Math.min(c0, c1), c1: Math.max(c0, c1) }
  }
  if (/^\$?\d+$/.test(p) && /^\$?\d+$/.test(q)) {
    const r0 = Number(p.replace('$', '')) - 1
    const r1 = Number(q.replace('$', '')) - 1
    return { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: 0, c1: maxC }
  }
  return null
}

export const inRange = (rg: Range, r: number, c: number) => r >= rg.r0 && r <= rg.r1 && c >= rg.c0 && c <= rg.c1

export const rangesIntersect = (a: Range, b: Range) => a.r0 <= b.r1 && b.r0 <= a.r1 && a.c0 <= b.c1 && b.c0 <= a.c1

const REF_RE = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?![\w(])/g

/**
 * Shift relative cell references in a formula by (dr, dc), leaving $-anchored parts,
 * string literals, function names and sheet names untouched. References pushed off
 * the grid become #REF!.
 */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!formula.startsWith('=') || (dr === 0 && dc === 0)) return formula
  let out = ''
  let i = 0
  while (i < formula.length) {
    const ch = formula[i]
    if (ch === '"' || ch === "'") {
      // copy quoted string / quoted sheet name verbatim
      let j = i + 1
      while (j < formula.length) {
        if (formula[j] === ch) {
          if (formula[j + 1] === ch) j += 2
          else break
        } else j++
      }
      out += formula.slice(i, j + 1)
      i = j + 1
      continue
    }
    // find next quote to bound this plain segment
    let k = i
    while (k < formula.length && formula[k] !== '"' && formula[k] !== "'") k++
    const seg = formula.slice(i, k)
    out += seg.replace(REF_RE, (m, d1: string, col: string, d2: string, row: string, offset: number) => {
      const prev = seg[offset - 1]
      // part of a longer identifier (e.g. LOG10, named range) – leave it
      if (prev && /[\w.]/.test(prev)) return m
      let c = colIndex(col)
      let r = Number(row) - 1
      if (!d1) c += dc
      if (!d2) r += dr
      if (c < 0 || r < 0 || c >= MAX_COLS || r >= MAX_ROWS) return '#REF!'
      return `${d1}${colName(c)}${d2}${r + 1}`
    })
    i = k
  }
  return out
}

/** Find cell/range references in a formula (for highlighting while editing). */
export function findRefs(formula: string): { sheet?: string; range: Range }[] {
  const out: { sheet?: string; range: Range }[] = []
  const re = /(?:('(?:[^']|'')+'|[A-Za-z0-9_.]+)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7})(?::(\$?[A-Za-z]{1,3}\$?\d{1,7}))?(?![\w(])/g
  let m: RegExpExecArray | null
  const noStrings = formula.replace(/"(?:[^"]|"")*"/g, (s) => ' '.repeat(s.length))
  while ((m = re.exec(noStrings))) {
    const prev = noStrings[m.index - 1]
    if (prev && /[\w.]/.test(prev)) continue
    const a = parseCell(m[2])
    const b = m[3] ? parseCell(m[3]) : a
    if (!a || !b) continue
    const sheet = m[1] ? splitSheet(`${m[1]}!x`).sheet : undefined
    out.push({ sheet, range: normRange(a, b) })
  }
  return out
}

export function quoteSheet(name: string) {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`
}
