import { DetailedCellError, HyperFormula, type CellValue, type ConfigParams, type RawCellContent } from 'hyperformula'
import { MAX_COLS, MAX_ROWS, inRange, parseRange, rangesIntersect, shiftFormula, splitSheet, type Range } from './addr'
import { editTextFor, formatValue, isDateFormat, parseInput } from './format'

export interface Borders {
  top?: string
  right?: string
  bottom?: string
  left?: string
}

export interface CellStyle {
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
  font?: string
  size?: number
  color?: string
  fill?: string
  align?: 'left' | 'center' | 'right'
  valign?: 'top' | 'middle' | 'bottom'
  wrap?: boolean
  fmt?: string
  border?: Borders
}

export type CondType = 'gt' | 'lt' | 'between' | 'eq' | 'contains' | 'blank' | 'duplicate' | 'scale'

export interface CondRule {
  id: string
  range: string
  type: CondType
  a?: string
  b?: string
  fill?: string
  color?: string
  bold?: boolean
  /** Two-color scale: [min color, max color]. */
  scale?: [string, string]
}

export interface FilterState {
  range: Range
  /** column index → allowed display values (absent = no filter on that column) */
  criteria: Record<number, string[]>
}

export interface GridSheet {
  kind: 'grid'
  id: string
  name: string
  colWidths: Record<number, number>
  rowHeights: Record<number, number>
  hiddenRows: Record<number, true>
  styles: Record<string, CellStyle>
  merges: Range[]
  freeze: { rows: number; cols: number }
  cond: CondRule[]
  filter?: FilterState
  tabColor?: string
}

export type ChartType = 'column' | 'stackedColumn' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'funnel'

export interface ChartSpec {
  type: ChartType
  range: string
  seriesBy: 'columns' | 'rows'
  headers: boolean
  legend: boolean
  smooth?: boolean
}

export interface Widget {
  id: string
  type: 'chart' | 'kpi' | 'text' | 'table'
  x: number
  y: number
  w: number
  h: number
  title?: string
  chart?: ChartSpec
  kpi?: { formula: string; fmt?: string; compare?: string; compareLabel?: string }
  text?: string
  table?: { range: string }
}

export interface Dashboard {
  kind: 'dashboard'
  id: string
  name: string
  widgets: Widget[]
}

export type Tab = GridSheet | Dashboard

export interface SheetData {
  name: string
  cells: RawCellContent[][]
  meta?: Partial<Omit<GridSheet, 'kind' | 'id' | 'name'>>
}

export interface WorkbookData {
  sheets: SheetData[]
  dashboards?: Dashboard[]
  /** Optional tab order by name (sheets + dashboards). */
  order?: string[]
}

export const DEFAULT_COL_WIDTH = 100
export const DEFAULT_ROW_HEIGHT = 24

const HF_CONFIG: Partial<ConfigParams> = {
  licenseKey: 'gpl-v3',
  maxRows: MAX_ROWS,
  maxColumns: MAX_COLS,
  evaluateNullToZero: true,
  currencySymbol: ['$', '€', '£', '¥', '₹'],
  dateFormats: ['MM/DD/YYYY', 'MM/DD/YY', 'YYYY-MM-DD', 'YYYY/MM/DD'],
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  undoLimit: 1000,
  useArrayArithmetic: true
}

let idSeq = 0
export const newId = (p: string) => `${p}${Date.now().toString(36)}${(idSeq++).toString(36)}`

export function blankGrid(name: string): GridSheet {
  return { kind: 'grid', id: newId('s'), name, colWidths: {}, rowHeights: {}, hiddenRows: {}, styles: {}, merges: [], freeze: { rows: 0, cols: 0 }, cond: [] }
}

export const styleKey = (r: number, c: number) => `${r}:${c}`

function cloneTab(t: Tab): Tab {
  if (t.kind === 'dashboard') return { ...t, widgets: t.widgets.map((w) => ({ ...w })) }
  return {
    ...t,
    colWidths: { ...t.colWidths },
    rowHeights: { ...t.rowHeights },
    hiddenRows: { ...t.hiddenRows },
    styles: { ...t.styles },
    merges: [...t.merges],
    freeze: { ...t.freeze },
    cond: [...t.cond],
    filter: t.filter && { ...t.filter, criteria: { ...t.filter.criteria } }
  }
}

interface UndoEntry {
  label: string
  hf: boolean
  before: Tab[]
  after?: Tab[]
}

export type Transact = (label: string, fn: () => void, opts?: { hf?: boolean }) => void

/**
 * The document model: HyperFormula holds cell contents and computes formulas; this class
 * holds everything else (formatting, layout, dashboards) and a unified undo history.
 */
export class Workbook {
  hf: HyperFormula
  tabs: Tab[]
  version = 0
  dirty = false
  private undoStack: UndoEntry[] = []
  private redoStack: UndoEntry[] = []
  private listeners = new Set<() => void>()
  private condCache = new Map<string, { version: number; range: Range | null; min: number; max: number; dupes?: Set<string> }>()

  private constructor(hf: HyperFormula, tabs: Tab[]) {
    this.hf = hf
    this.tabs = tabs
  }

  static create(): Workbook {
    return Workbook.fromData({ sheets: [{ name: 'Sheet1', cells: [] }] })
  }

  static fromData(data: WorkbookData): Workbook {
    const sheets: Record<string, RawCellContent[][]> = {}
    const tabs: Tab[] = []
    for (const s of data.sheets) {
      sheets[s.name] = s.cells
      tabs.push({ ...blankGrid(s.name), ...(s.meta ?? {}) })
    }
    if (!tabs.length) {
      sheets.Sheet1 = []
      tabs.push(blankGrid('Sheet1'))
    }
    for (const d of data.dashboards ?? []) tabs.push(d)
    if (data.order) {
      const pos = new Map(data.order.map((n, i) => [n, i]))
      tabs.sort((a, b) => (pos.get(a.name) ?? 1e9) - (pos.get(b.name) ?? 1e9))
    }
    const hf = HyperFormula.buildFromSheets(sheets, HF_CONFIG)
    hf.clearUndoStack()
    return new Workbook(hf, tabs)
  }

  destroy() {
    this.hf.destroy()
  }

  // ---------------- subscription ----------------

  subscribe = (fn: () => void) => {
    this.listeners.add(fn)
    return () => {
      this.listeners.delete(fn)
    }
  }

  getVersion = () => this.version

  emit() {
    this.version++
    for (const fn of this.listeners) fn()
  }

  // ---------------- lookup ----------------

  tab(id: string): Tab | undefined {
    return this.tabs.find((t) => t.id === id)
  }

  grids(): GridSheet[] {
    return this.tabs.filter((t): t is GridSheet => t.kind === 'grid')
  }

  findTab(name: string): Tab | undefined {
    const n = name.toLowerCase()
    return this.tabs.find((t) => t.name.toLowerCase() === n)
  }

  sid(tab: GridSheet): number {
    const id = this.hf.getSheetId(tab.name)
    if (id === undefined) throw new Error(`Sheet "${tab.name}" not found`)
    return id
  }

  /** Resolve "Sheet!A1:B2" (or a bare range on a default sheet). */
  resolveRef(ref: string, defaultSheet?: GridSheet): { tab: GridSheet; range: Range } | null {
    const { sheet, ref: r } = splitSheet(ref)
    const tab = sheet ? this.findTab(sheet) : defaultSheet
    if (!tab || tab.kind !== 'grid') return null
    const range = parseRange(r, this.size(tab))
    return range ? { tab, range } : null
  }

  size(tab: GridSheet): { rows: number; cols: number } {
    const d = this.hf.getSheetDimensions(this.sid(tab))
    return { rows: d.height, cols: d.width }
  }

  value(tab: GridSheet, r: number, c: number): CellValue {
    return this.hf.getCellValue({ sheet: this.sid(tab), row: r, col: c })
  }

  raw(tab: GridSheet, r: number, c: number): RawCellContent {
    return this.hf.getCellSerialized({ sheet: this.sid(tab), row: r, col: c })
  }

  isFormula(tab: GridSheet, r: number, c: number): boolean {
    return this.hf.doesCellHaveFormula({ sheet: this.sid(tab), row: r, col: c })
  }

  style(tab: GridSheet, r: number, c: number): CellStyle | undefined {
    return tab.styles[styleKey(r, c)]
  }

  /** Effective number format: explicit style or one implied by the formula engine (dates, %…). */
  numberFormat(tab: GridSheet, r: number, c: number): string | undefined {
    const fmt = tab.styles[styleKey(r, c)]?.fmt
    if (fmt) return fmt
    const addr = { sheet: this.sid(tab), row: r, col: c }
    const t = this.hf.getCellValueDetailedType(addr)
    switch (t) {
      case 'NUMBER_DATE':
        return 'yyyy-mm-dd'
      case 'NUMBER_TIME':
        return 'h:mm:ss'
      case 'NUMBER_DATETIME':
        return 'yyyy-mm-dd h:mm'
      case 'NUMBER_PERCENT':
        return '0%'
      case 'NUMBER_CURRENCY':
        return `${this.hf.getCellValueFormat(addr) ?? '$'}#,##0.00`
      default:
        return undefined
    }
  }

  display(tab: GridSheet, r: number, c: number): string {
    return formatValue(this.value(tab, r, c), this.numberFormat(tab, r, c))
  }

  /** Text placed in the editor when the user starts editing a cell. */
  editText(tab: GridSheet, r: number, c: number): string {
    const raw = this.raw(tab, r, c)
    return editTextFor(raw, tab.styles[styleKey(r, c)]?.fmt)
  }

  rangeValues(tab: GridSheet, rg: Range): CellValue[][] {
    return this.hf.getRangeValues({ start: { sheet: this.sid(tab), row: rg.r0, col: rg.c0 }, end: { sheet: this.sid(tab), row: rg.r1, col: rg.c1 } })
  }

  rangeRaw(tab: GridSheet, rg: Range): RawCellContent[][] {
    return this.hf.getRangeSerialized({ start: { sheet: this.sid(tab), row: rg.r0, col: rg.c0 }, end: { sheet: this.sid(tab), row: rg.r1, col: rg.c1 } })
  }

  /** Values of a "Sheet!A1:B9" reference (used by charts and tables). */
  refValues(ref: string): { values: CellValue[][]; tab: GridSheet; range: Range } | null {
    const res = this.resolveRef(ref, this.grids()[0])
    if (!res) return null
    const rg = clampRange(res.range, this.size(res.tab))
    return { values: rg ? this.rangeValues(res.tab, rg) : [], tab: res.tab, range: res.range }
  }

  /** Evaluate a free-standing formula (KPI cards). */
  evaluate(formula: string): CellValue {
    const sheet = this.sid(this.grids()[0])
    const f = formula.startsWith('=') ? formula : `=${formula}`
    try {
      const v = this.hf.calculateFormula(f, sheet)
      return Array.isArray(v) ? (v[0]?.[0] ?? null) : v
    } catch (err) {
      return `#ERROR: ${(err as Error).message}`
    }
  }

  /** Expand a cell to the surrounding block of non-empty cells (Excel's "current region"). */
  currentRegion(tab: GridSheet, r: number, c: number): Range {
    const size = this.size(tab)
    const filled = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < size.rows && cc < size.cols && this.raw(tab, rr, cc) !== null
    const rowHas = (rr: number, c0: number, c1: number) => {
      for (let cc = c0; cc <= c1; cc++) if (filled(rr, cc)) return true
      return false
    }
    const colHas = (cc: number, r0: number, r1: number) => {
      for (let rr = r0; rr <= r1; rr++) if (filled(rr, cc)) return true
      return false
    }
    const rg = { r0: r, c0: c, r1: r, c1: c }
    let changed = true
    while (changed) {
      changed = false
      if (rg.r0 > 0 && rowHas(rg.r0 - 1, rg.c0 - 1, rg.c1 + 1)) (rg.r0--, (changed = true))
      if (rg.r1 < size.rows - 1 && rowHas(rg.r1 + 1, rg.c0 - 1, rg.c1 + 1)) (rg.r1++, (changed = true))
      if (rg.c0 > 0 && colHas(rg.c0 - 1, rg.r0 - 1, rg.r1 + 1)) (rg.c0--, (changed = true))
      if (rg.c1 < size.cols - 1 && colHas(rg.c1 + 1, rg.r0 - 1, rg.r1 + 1)) (rg.c1++, (changed = true))
    }
    return rg
  }

  /** Heuristic: first row is a header if it is all text and the next row is not. */
  looksLikeHeader(tab: GridSheet, rg: Range): boolean {
    if (rg.r1 <= rg.r0) return false
    const first = this.rangeValues(tab, { ...rg, r1: rg.r0 })[0]
    const second = this.rangeValues(tab, { ...rg, r0: rg.r0 + 1, r1: rg.r0 + 1 })[0]
    return first.every((v) => typeof v === 'string' && v !== '') && second.some((v) => typeof v !== 'string')
  }

  mergeAt(tab: GridSheet, r: number, c: number): Range | undefined {
    return tab.merges.find((m) => inRange(m, r, c))
  }

  // ---------------- conditional formatting ----------------

  condStyle(tab: GridSheet, r: number, c: number): Partial<CellStyle> | undefined {
    if (!tab.cond.length) return undefined
    let out: Partial<CellStyle> | undefined
    for (const rule of tab.cond) {
      let cache = this.condCache.get(rule.id)
      if (!cache || cache.version !== this.version) {
        const range = parseRange(rule.range, this.size(tab))
        cache = { version: this.version, range, min: Infinity, max: -Infinity }
        if (range && (rule.type === 'scale' || rule.type === 'duplicate')) {
          const rg = clampRange(range, this.size(tab))
          const counts = new Map<string, number>()
          for (const row of rg ? this.rangeValues(tab, rg) : []) {
            for (const v of row) {
              if (typeof v === 'number') {
                cache.min = Math.min(cache.min, v)
                cache.max = Math.max(cache.max, v)
              }
              if (v !== null && v !== '') counts.set(String(v), (counts.get(String(v)) ?? 0) + 1)
            }
          }
          cache.dupes = new Set([...counts].filter(([, n]) => n > 1).map(([k]) => k))
        }
        this.condCache.set(rule.id, cache)
      }
      if (!cache.range || !inRange(cache.range, r, c)) continue
      const v = this.value(tab, r, c)
      let hit = false
      const num = typeof v === 'number' ? v : NaN
      const a = Number(rule.a)
      const b = Number(rule.b)
      switch (rule.type) {
        case 'gt':
          hit = num > a
          break
        case 'lt':
          hit = num < a
          break
        case 'between':
          hit = num >= Math.min(a, b) && num <= Math.max(a, b)
          break
        case 'eq':
          hit = Number.isNaN(a) ? String(v ?? '').toLowerCase() === String(rule.a ?? '').toLowerCase() : num === a
          break
        case 'contains':
          hit = String(v ?? '').toLowerCase().includes(String(rule.a ?? '').toLowerCase())
          break
        case 'blank':
          hit = v === null || v === ''
          break
        case 'duplicate':
          hit = v !== null && v !== '' && !!cache.dupes?.has(String(v))
          break
        case 'scale':
          if (!Number.isNaN(num) && cache.max > cache.min && rule.scale) {
            out = { ...out, fill: mixColor(rule.scale[0], rule.scale[1], (num - cache.min) / (cache.max - cache.min)) }
          }
          continue
      }
      if (hit) out = { ...out, ...(rule.fill ? { fill: rule.fill } : {}), ...(rule.color ? { color: rule.color } : {}), ...(rule.bold ? { bold: true } : {}) }
    }
    return out
  }

  // ---------------- transactions & undo ----------------

  /**
   * Run a mutation as one undoable step. Every HyperFormula change must happen inside a
   * transaction with `hf: true` (it becomes exactly one HyperFormula batch / undo entry).
   * No HyperFormula values can be read inside `fn` because evaluation is suspended.
   */
  transact: Transact = (label, fn, opts = {}) => {
    const before = this.tabs.map(cloneTab)
    if (opts.hf) {
      this.hf.batch(fn)
    } else {
      fn()
    }
    this.undoStack.push({ label, hf: !!opts.hf, before })
    if (this.undoStack.length > 300) this.undoStack.shift()
    this.redoStack = []
    this.dirty = true
    this.emit()
  }

  canUndo() {
    return this.undoStack.length > 0
  }
  canRedo() {
    return this.redoStack.length > 0
  }
  undoLabel() {
    return this.undoStack.at(-1)?.label
  }

  undo() {
    const e = this.undoStack.pop()
    if (!e) return
    e.after = this.tabs.map(cloneTab)
    if (e.hf && this.hf.isThereSomethingToUndo()) this.hf.undo()
    this.tabs = e.before
    this.redoStack.push(e)
    this.dirty = true
    this.emit()
  }

  redo() {
    const e = this.redoStack.pop()
    if (!e) return
    if (e.hf && this.hf.isThereSomethingToRedo()) this.hf.redo()
    this.tabs = e.after ?? this.tabs
    this.undoStack.push({ ...e, after: undefined })
    this.dirty = true
    this.emit()
  }

  // ---------------- cell contents ----------------

  /** Write user-typed text to a cell (with Excel-like auto formatting). */
  setInput(tab: GridSheet, r: number, c: number, text: string) {
    this.writeTexts(tab, r, c, [[text]], 'Typing')
  }

  /** Write a block of user-style text (from paste, AI, etc.). */
  writeTexts(tab: GridSheet, r: number, c: number, rows: string[][], label = 'Edit') {
    const raws: RawCellContent[][] = []
    const fmts: { r: number; c: number; fmt: string }[] = []
    rows.forEach((row, i) => {
      raws.push(
        row.map((text, j) => {
          const p = parseInput(text ?? '')
          const existing = tab.styles[styleKey(r + i, c + j)]?.fmt
          if (p.fmt && (!existing || existing === 'General')) fmts.push({ r: r + i, c: c + j, fmt: p.fmt })
          return p.raw
        })
      )
    })
    this.transact(
      label,
      () => {
        this.hf.setCellContents({ sheet: this.sid(tab), row: r, col: c }, raws)
        for (const f of fmts) this.patchStyle(tab, f.r, f.c, { fmt: f.fmt })
      },
      { hf: true }
    )
  }

  /** Write raw engine contents (formulas/values), optionally with styles. */
  writeRaw(tab: GridSheet, r: number, c: number, raws: RawCellContent[][], styles?: (CellStyle | undefined)[][], label = 'Paste') {
    this.transact(
      label,
      () => {
        this.hf.setCellContents({ sheet: this.sid(tab), row: r, col: c }, raws)
        if (styles) {
          styles.forEach((row, i) =>
            row.forEach((st, j) => {
              const k = styleKey(r + i, c + j)
              if (st) tab.styles[k] = st
              else delete tab.styles[k]
            })
          )
        }
      },
      { hf: true }
    )
  }

  clearContents(tab: GridSheet, rg: Range) {
    const size = this.size(tab)
    const clamped = clampRange(rg, size)
    if (!clamped) return
    const rows = clamped.r1 - clamped.r0 + 1
    const cols = clamped.c1 - clamped.c0 + 1
    const empty: null[][] = Array.from({ length: rows }, () => new Array(cols).fill(null))
    this.transact('Clear', () => this.hf.setCellContents({ sheet: this.sid(tab), row: clamped.r0, col: clamped.c0 }, empty), { hf: true })
  }

  // ---------------- formatting ----------------

  private patchStyle(tab: GridSheet, r: number, c: number, patch: Partial<CellStyle>) {
    const k = styleKey(r, c)
    const next: CellStyle = { ...(tab.styles[k] ?? {}), ...patch }
    for (const key of Object.keys(next) as (keyof CellStyle)[]) if (next[key] === undefined) delete next[key]
    if (Object.keys(next).length) tab.styles[k] = next
    else delete tab.styles[k]
  }

  /** Apply a style patch (or a function of the old style) to every cell in a range. */
  applyStyle(tab: GridSheet, rg: Range, patch: Partial<CellStyle> | ((old: CellStyle | undefined, r: number, c: number) => Partial<CellStyle>), label = 'Format') {
    const size = this.size(tab)
    // Whole-row/column selections are limited to the used area (+ a margin) to stay fast.
    const r1 = Math.min(rg.r1, Math.max(size.rows, rg.r0) + 200)
    const c1 = Math.min(rg.c1, Math.max(size.cols, rg.c0) + 50)
    this.transact(label, () => {
      for (let r = rg.r0; r <= r1; r++) {
        for (let c = rg.c0; c <= c1; c++) {
          const p = typeof patch === 'function' ? patch(tab.styles[styleKey(r, c)], r, c) : patch
          this.patchStyle(tab, r, c, p)
        }
      }
    })
  }

  clearFormats(tab: GridSheet, rg: Range) {
    this.transact('Clear Formats', () => {
      for (const k of Object.keys(tab.styles)) {
        const [r, c] = k.split(':').map(Number)
        if (inRange(rg, r, c)) delete tab.styles[k]
      }
    })
  }

  setBorders(tab: GridSheet, rg: Range, mode: 'all' | 'outside' | 'bottom' | 'none', color = '#1f2937') {
    this.applyStyle(
      tab,
      rg,
      (_old, r, c) => {
        if (mode === 'none') return { border: undefined }
        if (mode === 'all') return { border: { top: color, right: color, bottom: color, left: color } }
        const b: Borders = { ...(_old?.border ?? {}) }
        if (mode === 'bottom') {
          if (r === rg.r1) b.bottom = color
        } else {
          if (r === rg.r0) b.top = color
          if (r === rg.r1) b.bottom = color
          if (c === rg.c0) b.left = color
          if (c === rg.c1) b.right = color
        }
        return { border: b }
      },
      'Borders'
    )
  }

  // ---------------- layout ----------------

  setColWidth(tab: GridSheet, cols: number[], width: number) {
    this.transact('Column Width', () => {
      for (const c of cols) {
        if (Math.abs(width - DEFAULT_COL_WIDTH) < 0.5) delete tab.colWidths[c]
        else tab.colWidths[c] = Math.max(2, Math.round(width))
      }
    })
  }

  setRowHeight(tab: GridSheet, rows: number[], height: number) {
    this.transact('Row Height', () => {
      for (const r of rows) {
        if (Math.abs(height - DEFAULT_ROW_HEIGHT) < 0.5) delete tab.rowHeights[r]
        else tab.rowHeights[r] = Math.max(2, Math.round(height))
      }
    })
  }

  setFreeze(tab: GridSheet, rows: number, cols: number) {
    this.transact('Freeze Panes', () => {
      tab.freeze = { rows, cols }
    })
  }

  merge(tab: GridSheet, rg: Range) {
    if (rg.r0 === rg.r1 && rg.c0 === rg.c1) return
    const size = this.size(tab)
    const clamped = clampRange(rg, size)
    this.transact(
      'Merge Cells',
      () => {
        tab.merges = tab.merges.filter((m) => !rangesIntersect(m, rg))
        tab.merges.push(rg)
        // keep top-left value only, like Excel
        if (clamped) {
          for (let r = clamped.r0; r <= clamped.r1; r++)
            for (let c = clamped.c0; c <= clamped.c1; c++)
              if (r !== rg.r0 || c !== rg.c0) this.hf.setCellContents({ sheet: this.sid(tab), row: r, col: c }, null)
        }
        this.patchStyle(tab, rg.r0, rg.c0, { align: 'center', valign: 'middle' })
      },
      { hf: true }
    )
  }

  unmerge(tab: GridSheet, rg: Range) {
    this.transact('Unmerge Cells', () => {
      tab.merges = tab.merges.filter((m) => !rangesIntersect(m, rg))
    })
  }

  // ---------------- rows & columns ----------------

  insertRows(tab: GridSheet, at: number, n: number) {
    this.transact(
      'Insert Rows',
      () => {
        this.hf.addRows(this.sid(tab), [at, n])
        shiftMeta(tab, 'row', at, n)
      },
      { hf: true }
    )
  }

  deleteRows(tab: GridSheet, at: number, n: number) {
    const rows = this.size(tab).rows
    this.transact(
      'Delete Rows',
      () => {
        if (at < rows) this.hf.removeRows(this.sid(tab), [at, Math.min(n, rows - at)])
        shiftMeta(tab, 'row', at, -n)
      },
      { hf: true }
    )
  }

  insertCols(tab: GridSheet, at: number, n: number) {
    this.transact(
      'Insert Columns',
      () => {
        this.hf.addColumns(this.sid(tab), [at, n])
        shiftMeta(tab, 'col', at, n)
      },
      { hf: true }
    )
  }

  deleteCols(tab: GridSheet, at: number, n: number) {
    const cols = this.size(tab).cols
    this.transact(
      'Delete Columns',
      () => {
        if (at < cols) this.hf.removeColumns(this.sid(tab), [at, Math.min(n, cols - at)])
        shiftMeta(tab, 'col', at, -n)
      },
      { hf: true }
    )
  }

  // ---------------- tabs ----------------

  uniqueName(base: string): string {
    if (!this.findTab(base)) return base
    let i = 2
    const stem = base.replace(/\s*\d+$/, '')
    while (this.findTab(`${stem} ${i}`) || this.findTab(`${stem}${i}`)) i++
    return /\d$/.test(base) ? `${stem}${i}` : `${stem} ${i}`
  }

  addSheet(name?: string, index?: number): GridSheet {
    const n = name ? this.uniqueName(name) : this.uniqueName(`Sheet${this.grids().length + 1}`)
    const tab = blankGrid(n)
    this.transact(
      'Add Sheet',
      () => {
        this.hf.addSheet(n)
        this.tabs.splice(index ?? this.tabs.length, 0, tab)
      },
      { hf: true }
    )
    return tab
  }

  addDashboard(name?: string): Dashboard {
    const d: Dashboard = { kind: 'dashboard', id: newId('d'), name: this.uniqueName(name ?? 'Dashboard'), widgets: [] }
    this.transact('Add Dashboard', () => {
      this.tabs.push(d)
    })
    return d
  }

  renameTab(id: string, name: string) {
    const tab = this.tab(id)
    const clean = name.trim().replace(/[[\]:*?/\\]/g, '').slice(0, 31)
    if (!tab || !clean || clean === tab.name) return
    const clash = this.findTab(clean)
    if (clash && clash.id !== id) throw new Error(`A sheet named "${clean}" already exists.`)
    const old = tab.name
    this.transact(
      'Rename',
      () => {
        if (tab.kind === 'grid') this.hf.renameSheet(this.sid(tab), clean)
        tab.name = clean
        if (tab.kind === 'grid') this.renameRefsInDashboards(old, clean)
      },
      { hf: tab.kind === 'grid' }
    )
  }

  private renameRefsInDashboards(oldName: string, newName: string) {
    const esc = oldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const re = new RegExp(`(^|[^\\w.])('?)${esc}\\2!`, 'gi')
    const q = /^[A-Za-z_][\w.]*$/.test(newName) ? newName : `'${newName}'`
    const fix = (s?: string) => s?.replace(re, (_m, pre: string) => `${pre}${q}!`)
    for (const t of this.tabs) {
      if (t.kind !== 'dashboard') continue
      t.widgets = t.widgets.map((w) => ({
        ...w,
        chart: w.chart && { ...w.chart, range: fix(w.chart.range)! },
        kpi: w.kpi && { ...w.kpi, formula: fix(w.kpi.formula)!, compare: fix(w.kpi.compare) },
        table: w.table && { range: fix(w.table.range)! }
      }))
    }
  }

  deleteTab(id: string) {
    const tab = this.tab(id)
    if (!tab) return
    if (tab.kind === 'grid' && this.grids().length === 1) throw new Error('A workbook needs at least one sheet.')
    this.transact(
      'Delete Sheet',
      () => {
        if (tab.kind === 'grid') this.hf.removeSheet(this.sid(tab))
        this.tabs = this.tabs.filter((t) => t.id !== id)
      },
      { hf: tab.kind === 'grid' }
    )
  }

  moveTab(id: string, toIndex: number) {
    const from = this.tabs.findIndex((t) => t.id === id)
    if (from < 0 || from === toIndex) return
    this.transact('Move Sheet', () => {
      const [t] = this.tabs.splice(from, 1)
      this.tabs.splice(Math.min(toIndex, this.tabs.length), 0, t)
    })
  }

  duplicateTab(id: string): Tab | undefined {
    const tab = this.tab(id)
    if (!tab) return
    if (tab.kind === 'dashboard') {
      const d: Dashboard = { ...(cloneTab(tab) as Dashboard), id: newId('d'), name: this.uniqueName(`${tab.name} copy`) }
      d.widgets = d.widgets.map((w) => ({ ...w, id: newId('w') }))
      this.transact('Duplicate', () => this.tabs.splice(this.tabs.indexOf(tab) + 1, 0, d))
      return d
    }
    const copy = { ...(cloneTab(tab) as GridSheet), id: newId('s'), name: this.uniqueName(`${tab.name} copy`) }
    const cells = this.hf.getSheetSerialized(this.sid(tab))
    this.transact(
      'Duplicate',
      () => {
        this.hf.addSheet(copy.name)
        this.hf.setSheetContent(this.hf.getSheetId(copy.name)!, cells)
        this.tabs.splice(this.tabs.indexOf(tab) + 1, 0, copy)
      },
      { hf: true }
    )
    return copy
  }

  // ---------------- dashboards ----------------

  updateDashboard(id: string, fn: (d: Dashboard) => void, label = 'Edit Dashboard') {
    const d = this.tab(id)
    if (!d || d.kind !== 'dashboard') return
    this.transact(label, () => fn(d))
  }

  /** Find or create a dashboard by name and add a widget to it (placed below existing ones). */
  addWidget(dashboardName: string | undefined, widget: Omit<Widget, 'id' | 'x' | 'y'> & Partial<Pick<Widget, 'x' | 'y'>>): { dashboard: Dashboard; widget: Widget } {
    let d = dashboardName ? this.findTab(dashboardName) : this.tabs.find((t) => t.kind === 'dashboard')
    if (d && d.kind !== 'dashboard') d = undefined
    const target: Dashboard = d ?? { kind: 'dashboard', id: newId('d'), name: this.uniqueName(dashboardName ?? 'Dashboard'), widgets: [] }
    const w = { ...widget, id: newId('w') } as Widget
    if (w.x === undefined || w.y === undefined) {
      const pos = placeWidget(target.widgets, w.w, w.h)
      w.x = pos.x
      w.y = pos.y
    }
    this.transact('Add Widget', () => {
      if (!d) this.tabs.push(target)
      target.widgets = [...target.widgets, w]
    })
    return { dashboard: target, widget: w }
  }

  // ---------------- data operations ----------------

  sortRange(tab: GridSheet, rg: Range, byCol: number, desc: boolean, hasHeader: boolean) {
    const r0 = hasHeader ? rg.r0 + 1 : rg.r0
    const size = this.size(tab)
    const r1 = Math.min(rg.r1, size.rows - 1)
    const c1 = Math.min(rg.c1, size.cols - 1)
    if (r1 <= r0 || c1 < rg.c0) return
    const area = { r0, c0: rg.c0, r1, c1 }
    const raws = this.rangeRaw(tab, area)
    const vals = this.rangeValues(tab, area)
    const keyCol = byCol - rg.c0
    const order = raws.map((_, i) => i)
    const rank = (v: CellValue) => (typeof v === 'number' ? 0 : typeof v === 'string' && v !== '' ? 1 : typeof v === 'boolean' ? 2 : v instanceof DetailedCellError ? 3 : 4)
    order.sort((a, b) => {
      const va = vals[a][keyCol]
      const vb = vals[b][keyCol]
      const ra = rank(va)
      const rb = rank(vb)
      if (ra === 4 || rb === 4) return ra - rb // blanks always last
      let cmp = ra - rb
      if (!cmp) {
        if (typeof va === 'number' && typeof vb === 'number') cmp = va - vb
        else cmp = String(va).localeCompare(String(vb), undefined, { numeric: true, sensitivity: 'base' })
      }
      return desc ? -cmp : cmp
    })
    const out = order.map((src, dst) => raws[src].map((v) => (typeof v === 'string' && v.startsWith('=') ? shiftFormula(v, dst - src, 0) : v)))
    const styles = order.map((src) => raws[src].map((_, j) => tab.styles[styleKey(r0 + src, rg.c0 + j)]))
    this.writeRaw(tab, r0, rg.c0, out, styles, 'Sort')
  }

  removeDuplicates(tab: GridSheet, rg: Range, hasHeader: boolean): number {
    const size = this.size(tab)
    const area = clampRange({ ...rg, r0: hasHeader ? rg.r0 + 1 : rg.r0 }, size)
    if (!area) return 0
    const raws = this.rangeRaw(tab, area)
    const vals = this.rangeValues(tab, area)
    const seen = new Set<string>()
    const keep: RawCellContent[][] = []
    vals.forEach((row, i) => {
      const key = JSON.stringify(row.map((v) => (v instanceof DetailedCellError ? v.value : v)))
      if (!seen.has(key)) {
        seen.add(key)
        keep.push(raws[i])
      }
    })
    const removed = raws.length - keep.length
    if (!removed) return 0
    const width = area.c1 - area.c0 + 1
    while (keep.length < raws.length) keep.push(new Array(width).fill(null))
    this.writeRaw(tab, area.r0, area.c0, keep, undefined, 'Remove Duplicates')
    return removed
  }

  /** Excel-style fill (handle drag, Cmd+D, Cmd+R): extends series and shifts formulas. */
  fill(tab: GridSheet, src: Range, dst: Range) {
    const vertical = dst.c0 === src.c0 && dst.c1 === src.c1
    const forward = vertical ? dst.r1 > src.r1 : dst.c1 > src.c1
    const srcLen = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1
    const dstLen = vertical ? dst.r1 - dst.r0 + 1 : dst.c1 - dst.c0 + 1
    const extra = dstLen - srcLen
    if (extra <= 0) return
    const lanes = vertical ? src.c1 - src.c0 + 1 : src.r1 - src.r0 + 1
    const raws = this.rangeRaw(tab, src)
    const out: RawCellContent[][] = vertical
      ? Array.from({ length: extra }, () => new Array(lanes).fill(null))
      : Array.from({ length: lanes }, () => new Array(extra).fill(null))
    const styles: (CellStyle | undefined)[][] = vertical
      ? Array.from({ length: extra }, () => new Array(lanes))
      : Array.from({ length: lanes }, () => new Array(extra))

    for (let lane = 0; lane < lanes; lane++) {
      const seq = Array.from({ length: srcLen }, (_, k) => (vertical ? raws[k][lane] : raws[lane][k]))
      const seqStyles = Array.from({ length: srcLen }, (_, k) => (vertical ? tab.styles[styleKey(src.r0 + k, src.c0 + lane)] : tab.styles[styleKey(src.r0 + lane, src.c0 + k)]))
      const nums = seq.every((v) => typeof v === 'number')
      const fmt = seqStyles[0]?.fmt
      let step = 0
      if (nums && srcLen >= 2) step = ((seq[srcLen - 1] as number) - (seq[0] as number)) / (srcLen - 1)
      else if (nums && srcLen === 1 && isDateFormat(fmt)) step = 1
      const textNum = srcLen === 1 && typeof seq[0] === 'string' && !seq[0].startsWith('=') ? /^(.*?)(\d+)$/.exec(seq[0]) : null

      for (let k = 0; k < extra; k++) {
        // position of the generated cell relative to the source start (negative when filling backwards)
        const offset = forward ? srcLen + k : -(extra - k)
        const srcIdx = ((offset % srcLen) + srcLen) % srcLen
        let v: RawCellContent
        if (nums && step !== 0) {
          v = (seq[0] as number) + step * offset
        } else if (textNum) {
          v = `${textNum[1]}${Math.max(0, Number(textNum[2]) + offset)}`
        } else {
          v = seq[srcIdx]
          if (typeof v === 'string' && v.startsWith('=')) {
            const delta = offset - srcIdx
            v = vertical ? shiftFormula(v, delta, 0) : shiftFormula(v, 0, delta)
          }
        }
        if (vertical) {
          out[k][lane] = v
          styles[k][lane] = seqStyles[srcIdx]
        } else {
          out[lane][k] = v
          styles[lane][k] = seqStyles[srcIdx]
        }
      }
    }
    const startR = vertical ? (forward ? src.r1 + 1 : dst.r0) : src.r0
    const startC = vertical ? src.c0 : forward ? src.c1 + 1 : dst.c0
    this.writeRaw(tab, startR, startC, out, styles, 'Fill')
  }

  // ---------------- filters ----------------

  setFilter(tab: GridSheet, filter: FilterState | undefined) {
    const size = this.size(tab)
    const hidden: Record<number, true> = {}
    if (filter) {
      const r1 = Math.min(filter.range.r1, size.rows - 1)
      for (let r = filter.range.r0 + 1; r <= r1; r++) {
        for (const [col, allowed] of Object.entries(filter.criteria)) {
          if (!allowed.includes(this.display(tab, r, Number(col)))) {
            hidden[r] = true
            break
          }
        }
      }
    }
    this.transact('Filter', () => {
      tab.filter = filter
      tab.hiddenRows = hidden
    })
  }

  /** Distinct display values in a filter column (for the filter popup). */
  distinctValues(tab: GridSheet, rg: Range, col: number): string[] {
    const size = this.size(tab)
    const set = new Set<string>()
    for (let r = rg.r0 + 1; r <= Math.min(rg.r1, size.rows - 1); r++) set.add(this.display(tab, r, col))
    return [...set].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
  }

  setCondRules(tab: GridSheet, rules: CondRule[]) {
    this.transact('Conditional Formatting', () => {
      tab.cond = rules
    })
  }

  markSaved() {
    this.dirty = false
    this.emit()
  }
}

export function clampRange(rg: Range, size: { rows: number; cols: number }): Range | null {
  const r1 = Math.min(rg.r1, size.rows - 1)
  const c1 = Math.min(rg.c1, size.cols - 1)
  if (r1 < rg.r0 || c1 < rg.c0) return null
  return { r0: rg.r0, c0: rg.c0, r1, c1 }
}

function shiftMeta(tab: GridSheet, axis: 'row' | 'col', at: number, n: number) {
  const isRow = axis === 'row'
  const move = (i: number): number | null => {
    if (i < at) return i
    if (n < 0 && i < at - n) return null // deleted
    return i + n
  }
  const styles: Record<string, CellStyle> = {}
  for (const [k, st] of Object.entries(tab.styles)) {
    const [r, c] = k.split(':').map(Number)
    const m = move(isRow ? r : c)
    if (m === null) continue
    styles[isRow ? styleKey(m, c) : styleKey(r, m)] = st
  }
  tab.styles = styles
  const remap = <T,>(rec: Record<number, T>) => {
    const out: Record<number, T> = {}
    for (const [k, v] of Object.entries(rec)) {
      const m = move(Number(k))
      if (m !== null) out[m] = v
    }
    return out
  }
  if (isRow) {
    tab.rowHeights = remap(tab.rowHeights)
    tab.hiddenRows = remap(tab.hiddenRows)
  } else tab.colWidths = remap(tab.colWidths)

  const lo = isRow ? 'r0' : 'c0'
  const hi = isRow ? 'r1' : 'c1'
  const shiftRange = (rg: Range): Range | null => {
    let a = rg[lo]
    let b = rg[hi]
    if (n > 0) {
      if (a >= at) a += n
      if (b >= at) b += n
    } else {
      const delEnd = at - n - 1
      if (a >= at && b <= delEnd) return null
      const adj = (i: number) => (i > delEnd ? i + n : i >= at ? at : i)
      a = adj(a)
      b = b > delEnd ? b + n : b >= at ? at - 1 : b
      if (b < a) return null
    }
    return { ...rg, [lo]: a, [hi]: b }
  }
  tab.merges = tab.merges.map(shiftRange).filter((m): m is Range => !!m && !(m.r0 === m.r1 && m.c0 === m.c1))
  if (tab.filter) {
    const f = shiftRange(tab.filter.range)
    tab.filter = f ? { ...tab.filter, range: f } : undefined
  }
}

function placeWidget(widgets: Widget[], w: number, h: number): { x: number; y: number } {
  const cols = 12
  for (let y = 0; y < 200; y++) {
    for (let x = 0; x + w <= cols; x++) {
      const free = widgets.every((o) => x + w <= o.x || o.x + o.w <= x || y + h <= o.y || o.y + o.h <= y)
      if (free) return { x, y }
    }
  }
  return { x: 0, y: widgets.reduce((m, o) => Math.max(m, o.y + o.h), 0) }
}

function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace('#', '')
  const full = h.length === 3 ? h.split('').map((ch) => ch + ch).join('') : h
  const n = parseInt(full, 16)
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255]
}

function mixColor(a: string, b: string, t: number): string {
  const ca = hexToRgb(a)
  const cb = hexToRgb(b)
  const m = ca.map((v, i) => Math.round(v + (cb[i] - v) * Math.max(0, Math.min(1, t))))
  return `#${m.map((v) => v.toString(16).padStart(2, '0')).join('')}`
}
