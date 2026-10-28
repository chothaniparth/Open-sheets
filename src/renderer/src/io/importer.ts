import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'
import * as cptable from 'xlsx/dist/cpexcel.full.mjs'
import type { RawCellContent } from 'hyperformula'
import type { Range } from '../engine/addr'
import { parseRange } from '../engine/addr'
import { jsDateToSerial } from '../engine/format'
import { styleKey, type Borders, type CellStyle, type CondRule, type Dashboard, type SheetData, type WorkbookData } from '../engine/workbook'

XLSX.set_cptable(cptable)

export const META_SHEET = '_opensheets_meta'

export interface MetaPayload {
  version: 1
  dashboards: Dashboard[]
  order: string[]
  cond: Record<string, CondRule[]>
  tabColors?: Record<string, string>
}

const EXCELJS_EXT = new Set(['xlsx', 'xlsm', 'xltx', 'xltm'])

export async function importFile(name: string, data: Uint8Array): Promise<WorkbookData> {
  const ext = name.split('.').pop()?.toLowerCase() ?? ''
  if (EXCELJS_EXT.has(ext)) {
    try {
      return await importWithExcelJS(data)
    } catch (err) {
      console.warn('ExcelJS could not read the file, falling back to SheetJS', err)
    }
  }
  return importWithSheetJS(data, ext)
}

export function cleanFormula(f: string): string {
  return f.replace(/_xlfn\.|_xlws\.|_xlpm\./g, '')
}

// ---------------------------------------------------------------- ExcelJS (styled OOXML)

// Office default theme colors (indices used by <color theme="n">).
const THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47']

function colorOf(c: Partial<ExcelJS.Color> & { tint?: number; indexed?: number } | undefined): string | undefined {
  if (!c) return undefined
  if (c.argb && /^[0-9A-Fa-f]{8}$/.test(c.argb)) return `#${c.argb.slice(2).toLowerCase()}`
  if (c.argb && /^[0-9A-Fa-f]{6}$/.test(c.argb)) return `#${c.argb.toLowerCase()}`
  if (typeof c.theme === 'number' && THEME[c.theme]) {
    let rgb = THEME[c.theme].match(/../g)!.map((h) => parseInt(h, 16))
    const tint = c.tint ?? 0
    if (tint > 0) rgb = rgb.map((v) => Math.round(v + (255 - v) * tint))
    if (tint < 0) rgb = rgb.map((v) => Math.round(v * (1 + tint)))
    return `#${rgb.map((v) => v.toString(16).padStart(2, '0')).join('')}`
  }
  return undefined
}

function styleFromExcel(cell: ExcelJS.Cell): CellStyle | undefined {
  const st: CellStyle = {}
  const f = cell.font
  if (f) {
    if (f.bold) st.bold = true
    if (f.italic) st.italic = true
    if (f.underline) st.underline = true
    if (f.strike) st.strike = true
    if (f.size && f.size !== 11) st.size = Math.round(f.size * (4 / 3))
    if (f.name && !/^(Calibri|Aptos|Arial)/.test(f.name)) st.font = f.name
    const col = colorOf(f.color as never)
    if (col && col !== '#000000') st.color = col
  }
  const fill = cell.fill as ExcelJS.FillPattern | undefined
  if (fill?.type === 'pattern' && fill.pattern === 'solid') {
    const col = colorOf(fill.fgColor as never)
    if (col) st.fill = col
  }
  const a = cell.alignment
  if (a) {
    if (a.horizontal === 'left' || a.horizontal === 'right') st.align = a.horizontal
    if (a.horizontal === 'center' || a.horizontal === 'centerContinuous') st.align = 'center'
    if (a.vertical === 'top' || a.vertical === 'bottom') st.valign = a.vertical
    if (a.vertical === 'middle') st.valign = 'middle'
    if (a.wrapText) st.wrap = true
  }
  if (cell.numFmt && cell.numFmt !== 'General') st.fmt = cell.numFmt
  const b = cell.border
  if (b) {
    const side = (s?: Partial<ExcelJS.Border>) => (s?.style ? (colorOf(s.color as never) ?? '#000000') : undefined)
    const border: Borders = { top: side(b.top), right: side(b.right), bottom: side(b.bottom), left: side(b.left) }
    if (border.top || border.right || border.bottom || border.left) st.border = border
  }
  return Object.keys(st).length ? st : undefined
}

function valueFromExcel(cell: ExcelJS.Cell): RawCellContent {
  const v = cell.value as unknown
  if (v === null || v === undefined) return null
  if (typeof v === 'number' || typeof v === 'boolean') return v
  if (typeof v === 'string') return v.startsWith('=') ? `'${v}` : v
  if (v instanceof Date) return jsDateToSerial(v)
  if (typeof v === 'object') {
    const o = v as Record<string, any>
    const formula = (cell as any).formula ?? o.formula
    if (formula) return `=${cleanFormula(formula)}`
    if (o.sharedFormula && o.result !== undefined) return o.result instanceof Date ? jsDateToSerial(o.result) : o.result
    if (Array.isArray(o.richText)) return o.richText.map((t: { text: string }) => t.text).join('')
    if (o.text !== undefined) return typeof o.text === 'string' ? o.text : (o.text.richText?.map((t: { text: string }) => t.text).join('') ?? '')
    if (o.error) return o.error
  }
  return String(v)
}

async function importWithExcelJS(data: Uint8Array): Promise<WorkbookData> {
  const book = new ExcelJS.Workbook()
  await book.xlsx.load(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) as ArrayBuffer)
  const sheets: SheetData[] = []
  let meta: MetaPayload | undefined

  for (const ws of book.worksheets) {
    if (ws.name === META_SHEET) {
      meta = readMeta(ws)
      continue
    }
    const cells: RawCellContent[][] = []
    const styles: Record<string, CellStyle> = {}
    const rowHeights: Record<number, number> = {}
    const hiddenRows: Record<number, true> = {}
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      const r = rowNumber - 1
      if (row.height && Math.abs(row.height - 15) > 0.5) rowHeights[r] = Math.round(row.height * (4 / 3))
      if (row.hidden) hiddenRows[r] = true
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const c = colNumber - 1
        if (cell.isMerged && cell.master !== cell) return
        const v = valueFromExcel(cell)
        if (v !== null && v !== '') (cells[r] ??= [])[c] = v
        const st = styleFromExcel(cell)
        if (st) styles[styleKey(r, c)] = st
      })
    })
    for (let r = 0; r < cells.length; r++) {
      const row = (cells[r] ??= [])
      for (let c = 0; c < row.length; c++) if (row[c] === undefined) row[c] = null
    }
    const colWidths: Record<number, number> = {}
    ;(ws.columns ?? []).forEach((col, i) => {
      if (col?.width && Math.abs(col.width - 8.43) > 0.1) colWidths[i] = Math.round(col.width * 7 + 5)
    })
    const merges: Range[] = ((ws.model as any)?.merges ?? []).map((m: string) => parseRange(m)).filter(Boolean)
    const view = ws.views?.[0] as { state?: string; xSplit?: number; ySplit?: number } | undefined
    const freeze = view?.state === 'frozen' ? { rows: view.ySplit ?? 0, cols: view.xSplit ?? 0 } : { rows: 0, cols: 0 }
    sheets.push({ name: ws.name, cells, meta: { styles, colWidths, rowHeights, hiddenRows, merges, freeze } })
  }

  applyMeta(sheets, meta)
  return { sheets, dashboards: meta?.dashboards ?? [], order: meta?.order }
}

function readMeta(ws: ExcelJS.Worksheet): MetaPayload | undefined {
  let json = ''
  ws.getColumn(1).eachCell((cell) => {
    json += String(cell.value ?? '')
  })
  try {
    return JSON.parse(json) as MetaPayload
  } catch {
    return undefined
  }
}

function applyMeta(sheets: SheetData[], meta?: MetaPayload) {
  if (!meta) return
  for (const s of sheets) {
    const rules = meta.cond?.[s.name]
    if (rules) s.meta = { ...s.meta, cond: rules }
    const color = meta.tabColors?.[s.name]
    if (color) s.meta = { ...s.meta, tabColor: color }
  }
}

// ---------------------------------------------------------------- SheetJS (everything else)

function importWithSheetJS(data: Uint8Array, ext: string): WorkbookData {
  const book = XLSX.read(data, {
    type: 'array',
    cellFormula: true,
    cellNF: true,
    cellStyles: true,
    cellDates: false,
    raw: ext === 'csv' || ext === 'tsv' ? false : undefined
  })
  const sheets: SheetData[] = []
  let meta: MetaPayload | undefined
  for (const name of book.SheetNames) {
    const ws = book.Sheets[name]
    if (name === META_SHEET) {
      const rows = XLSX.utils.sheet_to_json<string[]>(ws, { header: 1 })
      try {
        meta = JSON.parse(rows.map((r) => r[0] ?? '').join(''))
      } catch {
        /* ignore */
      }
      continue
    }
    const cells: RawCellContent[][] = []
    const styles: Record<string, CellStyle> = {}
    for (const key of Object.keys(ws)) {
      if (key[0] === '!') continue
      const cell = ws[key] as XLSX.CellObject
      const { r, c } = XLSX.utils.decode_cell(key)
      let v: RawCellContent = null
      if (cell.f) v = `=${cleanFormula(cell.f)}`
      else if (cell.t === 'n' || cell.t === 'b') v = cell.v as number | boolean
      else if (cell.t === 's') v = String(cell.v ?? '')
      else if (cell.t === 'e') v = cell.w ?? '#N/A'
      else if (cell.t === 'd') v = jsDateToSerial(cell.v as Date)
      if (v !== null && v !== '') {
        const row = (cells[r] ??= [])
        row[c] = v
      }
      const fmt = typeof cell.z === 'string' ? cell.z : undefined
      if (fmt && fmt !== 'General') styles[styleKey(r, c)] = { fmt }
    }
    for (let r = 0; r < cells.length; r++) {
      const row = (cells[r] ??= [])
      for (let c = 0; c < row.length; c++) if (row[c] === undefined) row[c] = null
    }
    const merges: Range[] = (ws['!merges'] ?? []).map((m) => ({ r0: m.s.r, c0: m.s.c, r1: m.e.r, c1: m.e.c }))
    const colWidths: Record<number, number> = {}
    ;(ws['!cols'] ?? []).forEach((col, i) => {
      const px = col?.wpx ?? (col?.wch ? col.wch * 7 + 5 : undefined)
      if (px && Math.abs(px - 64) > 2) colWidths[i] = Math.round(px)
    })
    const rowHeights: Record<number, number> = {}
    const hiddenRows: Record<number, true> = {}
    ;(ws['!rows'] ?? []).forEach((row, i) => {
      const px = row?.hpx ?? (row?.hpt ? row.hpt * (4 / 3) : undefined)
      if (px) rowHeights[i] = Math.round(px)
      if (row?.hidden) hiddenRows[i] = true
    })
    sheets.push({ name: name.slice(0, 31) || `Sheet${sheets.length + 1}`, cells, meta: { styles, merges, colWidths, rowHeights, hiddenRows } })
  }
  if (!sheets.length) sheets.push({ name: 'Sheet1', cells: [] })
  applyMeta(sheets, meta)
  return { sheets, dashboards: meta?.dashboards ?? [], order: meta?.order }
}
