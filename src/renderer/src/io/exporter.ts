import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'
import { DetailedCellError, type CellValue } from 'hyperformula'
import { cellName, colName, parseRange, rangeName } from '../engine/addr'
import { styleKey, type CellStyle, type CondRule, type GridSheet, type Workbook } from '../engine/workbook'
import { META_SHEET, type MetaPayload } from './importer'

const enc = new TextEncoder()

/** Formats we write natively with full formatting; everything else goes through SheetJS. */
export const LOSSLESS_EXT = new Set(['xlsx'])

const SHEETJS_BOOK: Record<string, XLSX.BookType> = {
  xls: 'biff8',
  xlsb: 'xlsb',
  ods: 'ods',
  fods: 'fods',
  xml: 'xlml',
  slk: 'sylk',
  dif: 'dif',
  dbf: 'dbf',
  numbers: 'numbers'
}

export async function exportWorkbook(wb: Workbook, ext: string, active: GridSheet): Promise<Uint8Array> {
  switch (ext) {
    case 'xlsx':
      return exportXlsx(wb)
    case 'csv':
      return enc.encode(`﻿${toDelimited(wb, active, ',')}`)
    case 'tsv':
    case 'txt':
    case 'tab':
      return enc.encode(toDelimited(wb, active, '\t'))
    case 'html':
    case 'htm':
      return enc.encode(sheetToHtml(wb, active, false))
    default: {
      const bookType = SHEETJS_BOOK[ext]
      if (!bookType) throw new Error(`Saving as .${ext} is not supported.`)
      return exportSheetJS(wb, bookType)
    }
  }
}

function plainValue(v: CellValue): string | number | boolean | null {
  if (v instanceof DetailedCellError) return v.value
  return v as string | number | boolean | null
}

// ---------------------------------------------------------------- XLSX via ExcelJS

function argb(hex: string) {
  return `FF${hex.replace('#', '').toUpperCase().padStart(6, '0')}`
}

function toExcelStyle(st: CellStyle): Partial<ExcelJS.Style> {
  const out: Partial<ExcelJS.Style> = {}
  const font: Partial<ExcelJS.Font> = {}
  if (st.bold) font.bold = true
  if (st.italic) font.italic = true
  if (st.underline) font.underline = true
  if (st.strike) font.strike = true
  if (st.size) font.size = Math.round(st.size * 0.75)
  if (st.font) font.name = st.font
  if (st.color) font.color = { argb: argb(st.color) }
  if (Object.keys(font).length) out.font = font
  if (st.fill) out.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(st.fill) } }
  if (st.align || st.valign || st.wrap) {
    out.alignment = { horizontal: st.align, vertical: st.valign, wrapText: st.wrap || undefined }
  }
  if (st.fmt) out.numFmt = st.fmt
  if (st.border) {
    const side = (c?: string): Partial<ExcelJS.Border> | undefined => (c ? { style: 'thin', color: { argb: argb(c) } } : undefined)
    out.border = { top: side(st.border.top), right: side(st.border.right), bottom: side(st.border.bottom), left: side(st.border.left) }
  }
  return out
}

function condToExcel(rule: CondRule): ExcelJS.ConditionalFormattingRule | null {
  const style: Partial<ExcelJS.Style> = {}
  if (rule.fill) style.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(rule.fill) } }
  if (rule.color || rule.bold) style.font = { color: rule.color ? { argb: argb(rule.color) } : undefined, bold: rule.bold }
  const priority = 1
  switch (rule.type) {
    case 'gt':
      return { type: 'cellIs', operator: 'greaterThan', formulae: [rule.a ?? '0'], style, priority }
    case 'lt':
      return { type: 'cellIs', operator: 'lessThan', formulae: [rule.a ?? '0'], style, priority }
    case 'between':
      return { type: 'cellIs', operator: 'between', formulae: [rule.a ?? '0', rule.b ?? '0'], style, priority }
    case 'eq': {
      const isNum = rule.a !== undefined && rule.a !== '' && !Number.isNaN(Number(rule.a))
      return { type: 'cellIs', operator: 'equal', formulae: [isNum ? rule.a! : `"${rule.a ?? ''}"`], style, priority }
    }
    case 'contains':
      return { type: 'containsText', operator: 'containsText', text: rule.a ?? '', style, priority } as ExcelJS.ConditionalFormattingRule
    case 'duplicate':
      return { type: 'duplicateValues', style, priority } as unknown as ExcelJS.ConditionalFormattingRule
    case 'scale':
      return {
        type: 'colorScale',
        priority,
        cfvo: [{ type: 'min' }, { type: 'max' }],
        color: [{ argb: argb(rule.scale?.[0] ?? '#ffffff') }, { argb: argb(rule.scale?.[1] ?? '#63be7b') }]
      } as ExcelJS.ConditionalFormattingRule
    default:
      return null
  }
}

async function exportXlsx(wb: Workbook): Promise<Uint8Array> {
  const book = new ExcelJS.Workbook()
  book.creator = 'Open Sheets'
  book.created = new Date()
  const cond: Record<string, CondRule[]> = {}
  const tabColors: Record<string, string> = {}

  for (const tab of wb.grids()) {
    const ws = book.addWorksheet(tab.name, {
      views: tab.freeze.rows || tab.freeze.cols ? [{ state: 'frozen', xSplit: tab.freeze.cols, ySplit: tab.freeze.rows }] : [],
      properties: tab.tabColor ? { tabColor: { argb: argb(tab.tabColor) } } : undefined
    })
    if (tab.tabColor) tabColors[tab.name] = tab.tabColor
    const sid = wb.sid(tab)
    const raws = wb.hf.getSheetSerialized(sid)
    const vals = wb.hf.getSheetValues(sid)
    raws.forEach((row, r) =>
      row.forEach((raw, c) => {
        if (raw === null || raw === undefined || raw === '') return
        const cell = ws.getCell(r + 1, c + 1)
        if (typeof raw === 'string' && raw.startsWith('=')) {
          const v = vals[r]?.[c]
          const result = v instanceof DetailedCellError ? { error: v.value } : (v ?? undefined)
          cell.value = { formula: raw.slice(1), result } as ExcelJS.CellFormulaValue
        } else {
          cell.value = raw as ExcelJS.CellValue
        }
      })
    )
    for (const [k, st] of Object.entries(tab.styles)) {
      const [r, c] = k.split(':').map(Number)
      const cell = ws.getCell(r + 1, c + 1)
      Object.assign(cell, toExcelStyle(st))
    }
    for (const [c, px] of Object.entries(tab.colWidths)) ws.getColumn(Number(c) + 1).width = Math.max(0, (px - 5) / 7)
    for (const [r, px] of Object.entries(tab.rowHeights)) ws.getRow(Number(r) + 1).height = px * 0.75
    for (const r of Object.keys(tab.hiddenRows)) ws.getRow(Number(r) + 1).hidden = true
    for (const m of tab.merges) ws.mergeCells(m.r0 + 1, m.c0 + 1, m.r1 + 1, m.c1 + 1)
    if (tab.filter) ws.autoFilter = rangeName(tab.filter.range)
    if (tab.cond.length) {
      cond[tab.name] = tab.cond
      for (const rule of tab.cond) {
        const x = condToExcel(rule)
        if (x && parseRange(rule.range)) ws.addConditionalFormatting({ ref: rule.range, rules: [x] })
      }
    }
  }

  // Dashboards and app-only settings travel in a hidden sheet so the file stays a normal .xlsx.
  const meta: MetaPayload = {
    version: 1,
    dashboards: wb.tabs.filter((t) => t.kind === 'dashboard') as MetaPayload['dashboards'],
    order: wb.tabs.map((t) => t.name),
    cond,
    tabColors
  }
  const json = JSON.stringify(meta)
  const metaWs = book.addWorksheet(META_SHEET, { state: 'veryHidden' })
  for (let i = 0, row = 1; i < json.length; i += 30000, row++) metaWs.getCell(row, 1).value = json.slice(i, i + 30000)

  const buf = await book.xlsx.writeBuffer()
  return new Uint8Array(buf as ArrayBuffer)
}

// ---------------------------------------------------------------- other formats via SheetJS

async function exportSheetJS(wb: Workbook, bookType: XLSX.BookType): Promise<Uint8Array> {
  const book = XLSX.utils.book_new()
  for (const tab of wb.grids()) {
    const sid = wb.sid(tab)
    const raws = wb.hf.getSheetSerialized(sid)
    const vals = wb.hf.getSheetValues(sid)
    const ws: XLSX.WorkSheet = {}
    let maxR = 0
    let maxC = 0
    raws.forEach((row, r) =>
      row.forEach((raw, c) => {
        if (raw === null || raw === undefined || raw === '') return
        const v = plainValue(vals[r]?.[c] ?? null)
        const cell: XLSX.CellObject =
          typeof v === 'number' ? { t: 'n', v } : typeof v === 'boolean' ? { t: 'b', v } : vals[r]?.[c] instanceof DetailedCellError ? { t: 'e', v: 0, w: String(v) } : { t: 's', v: v ?? '' }
        if (typeof raw === 'string' && raw.startsWith('=')) cell.f = raw.slice(1)
        const fmt = tab.styles[styleKey(r, c)]?.fmt
        if (fmt) cell.z = fmt
        ws[cellName(r, c)] = cell
        maxR = Math.max(maxR, r)
        maxC = Math.max(maxC, c)
      })
    )
    ws['!ref'] = `A1:${colName(maxC)}${maxR + 1}`
    ws['!merges'] = tab.merges.map((m) => ({ s: { r: m.r0, c: m.c0 }, e: { r: m.r1, c: m.c1 } }))
    ws['!cols'] = Array.from({ length: maxC + 1 }, (_, c) => (tab.colWidths[c] ? { wpx: tab.colWidths[c] } : {}))
    XLSX.utils.book_append_sheet(book, ws, tab.name)
  }
  const opts: XLSX.WritingOptions = { bookType, type: 'array', compression: true }
  if (bookType === 'numbers') {
    const { default: payload } = await import('xlsx/dist/xlsx.zahl.mjs')
    ;(opts as XLSX.WritingOptions & { numbers: string }).numbers = payload
  }
  if (bookType === 'dif' || bookType === 'sylk' || bookType === 'dbf') {
    // single-sheet formats: keep only the first sheet
    book.SheetNames = book.SheetNames.slice(0, 1)
  }
  const out = XLSX.write(book, opts) as ArrayBuffer
  return new Uint8Array(out)
}

// ---------------------------------------------------------------- text formats

function toDelimited(wb: Workbook, tab: GridSheet, sep: string): string {
  const { rows, cols } = wb.size(tab)
  const lines: string[] = []
  for (let r = 0; r < rows; r++) {
    const cells: string[] = []
    for (let c = 0; c < cols; c++) {
      const s = wb.display(tab, r, c)
      cells.push(/[",\n\r\t]/.test(s) || s.includes(sep) ? `"${s.replace(/"/g, '""')}"` : s)
    }
    lines.push(cells.join(sep))
  }
  return lines.join('\r\n')
}

const escapeHtml = (s: string) => s.replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]!)

/** Styled HTML table of a sheet's used range (HTML export, print and PDF). */
export function sheetToHtml(wb: Workbook, tab: GridSheet, forPrint: boolean): string {
  const { rows, cols } = wb.size(tab)
  const covered = new Set<string>()
  const spans = new Map<string, { rs: number; cs: number }>()
  for (const m of tab.merges) {
    spans.set(styleKey(m.r0, m.c0), { rs: m.r1 - m.r0 + 1, cs: m.c1 - m.c0 + 1 })
    for (let r = m.r0; r <= m.r1; r++) for (let c = m.c0; c <= m.c1; c++) if (r !== m.r0 || c !== m.c0) covered.add(styleKey(r, c))
  }
  let body = ''
  for (let r = 0; r < rows; r++) {
    if (tab.hiddenRows[r]) continue
    body += '<tr>'
    for (let c = 0; c < cols; c++) {
      const k = styleKey(r, c)
      if (covered.has(k)) continue
      const st = { ...(tab.styles[k] ?? {}), ...(wb.condStyle(tab, r, c) ?? {}) }
      const v = wb.value(tab, r, c)
      const css: string[] = []
      if (st.bold) css.push('font-weight:600')
      if (st.italic) css.push('font-style:italic')
      if (st.underline || st.strike) css.push(`text-decoration:${[st.underline && 'underline', st.strike && 'line-through'].filter(Boolean).join(' ')}`)
      if (st.color) css.push(`color:${st.color}`)
      if (st.fill) css.push(`background:${st.fill}`)
      if (st.size) css.push(`font-size:${st.size}px`)
      if (st.font) css.push(`font-family:'${st.font}'`)
      css.push(`text-align:${st.align ?? (typeof v === 'number' ? 'right' : 'left')}`)
      if (st.valign) css.push(`vertical-align:${st.valign}`)
      if (!st.wrap) css.push('white-space:nowrap')
      for (const side of ['top', 'right', 'bottom', 'left'] as const) {
        const col = st.border?.[side]
        if (col) css.push(`border-${side}:1px solid ${col}`)
      }
      const span = spans.get(k)
      const attrs = span ? ` rowspan="${span.rs}" colspan="${span.cs}"` : ''
      body += `<td${attrs} style="${css.join(';')}">${escapeHtml(wb.display(tab, r, c))}</td>`
    }
    body += '</tr>\n'
  }
  const colgroup = Array.from({ length: cols }, (_, c) => `<col style="width:${tab.colWidths[c] ?? 100}px">`).join('')
  return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(tab.name)}</title>
<style>
  body{font-family:-apple-system,"Helvetica Neue",Arial,sans-serif;font-size:12px;color:#111;margin:${forPrint ? '0' : '24px'}}
  table{border-collapse:collapse;table-layout:fixed}
  td{padding:3px 6px;overflow:hidden;${forPrint ? '' : 'border:1px solid #e5e7eb;'}}
  h1{font-size:15px;margin:0 0 10px}
  @page{margin:12mm}
</style></head><body>${forPrint ? '' : `<h1>${escapeHtml(tab.name)}</h1>`}
<table><colgroup>${colgroup}</colgroup>${body}</table></body></html>`
}
