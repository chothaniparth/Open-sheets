import { DetailedCellError } from 'hyperformula'
import { colIndex, parseCell, parseRange, quoteSheet, rangeName, splitSheet, type Range } from '../engine/addr'
import { clampRange, type CellStyle, type ChartType, type GridSheet, type Workbook } from '../engine/workbook'

export interface ToolHost {
  wb: () => Workbook
  activeGrid: () => GridSheet
  selection: () => Range | null
  setActive: (tabId: string) => void
}

const MAX_CELLS = 2500

function sheetFor(host: ToolHost, name?: string): GridSheet {
  if (!name) return host.activeGrid()
  const t = host.wb().findTab(name)
  if (!t) throw new Error(`No sheet named "${name}". Call get_workbook_info to see sheet names.`)
  if (t.kind !== 'grid') throw new Error(`"${name}" is a dashboard, not a worksheet.`)
  return t
}

function target(host: ToolHost, sheet: string | undefined, ref: string): { tab: GridSheet; range: Range } {
  const split = splitSheet(ref)
  const tab = sheetFor(host, split.sheet ?? sheet)
  const range = parseRange(split.ref, host.wb().size(tab))
  if (!range) throw new Error(`"${ref}" is not a valid A1 range.`)
  return { tab, range }
}

function plain(v: unknown) {
  if (v instanceof DetailedCellError) return v.value
  return v as string | number | boolean | null
}

function qualify(host: ToolHost, ref: string) {
  return splitSheet(ref).sheet ? ref : `${quoteSheet(host.activeGrid().name)}!${ref}`
}

type Input = Record<string, any>

export async function runTool(host: ToolHost, name: string, input: Input): Promise<unknown> {
  const wb = host.wb()
  switch (name) {
    case 'get_workbook_info': {
      const sel = host.selection()
      return {
        activeSheet: host.activeGrid().name,
        selection: sel ? rangeName(sel) : null,
        tabs: wb.tabs.map((t) => {
          if (t.kind === 'dashboard') return { name: t.name, kind: 'dashboard', widgets: t.widgets.map((w) => ({ type: w.type, title: w.title })) }
          const { rows, cols } = wb.size(t)
          return { name: t.name, kind: 'worksheet', usedRange: rows && cols ? rangeName({ r0: 0, c0: 0, r1: rows - 1, c1: cols - 1 }) : 'empty', frozen: t.freeze }
        })
      }
    }
    case 'read_range': {
      const { tab, range } = target(host, input.sheet, input.range)
      const rg = clampRange(range, wb.size(tab))
      if (!rg) return { sheet: tab.name, range: rangeName(range), values: [], note: 'Range is empty.' }
      const width = rg.c1 - rg.c0 + 1
      const maxRows = Math.max(1, Math.floor(MAX_CELLS / width))
      const truncated = rg.r1 - rg.r0 + 1 > maxRows
      const read = { ...rg, r1: Math.min(rg.r1, rg.r0 + maxRows - 1) }
      const values = wb.rangeValues(tab, read).map((row) => row.map(plain))
      const formulas: Record<string, string> = {}
      const formats: Record<string, string> = {}
      for (let r = read.r0; r <= read.r1; r++)
        for (let c = read.c0; c <= read.c1; c++) {
          if (wb.isFormula(tab, r, c)) formulas[rangeName({ r0: r, c0: c, r1: r, c1: c })] = String(wb.raw(tab, r, c))
          const fmt = tab.styles[`${r}:${c}`]?.fmt
          if (fmt) formats[rangeName({ r0: r, c0: c, r1: r, c1: c })] = fmt
        }
      return { sheet: tab.name, range: rangeName(read), values, formulas, numberFormats: formats, truncated: truncated ? `Only the first ${maxRows} rows were returned; read the rest in another call.` : undefined }
    }
    case 'write_range': {
      const tab = sheetFor(host, splitSheet(input.start).sheet ?? input.sheet)
      const start = parseCell(splitSheet(input.start).ref)
      if (!start) throw new Error(`"${input.start}" is not a valid cell.`)
      const values: string[][] = (input.values ?? []).map((row: unknown[]) => (Array.isArray(row) ? row : [row]).map((v) => (v === null || v === undefined ? '' : String(v))))
      if (!values.length) throw new Error('values is empty.')
      const width = Math.max(...values.map((r) => r.length))
      const rows = values.map((r) => [...r, ...new Array(width - r.length).fill('')])
      wb.writeTexts(tab, start.r, start.c, rows, 'AI: Write cells')
      const written = { r0: start.r, c0: start.c, r1: start.r + rows.length - 1, c1: start.c + width - 1 }
      // report formula errors so the model can fix them
      const errors: string[] = []
      for (let r = written.r0; r <= written.r1 && errors.length < 10; r++)
        for (let c = written.c0; c <= written.c1 && errors.length < 10; c++) {
          const v = wb.value(tab, r, c)
          if (v instanceof DetailedCellError) errors.push(`${rangeName({ r0: r, c0: c, r1: r, c1: c })}: ${v.value} ${v.message ?? ''}`.trim())
        }
      return { sheet: tab.name, written: rangeName(written), errors: errors.length ? errors : undefined }
    }
    case 'clear_range': {
      const { tab, range } = target(host, input.sheet, input.range)
      wb.clearContents(tab, range)
      return { cleared: `${tab.name}!${rangeName(range)}` }
    }
    case 'format_range': {
      const { tab, range } = target(host, input.sheet, input.range)
      const patch: Partial<CellStyle> = {}
      if (input.bold !== undefined) patch.bold = input.bold
      if (input.italic !== undefined) patch.italic = input.italic
      if (input.underline !== undefined) patch.underline = input.underline
      if (input.fontSize) patch.size = Math.round(Number(input.fontSize) * (4 / 3))
      if (input.fontColor) patch.color = input.fontColor
      if (input.fillColor) patch.fill = input.fillColor === 'none' ? undefined : input.fillColor
      if (input.horizontalAlign) patch.align = input.horizontalAlign
      if (input.wrap !== undefined) patch.wrap = input.wrap
      if (input.numberFormat) patch.fmt = input.numberFormat === 'General' ? undefined : input.numberFormat
      if (Object.keys(patch).length) wb.applyStyle(tab, range, patch, 'AI: Format')
      if (input.borders) wb.setBorders(tab, range, input.borders)
      return { formatted: `${tab.name}!${rangeName(range)}` }
    }
    case 'set_column_width': {
      const tab = sheetFor(host, input.sheet)
      const [a, b = a] = String(input.columns).toUpperCase().split(':')
      const c0 = colIndex(a)
      const c1 = colIndex(b)
      const cols = Array.from({ length: Math.abs(c1 - c0) + 1 }, (_, i) => Math.min(c0, c1) + i)
      if (input.width === 'auto') {
        const { rows } = wb.size(tab)
        for (const c of cols) {
          let max = 4
          for (let r = 0; r < Math.min(rows, 2000); r++) max = Math.max(max, wb.display(tab, r, c).length)
          wb.setColWidth(tab, [c], Math.min(420, max * 7.5 + 18))
        }
      } else wb.setColWidth(tab, cols, Number(input.width))
      return { ok: true }
    }
    case 'create_sheet': {
      const tab = wb.addSheet(String(input.name))
      host.setActive(tab.id)
      return { created: tab.name }
    }
    case 'sort_range': {
      const { tab, range } = target(host, input.sheet, input.range)
      const col = colIndex(String(input.column).toUpperCase())
      if (col < range.c0 || col > range.c1) throw new Error('The sort column must be inside the range.')
      wb.sortRange(tab, range, col, !!input.descending, input.hasHeader ?? true)
      return { sorted: `${tab.name}!${rangeName(range)}` }
    }
    case 'freeze_panes': {
      const tab = sheetFor(host, input.sheet)
      wb.setFreeze(tab, input.rows, input.columns)
      return { ok: true }
    }
    case 'create_chart': {
      const range = qualify(host, String(input.range))
      if (!wb.refValues(range)) throw new Error(`Range "${range}" not found.`)
      const { dashboard, widget } = wb.addWidget(input.dashboard || 'Dashboard', {
        type: 'chart',
        w: ['pie', 'doughnut', 'funnel', 'radar'].includes(input.chartType) ? 4 : 6,
        h: 5,
        title: input.title,
        chart: { type: input.chartType as ChartType, range, seriesBy: 'columns', headers: true, legend: true }
      })
      return { dashboard: dashboard.name, widgetId: widget.id }
    }
    case 'add_kpi_card': {
      const formula = String(input.formula).startsWith('=') ? String(input.formula) : `=${input.formula}`
      const value = plain(wb.evaluate(formula))
      if (typeof value === 'string' && value.startsWith('#')) throw new Error(`The formula evaluates to ${value}. Use explicit sheet references like =SUM(Sheet1!B2:B10).`)
      const { dashboard } = wb.addWidget(input.dashboard || 'Dashboard', { type: 'kpi', w: 3, h: 2, title: input.label, kpi: { formula, fmt: input.numberFormat } })
      return { dashboard: dashboard.name, value }
    }
    default:
      throw new Error(`Unknown tool ${name}`)
  }
}

/** Short description of the workbook sent with every chat message. */
export function workbookContext(wb: Workbook, active: GridSheet, sel: Range | null): string {
  const lines: string[] = []
  lines.push(`Active sheet: ${active.name}`)
  if (sel) lines.push(`Selection: ${rangeName(sel)}`)
  lines.push(
    `Tabs: ${wb.tabs
      .map((t) => {
        if (t.kind === 'dashboard') return `${t.name} (dashboard, ${t.widgets.length} widgets)`
        const { rows, cols } = wb.size(t)
        return `${t.name} (${rows && cols ? rangeName({ r0: 0, c0: 0, r1: rows - 1, c1: cols - 1 }) : 'empty'})`
      })
      .join(', ')}`
  )
  // A small preview of the selection (or the top of the sheet) saves the model a tool call.
  const { rows, cols } = wb.size(active)
  const base = sel && !(sel.r0 === sel.r1 && sel.c0 === sel.c1) ? sel : { r0: 0, c0: 0, r1: rows - 1, c1: cols - 1 }
  const prev = clampRange({ ...base, r1: Math.min(base.r1, base.r0 + 14), c1: Math.min(base.c1, base.c0 + 9) }, { rows, cols })
  if (prev) {
    lines.push(`Preview of ${active.name}!${rangeName(prev)} (tab-separated):`)
    for (let r = prev.r0; r <= prev.r1; r++) {
      const row: string[] = []
      for (let c = prev.c0; c <= prev.c1; c++) row.push(wb.display(active, r, c))
      lines.push(row.join('\t'))
    }
  }
  return lines.join('\n')
}
