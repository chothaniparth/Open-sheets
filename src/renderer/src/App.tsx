import { useCallback, useEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import { Settings as SettingsIcon, Sparkles } from 'lucide-react'
import type { RawCellContent } from 'hyperformula'
import { FILE_EXTENSIONS, type SettingsSnapshot } from '@shared/types'
import { MAX_COLS, MAX_ROWS, cellName, normRange, parseCell, parseRange, quoteSheet, rangeName, shiftFormula, splitSheet, type Range } from './engine/addr'
import { adjustDecimals } from './engine/format'
import { clampRange, styleKey, Workbook, type CellStyle, type GridSheet, type Widget } from './engine/workbook'
import { exportWorkbook, sheetToHtml } from './io/exporter'
import { importFile } from './io/importer'
import { runTool, workbookContext, type ToolHost } from './ai/tools'
import { AIPanel } from './components/AIPanel'
import { FormulaBar, SheetTabs, StatusBar } from './components/Chrome'
import { DashboardView } from './components/Dashboard'
import { ChartDialog, CondFormatDialog, FilterPopup, FindBar, FunctionDialog, KpiDialog, PromptDialog, TableDialog, type WidgetSave } from './components/Dialogs'
import { Grid } from './components/Grid'
import { Ribbon, type RibbonTab } from './components/Ribbon'
import { SettingsDialog } from './components/Settings'
import type { ClipboardData, EditState, Sel } from './components/types'

type DialogState =
  | { kind: 'chart'; range: string; widget?: Widget; dashId?: string }
  | { kind: 'kpi'; formula: string; widget?: Widget; dashId?: string }
  | { kind: 'table'; range: string; widget?: Widget; dashId?: string }
  | { kind: 'cond' }
  | { kind: 'function' }
  | { kind: 'numfmt'; initial: string }
  | { kind: 'settings'; section?: 'providers' | 'general' | 'about' }

const WRITABLE = new Set<string>(FILE_EXTENSIONS.save)
const basename = (p: string) => p.split('/').pop() ?? p
const extOf = (p: string) => p.split('.').pop()?.toLowerCase() ?? ''
const DEFAULT_SEL: Sel = { anchor: { r: 0, c: 0 }, focus: { r: 0, c: 0 } }

/** Parse clipboard text (tab-separated, with Excel-style quoted fields). */
function parseTsv(text: string): string[][] {
  const rows: string[][] = []
  let row: string[] = []
  let cell = ''
  let quoted = false
  const t = text.replace(/\r\n?/g, '\n').replace(/\n$/, '')
  for (let i = 0; i < t.length; i++) {
    const ch = t[i]
    if (quoted) {
      if (ch === '"' && t[i + 1] === '"') (cell += '"'), i++
      else if (ch === '"') quoted = false
      else cell += ch
    } else if (ch === '"' && cell === '') quoted = true
    else if (ch === '\t') row.push(cell), (cell = '')
    else if (ch === '\n') row.push(cell), rows.push(row), (row = []), (cell = '')
    else cell += ch
  }
  row.push(cell)
  rows.push(row)
  return rows
}

function useDark(theme: string | undefined) {
  const [sys, setSys] = useState(() => matchMedia('(prefers-color-scheme: dark)').matches)
  useEffect(() => {
    const mq = matchMedia('(prefers-color-scheme: dark)')
    const fn = () => setSys(mq.matches)
    mq.addEventListener('change', fn)
    return () => mq.removeEventListener('change', fn)
  }, [])
  useEffect(() => {
    if (theme === 'light' || theme === 'dark') document.documentElement.dataset.theme = theme
    else delete document.documentElement.dataset.theme
  }, [theme])
  return theme === 'dark' || (theme !== 'light' && sys)
}

export default function App() {
  const [wb, setWb] = useState(() => Workbook.create())
  const version = useSyncExternalStore(wb.subscribe, wb.getVersion)
  const [doc, setDoc] = useState<{ path?: string; name: string }>({ name: 'Untitled' })
  const [activeId, setActiveId] = useState(wb.tabs[0].id)
  const [sels, setSels] = useState<Record<string, Sel>>({})
  const [edit, setEditState] = useState<EditState | null>(null)
  const [ribbonTab, setRibbonTab] = useState<RibbonTab>('home')
  const [showAI, setShowAI] = useState(true)
  const [showFormulas, setShowFormulas] = useState(false)
  const [gridlines, setGridlines] = useState(true)
  const [snap, setSnap] = useState<SettingsSnapshot | null>(null)
  const [dialog, setDialog] = useState<DialogState | null>(null)
  const [find, setFind] = useState(false)
  const [filterPop, setFilterPop] = useState<{ col: number; x: number; y: number } | null>(null)
  const [clip, setClip] = useState<ClipboardData | null>(null)
  const [toast, setToastText] = useState<string | null>(null)
  const [loading, setLoading] = useState<string | null>(null)
  const [focusToken, setFocusToken] = useState(0)
  const lastGridId = useRef(activeId)
  const dark = useDark(snap?.settings.theme)

  const tab = wb.tab(activeId) ?? wb.tabs[0]
  if (tab.kind === 'grid') lastGridId.current = tab.id
  const grid = (tab.kind === 'grid' ? tab : (wb.tab(lastGridId.current) ?? wb.grids()[0])) as GridSheet
  const sel = sels[grid.id] ?? DEFAULT_SEL
  const range = normRange(sel.anchor, sel.focus)
  const functions = useMemo(() => wb.hf.getRegisteredFunctionNames().sort(), [wb])

  // live refs for callbacks registered once (menu, AI tools)
  const editRef = useRef(edit)
  editRef.current = edit
  const wbRef = useRef(wb)
  wbRef.current = wb
  const gridRef = useRef(grid)
  gridRef.current = grid
  const rangeRef = useRef<Range | null>(range)
  rangeRef.current = tab.kind === 'grid' ? range : null

  const setEdit = useCallback((e: EditState | null) => setEditState(e), [])
  const setSel = useCallback((s: Sel, tabId?: string) => setSels((prev) => ({ ...prev, [tabId ?? gridRef.current.id]: s })), [])

  const toastTimer = useRef<ReturnType<typeof setTimeout>>(undefined)
  const showToast = (text: string) => {
    setToastText(text)
    clearTimeout(toastTimer.current)
    toastTimer.current = setTimeout(() => setToastText(null), 3200)
  }
  const fail = (err: unknown) => showToast((err as Error).message ?? String(err))

  // ---------------------------------------------------------------- settings
  useEffect(() => {
    void window.api.settings.get().then((s) => {
      setSnap(s)
      setGridlines(s.settings.showGridlines)
    })
    return window.api.settings.onChange(setSnap)
  }, [])

  // ---------------------------------------------------------------- documents
  const replaceWorkbook = (next: Workbook, info: { path?: string; name: string }) => {
    const old = wbRef.current
    setWb(next)
    setDoc(info)
    setActiveId(next.tabs[0].id)
    setSels({})
    setEditState(null)
    setClip(null)
    old.destroy()
  }

  useEffect(
    () =>
      window.api.file.onLoad(async (file) => {
        setLoading(`Opening ${file.name}…`)
        await new Promise((r) => setTimeout(r, 30))
        try {
          const data = await importFile(file.name, file.data)
          replaceWorkbook(Workbook.fromData(data), { path: WRITABLE.has(extOf(file.name)) ? file.path : undefined, name: file.name })
          if (['xlsx', 'xlsm', 'xltx', 'xltm'].includes(extOf(file.name)) === false) showToast(`Opened ${file.name}. Save as .xlsx to keep formatting and dashboards.`)
        } catch (err) {
          showToast(`Could not open ${file.name}: ${(err as Error).message}`)
        } finally {
          setLoading(null)
        }
      }),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    []
  )

  useEffect(() => {
    void window.api.win.setState({ title: doc.name, edited: wb.dirty })
  }, [doc.name, wb.dirty, version])

  const writeTo = async (path: string, asDocument: boolean) => {
    const ext = extOf(path)
    setLoading('Saving…')
    try {
      const bytes = await exportWorkbook(wb, ext, grid)
      await window.api.file.write(path, bytes, { isDocument: asDocument })
      if (asDocument) {
        setDoc({ path, name: basename(path) })
        wb.markSaved()
      }
      showToast(`Saved ${basename(path)}`)
      return true
    } catch (err) {
      fail(err)
      return false
    } finally {
      setLoading(null)
    }
  }

  const saveAs = async (ext?: string): Promise<boolean> => {
    const base = doc.name.replace(/\.[^.]+$/, '')
    const path = await window.api.file.saveDialog(`${base}.${ext ?? 'xlsx'}`, ext ? [ext] : undefined)
    if (!path) return false
    return writeTo(path, !ext)
  }

  const save = async (): Promise<boolean> => {
    if (!doc.path) return saveAs()
    const ext = extOf(doc.path)
    if (ext === 'xlsx') return writeTo(doc.path, true)
    if (WRITABLE.has(ext)) {
      const choice = await window.api.app.confirm(
        `Keep using the .${ext} format?`,
        ['Keep Format', 'Save as .xlsx…', 'Cancel'],
        'Formatting, charts and dashboards are only fully preserved in the Excel Workbook (.xlsx) format.'
      )
      if (choice === 0) return writeTo(doc.path, true)
      if (choice === 1) return saveAs()
      return false
    }
    return saveAs()
  }

  const printSheet = (mode: 'print' | 'pdf') => {
    void window.api.app.printHtml(sheetToHtml(wb, grid, true), mode, `${doc.name.replace(/\.[^.]+$/, '')} - ${grid.name}.pdf`).catch(fail)
  }

  // ---------------------------------------------------------------- editing
  const startEdit = (initial?: string) => {
    if (tab.kind !== 'grid') return
    const { r, c } = sel.anchor
    const original = wb.editText(grid, r, c)
    setEditState({ tabId: grid.id, r, c, text: initial ?? original, original, origin: 'cell', mode: initial !== undefined ? 'enter' : 'edit' })
  }

  const commitEdit = (move?: [number, number]) => {
    const e = editRef.current
    if (!e) return
    const t = wb.tab(e.tabId)
    setEditState(null)
    if (t?.kind === 'grid' && e.text !== e.original) {
      try {
        wb.setInput(t, e.r, e.c, e.text)
      } catch (err) {
        fail(err)
      }
    }
    if (e.tabId !== activeId) setActiveId(e.tabId)
    if (move && (move[0] || move[1])) {
      const m = t?.kind === 'grid' ? wb.mergeAt(t, e.r, e.c) : undefined
      const r = Math.max(0, Math.min(MAX_ROWS - 1, (move[0] > 0 && m ? m.r1 : e.r) + move[0]))
      const c = Math.max(0, Math.min(MAX_COLS - 1, (move[1] > 0 && m ? m.c1 : e.c) + move[1]))
      setSel({ anchor: { r, c }, focus: { r, c } }, e.tabId)
    }
    setFocusToken((x) => x + 1)
  }

  const cancelEdit = () => {
    const e = editRef.current
    setEditState(null)
    if (e && e.tabId !== activeId) setActiveId(e.tabId)
    setFocusToken((x) => x + 1)
  }

  const selectTab = (id: string) => {
    const e = editRef.current
    // keep a formula edit open so cells on other sheets can be referenced
    if (e && !(e.text.startsWith('=') && wb.tab(id)?.kind === 'grid')) commitEdit()
    setActiveId(id)
  }

  // ---------------------------------------------------------------- clipboard
  const pasteInternal = (c: ClipboardData, valuesOnly: boolean) => {
    const src = wb.tab(c.tabId)
    if (!src || src.kind !== 'grid') return
    const h = c.range.r1 - c.range.r0 + 1
    const w = c.range.c1 - c.range.c0 + 1
    // tile the copied block when the selection is an exact multiple of it
    const tilesR = (range.r1 - range.r0 + 1) % h === 0 && range.r1 < MAX_ROWS - 1 ? (range.r1 - range.r0 + 1) / h : 1
    const tilesC = (range.c1 - range.c0 + 1) % w === 0 && range.c1 < MAX_COLS - 1 ? (range.c1 - range.c0 + 1) / w : 1
    const raws = wb.rangeRaw(src, c.range)
    const vals = valuesOnly ? wb.rangeValues(src, c.range) : null
    const out: RawCellContent[][] = []
    const styles: (CellStyle | undefined)[][] = []
    for (let tr = 0; tr < tilesR * h; tr++) {
      const row: RawCellContent[] = []
      const srow: (CellStyle | undefined)[] = []
      for (let tc = 0; tc < tilesC * w; tc++) {
        const sr = tr % h
        const sc = tc % w
        let v: RawCellContent = vals ? (vals[sr][sc] as RawCellContent) : raws[sr][sc]
        if (vals && typeof v === 'object' && v !== null) v = String((v as { value?: string }).value ?? '')
        if (!vals && !c.cut && typeof v === 'string' && v.startsWith('=')) {
          v = shiftFormula(v, range.r0 + tr - (c.range.r0 + sr), range.c0 + tc - (c.range.c0 + sc))
        }
        row.push(v)
        srow.push(src.styles[styleKey(c.range.r0 + sr, c.range.c0 + sc)])
      }
      out.push(row)
      styles.push(srow)
    }
    const dest = grid
    wb.transact(
      c.cut ? 'Move' : 'Paste',
      () => {
        if (c.cut) {
          const empty = raws.map((r) => r.map(() => null))
          wb.hf.setCellContents({ sheet: wb.sid(src), row: c.range.r0, col: c.range.c0 }, empty)
          for (let r = c.range.r0; r <= c.range.r1; r++) for (let cc = c.range.c0; cc <= c.range.c1; cc++) delete src.styles[styleKey(r, cc)]
        }
        wb.hf.setCellContents({ sheet: wb.sid(dest), row: range.r0, col: range.c0 }, out)
        if (!valuesOnly)
          styles.forEach((row, i) =>
            row.forEach((st, j) => {
              const k = styleKey(range.r0 + i, range.c0 + j)
              if (st) dest.styles[k] = st
              else delete dest.styles[k]
            })
          )
      },
      { hf: true }
    )
    setSel({ anchor: { r: range.r0, c: range.c0 }, focus: { r: range.r0 + out.length - 1, c: range.c0 + out[0].length - 1 } })
    if (c.cut) setClip(null)
  }

  const pasteText = (text: string, valuesOnly = false) => {
    if (tab.kind !== 'grid') return
    if (clip && text === clip.tsv) return pasteInternal(clip, valuesOnly)
    const rows = parseTsv(text)
    if (!rows.length) return
    wb.writeTexts(grid, range.r0, range.c0, rows, 'Paste')
    setSel({ anchor: { r: range.r0, c: range.c0 }, focus: { r: range.r0 + rows.length - 1, c: range.c0 + Math.max(...rows.map((r) => r.length)) - 1 } })
  }

  const copySelection = (cut: boolean): ClipboardData | null => {
    const rg = clampRange(range, wb.size(grid)) ?? { ...range, r1: range.r0, c1: range.c0 }
    const lines: string[] = []
    for (let r = rg.r0; r <= rg.r1; r++) {
      const cells: string[] = []
      for (let c = rg.c0; c <= rg.c1; c++) {
        const s = wb.display(grid, r, c)
        cells.push(/[\t\n"]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s)
      }
      lines.push(cells.join('\t'))
    }
    const data = { tabId: grid.id, range: rg, tsv: lines.join('\n'), cut }
    setClip(data)
    return data
  }

  const onClipboard = (kind: 'copy' | 'cut' | 'paste', e: ClipboardEvent) => {
    e.preventDefault()
    if (kind === 'paste') return pasteText(e.clipboardData?.getData('text/plain') ?? '')
    const data = copySelection(kind === 'cut')
    if (data) e.clipboardData?.setData('text/plain', data.tsv)
  }

  // ---------------------------------------------------------------- helpers
  const region = (): Range => (range.r0 === range.r1 && range.c0 === range.c1 ? wb.currentRegion(grid, range.r0, range.c0) : clampRange(range, wb.size(grid)) ?? range)
  const qualified = (rg: Range) => `${quoteSheet(grid.name)}!${rangeName(rg)}`
  const activeStyle = tab.kind === 'grid' ? wb.style(grid, sel.anchor.r, sel.anchor.c) : undefined
  const rowsCount = Math.min(range.r1 - range.r0 + 1, 1000)
  const colsCount = Math.min(range.c1 - range.c0 + 1, 200)

  const saveWidget = (s: WidgetSave, existing?: Widget, dashId?: string) => {
    if (existing && dashId) {
      wb.updateDashboard(dashId, (d) => {
        d.widgets = d.widgets.map((w) => (w.id === existing.id ? { ...w, ...s.widget, id: w.id, x: w.x, y: w.y, w: w.w, h: w.h } : w))
      })
    } else {
      const { dashboard } = wb.addWidget(s.dashboard, s.widget)
      setActiveId(dashboard.id)
    }
    setDialog(null)
  }

  const autoSum = () => {
    if (range.r0 !== range.r1 || range.c0 !== range.c1) {
      const rg = clampRange(range, wb.size(grid))
      if (!rg) return
      const row = Array.from({ length: rg.c1 - rg.c0 + 1 }, (_, i) => `=SUM(${rangeName({ r0: rg.r0, c0: rg.c0 + i, r1: rg.r1, c1: rg.c0 + i })})`)
      wb.writeTexts(grid, rg.r1 + 1, rg.c0, [row], 'AutoSum')
      return
    }
    const { r, c } = sel.anchor
    const isNum = (rr: number, cc: number) => rr >= 0 && cc >= 0 && typeof wb.value(grid, rr, cc) === 'number'
    let k = r - 1
    while (isNum(k, c)) k--
    if (k < r - 1) return setEditState({ tabId: grid.id, r, c, text: `=SUM(${rangeName({ r0: k + 1, c0: c, r1: r - 1, c1: c })})`, original: '', origin: 'cell', mode: 'edit' })
    k = c - 1
    while (isNum(r, k)) k--
    const text = k < c - 1 ? `=SUM(${rangeName({ r0: r, c0: k + 1, r1: r, c1: c - 1 })})` : '=SUM()'
    setEditState({ tabId: grid.id, r, c, text, original: '', origin: 'cell', mode: 'edit' })
  }

  // ---------------------------------------------------------------- commands
  const run = (cmd: string, arg?: string) => {
    const isGrid = tab.kind === 'grid'
    const inTextField = () => {
      const el = document.activeElement as HTMLElement | null
      return !!el && (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA' || el.isContentEditable)
    }
    const style = (patch: Partial<CellStyle> | ((old: CellStyle | undefined) => Partial<CellStyle>), label?: string) => {
      if (!isGrid) return
      if (editRef.current) commitEdit()
      wb.applyStyle(grid, range, patch, label)
    }
    try {
      switch (cmd) {
        // file
        case 'save':
          return void save()
        case 'saveAs':
          return void saveAs()
        case 'saveAndClose':
          return void save().then((ok) => ok && window.api.win.close())
        case 'print':
          return printSheet('print')
        case 'export:pdf':
          return printSheet('pdf')
        case 'settings':
          return setDialog({ kind: 'settings' })
        // edit
        case 'undo':
          if (inTextField()) return void document.execCommand('undo')
          setEditState(null)
          return wb.undo()
        case 'redo':
          if (inTextField()) return void document.execCommand('redo')
          return wb.redo()
        case 'copy':
        case 'cut':
          if (!isGrid) return
          setFocusToken((x) => x + 1)
          return void setTimeout(() => document.execCommand(cmd))
        case 'paste':
        case 'pasteValues':
          return void navigator.clipboard
            .readText()
            .then((t) => pasteText(t, cmd === 'pasteValues'))
            .catch(fail)
        case 'selectAll':
          if (inTextField()) return void document.execCommand('selectAll')
          return setSel({ anchor: { r: 0, c: 0 }, focus: { r: MAX_ROWS - 1, c: MAX_COLS - 1 } })
        case 'clear':
          return isGrid && wb.clearContents(grid, range)
        case 'clearFormats':
          return isGrid && wb.clearFormats(grid, range)
        case 'fillDown':
          if (isGrid && range.r1 > range.r0) wb.fill(grid, { ...range, r1: range.r0 }, range)
          return
        case 'fillRight':
          if (isGrid && range.c1 > range.c0) wb.fill(grid, { ...range, c1: range.c0 }, range)
          return
        case 'find':
          return isGrid && setFind(true)
        case 'autosum':
          return isGrid && autoSum()
        // formatting
        case 'format:bold':
          return style({ bold: !activeStyle?.bold }, 'Bold')
        case 'format:italic':
          return style({ italic: !activeStyle?.italic }, 'Italic')
        case 'format:underline':
          return style({ underline: !activeStyle?.underline }, 'Underline')
        case 'format:strike':
          return style({ strike: !activeStyle?.strike }, 'Strikethrough')
        case 'format:wrap':
          return style({ wrap: !activeStyle?.wrap }, 'Wrap Text')
        case 'format:align':
          return style({ align: activeStyle?.align === arg ? undefined : (arg as CellStyle['align']) }, 'Align')
        case 'format:font':
          return style({ font: arg || undefined }, 'Font')
        case 'format:size':
          return style({ size: arg ? Number(arg) : undefined }, 'Font Size')
        case 'format:color':
          return style({ color: arg || undefined }, 'Text Color')
        case 'format:fill':
          return style({ fill: arg || undefined }, 'Fill Color')
        case 'format:numfmt':
          if (arg === 'custom') return setDialog({ kind: 'numfmt', initial: activeStyle?.fmt ?? 'General' })
          return style({ fmt: !arg || arg === 'General' ? undefined : arg }, 'Number Format')
        case 'format:decimals': {
          const sample = wb.value(grid, sel.anchor.r, sel.anchor.c)
          return style((old) => ({ fmt: adjustDecimals(old?.fmt ?? wb.numberFormat(grid, sel.anchor.r, sel.anchor.c), Number(arg) as 1 | -1, typeof sample === 'number' ? sample : undefined) }), 'Decimals')
        }
        case 'format:borders':
          return isGrid && wb.setBorders(grid, range, arg as 'all' | 'outside' | 'bottom' | 'none')
        case 'format:merge':
          if (!isGrid) return
          if (grid.merges.some((m) => m.r0 <= range.r1 && range.r0 <= m.r1 && m.c0 <= range.c1 && range.c0 <= m.c1)) return wb.unmerge(grid, range)
          return wb.merge(grid, range)
        case 'format:conditional':
          return isGrid && setDialog({ kind: 'cond' })
        // insert / delete
        case 'insert:rowAbove':
          return isGrid && wb.insertRows(grid, range.r0, rowsCount)
        case 'insert:rowBelow':
          return isGrid && wb.insertRows(grid, Math.min(range.r1, MAX_ROWS - 2) + 1, rowsCount)
        case 'insert:colLeft':
          return isGrid && wb.insertCols(grid, range.c0, colsCount)
        case 'insert:colRight':
          return isGrid && wb.insertCols(grid, Math.min(range.c1, MAX_COLS - 2) + 1, colsCount)
        case 'delete:rows':
          return isGrid && wb.deleteRows(grid, range.r0, range.r1 - range.r0 + 1)
        case 'delete:cols':
          return isGrid && wb.deleteCols(grid, range.c0, range.c1 - range.c0 + 1)
        case 'insert:sheet': {
          const t = wb.addSheet(undefined, wb.tabs.indexOf(tab) + 1)
          return setActiveId(t.id)
        }
        case 'insert:dashboard': {
          const d = wb.addDashboard()
          return setActiveId(d.id)
        }
        case 'insert:chart':
          return setDialog({ kind: 'chart', range: qualified(region()), dashId: tab.kind === 'dashboard' ? tab.id : undefined })
        case 'insert:kpi': {
          const rg = clampRange(range, wb.size(grid))
          return setDialog({ kind: 'kpi', formula: rg ? `=SUM(${qualified(rg)})` : '', dashId: tab.kind === 'dashboard' ? tab.id : undefined })
        }
        case 'insert:function':
        case 'help:functions':
          return setDialog({ kind: 'function' })
        // data
        case 'data:sortAsc':
        case 'data:sortDesc': {
          if (!isGrid) return
          const rg = region()
          return wb.sortRange(grid, rg, sel.anchor.c, cmd === 'data:sortDesc', wb.looksLikeHeader(grid, rg))
        }
        case 'data:filter':
          if (!isGrid) return
          if (grid.filter) return wb.setFilter(grid, undefined)
          return wb.setFilter(grid, { range: region(), criteria: {} })
        case 'data:removeDuplicates': {
          if (!isGrid) return
          const rg = region()
          const n = wb.removeDuplicates(grid, rg, wb.looksLikeHeader(grid, rg))
          return showToast(n ? `Removed ${n} duplicate row${n > 1 ? 's' : ''}.` : 'No duplicate rows found.')
        }
        // view
        case 'view:ai':
          return setShowAI((v) => !v)
        case 'view:formulas':
          return setShowFormulas((v) => !v)
        case 'view:gridlines':
          return setGridlines((v) => !v)
        case 'view:freezeTop':
          return isGrid && wb.setFreeze(grid, 1, 0)
        case 'view:freezeFirst':
          return isGrid && wb.setFreeze(grid, 0, 1)
        case 'view:freezeSelection':
          return isGrid && wb.setFreeze(grid, sel.anchor.r, sel.anchor.c)
        case 'view:unfreeze':
          return isGrid && wb.setFreeze(grid, 0, 0)
        default:
          if (cmd.startsWith('export:')) return void saveAs(cmd.slice(7))
      }
    } catch (err) {
      fail(err)
    }
  }
  const runRef = useRef(run)
  runRef.current = run
  useEffect(() => window.api.app.onMenu((c) => runRef.current(c)), [])

  // ---------------------------------------------------------------- AI tool bridge
  const hostRef = useRef<ToolHost>(null!)
  hostRef.current = {
    wb: () => wbRef.current,
    activeGrid: () => gridRef.current,
    selection: () => rangeRef.current,
    setActive: (id) => setActiveId(id)
  }
  useEffect(
    () =>
      window.api.ai.onToolRequest(async (req) => {
        try {
          const result = await runTool(hostRef.current, req.name, (req.input ?? {}) as Record<string, unknown>)
          window.api.ai.respondTool(req.reqId, true, JSON.parse(JSON.stringify(result ?? null)))
        } catch (err) {
          window.api.ai.respondTool(req.reqId, false, undefined, (err as Error).message)
        }
      }),
    []
  )

  // ---------------------------------------------------------------- context menus
  const cellMenu = async (kind: 'cell' | 'row' | 'col' | 'corner') => {
    const items = [
      { id: 'cut', label: 'Cut' },
      { id: 'copy', label: 'Copy' },
      { id: 'paste', label: 'Paste' },
      { id: 'pasteValues', label: 'Paste Values' },
      { type: 'separator' as const },
      ...(kind !== 'col' ? [{ id: 'insert:rowAbove', label: `Insert ${rowsCount > 1 ? `${rowsCount} Rows` : 'Row'} Above` }, { id: 'insert:rowBelow', label: 'Insert Row Below' }] : []),
      ...(kind !== 'row' ? [{ id: 'insert:colLeft', label: `Insert ${colsCount > 1 ? `${colsCount} Columns` : 'Column'} Left` }, { id: 'insert:colRight', label: 'Insert Column Right' }] : []),
      ...(kind !== 'col' ? [{ id: 'delete:rows', label: 'Delete Rows' }] : []),
      ...(kind !== 'row' ? [{ id: 'delete:cols', label: 'Delete Columns' }] : []),
      { type: 'separator' as const },
      { id: 'clear', label: 'Clear Contents' },
      { id: 'clearFormats', label: 'Clear Formats' },
      { type: 'separator' as const },
      { id: 'data:sortAsc', label: 'Sort A → Z' },
      { id: 'data:sortDesc', label: 'Sort Z → A' },
      { id: 'data:filter', label: grid.filter ? 'Remove Filter' : 'Filter' },
      { type: 'separator' as const },
      { id: 'insert:chart', label: 'Insert Chart…' },
      { id: 'format:conditional', label: 'Conditional Formatting…' },
      { id: 'ai', label: 'Ask AI about this…' }
    ]
    const chosen = await window.api.app.contextMenu(items)
    if (chosen === 'ai') setShowAI(true)
    else if (chosen) run(chosen)
  }

  const tabMenu = async (id: string) => {
    const t = wb.tab(id)
    if (!t) return
    const colors: [string, string][] = [
      ['Blue', '#2a78d6'],
      ['Orange', '#eb6834'],
      ['Green', '#1baf7a'],
      ['Violet', '#6250d6'],
      ['Red', '#e34948']
    ]
    const chosen = await window.api.app.contextMenu([
      { id: 'duplicate', label: 'Duplicate' },
      { id: 'delete', label: 'Delete', enabled: !(t.kind === 'grid' && wb.grids().length === 1) },
      { type: 'separator' },
      ...(t.kind === 'grid' ? [...colors.map(([n]) => ({ id: `color:${n}`, label: `Tab Color: ${n}` })), { id: 'color:none', label: 'No Tab Color' }] : [])
    ])
    try {
      if (chosen === 'duplicate') {
        const copy = wb.duplicateTab(id)
        if (copy) setActiveId(copy.id)
      } else if (chosen === 'delete') {
        const ok = (await window.api.app.confirm(`Delete "${t.name}"?`, ['Delete', 'Cancel'], 'You can undo this with ⌘Z.')) === 0
        if (ok) {
          const idx = wb.tabs.indexOf(t)
          wb.deleteTab(id)
          setActiveId(wb.tabs[Math.max(0, idx - 1)].id)
        }
      } else if (chosen?.startsWith('color:') && t.kind === 'grid') {
        const c = colors.find(([n]) => `color:${n}` === chosen)?.[1]
        wb.transact('Tab Color', () => {
          t.tabColor = c
        })
      }
    } catch (err) {
      fail(err)
    }
  }

  // ---------------------------------------------------------------- render
  const formulaValue = tab.kind === 'grid' ? (showFormulas ? String(wb.raw(grid, sel.anchor.r, sel.anchor.c) ?? '') : wb.editText(grid, sel.anchor.r, sel.anchor.c)) : ''
  const nameBox = tab.kind !== 'grid' ? '' : range.r0 === range.r1 && range.c0 === range.c1 ? cellName(range.r0, range.c0) : rangeName({ ...range, r1: Math.min(range.r1, MAX_ROWS - 1) })

  return (
    <div
      className="app"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => {
        e.preventDefault()
        const f = e.dataTransfer.files[0]
        if (f) void window.api.file.openDropped(f)
      }}
    >
      <div className="titlebar">
        <div className="ribbon-tabs">
          {(['home', 'insert', 'data', 'view'] as RibbonTab[]).map((t) => (
            <button key={t} className={`ribbon-tab${ribbonTab === t ? ' active' : ''}`} onClick={() => setRibbonTab(t)}>
              {t[0].toUpperCase() + t.slice(1)}
            </button>
          ))}
        </div>
        <div className="doc-title">
          {doc.name}
          {wb.dirty && <span className="edited"> — Edited</span>}
        </div>
        <button className={`tbtn${showAI ? ' on' : ''}`} title="AI assistant (⌘J)" onClick={() => setShowAI(!showAI)}>
          <Sparkles size={15} />
          <span className="lbl">AI</span>
        </button>
        <button className="tbtn" title="Settings (⌘,)" onClick={() => setDialog({ kind: 'settings' })}>
          <SettingsIcon size={15} />
        </button>
      </div>

      <Ribbon tab={ribbonTab} style={activeStyle} canUndo={wb.canUndo()} canRedo={wb.canRedo()} isGrid={tab.kind === 'grid'} showFormulas={showFormulas} showGridlines={gridlines} run={run} />

      <FormulaBar
        nameBox={nameBox}
        value={formulaValue}
        edit={edit}
        editingElsewhere={!!edit && edit.tabId !== tab.id}
        functions={functions}
        disabled={tab.kind !== 'grid'}
        onGoto={(ref) => {
          const { sheet, ref: r } = splitSheet(ref)
          const target = sheet ? wb.findTab(sheet) : grid
          const rg = target?.kind === 'grid' ? (parseRange(r) ?? (parseCell(r) && { r0: parseCell(r)!.r, c0: parseCell(r)!.c, r1: parseCell(r)!.r, c1: parseCell(r)!.c })) : null
          if (!target || !rg) return showToast(`"${ref}" is not a valid reference.`)
          setActiveId(target.id)
          setSel({ anchor: { r: rg.r0, c: rg.c0 }, focus: { r: rg.r1, c: rg.c1 } }, target.id)
          setFocusToken((x) => x + 1)
        }}
        onFocus={() => {
          if (tab.kind !== 'grid') return
          if (!editRef.current) {
            const { r, c } = sel.anchor
            const original = wb.editText(grid, r, c)
            setEditState({ tabId: grid.id, r, c, text: original, original, origin: 'bar', mode: 'edit' })
          } else if (editRef.current.origin !== 'bar') setEditState({ ...editRef.current, origin: 'bar' })
        }}
        onChange={(text) => editRef.current && setEditState({ ...editRef.current, text, pick: undefined })}
        onCommit={(move) => commitEdit(move)}
        onCancel={cancelEdit}
      />

      <div className="workspace">
        <div className="sheet-area">
          {tab.kind === 'grid' ? (
            <Grid
              wb={wb}
              tab={grid}
              version={version}
              sel={sel}
              setSel={(s) => setSel(s)}
              edit={edit}
              setEdit={setEdit}
              startEdit={startEdit}
              commitEdit={commitEdit}
              cancelEdit={cancelEdit}
              showFormulas={showFormulas}
              showGridlines={gridlines}
              copyRange={clip && clip.tabId === grid.id ? clip.range : null}
              functions={functions}
              defaultFont={snap?.settings.defaultFont ?? 'Helvetica Neue'}
              defaultFontSize={snap?.settings.defaultFontSize ?? 13}
              onContextMenu={(k) => void cellMenu(k)}
              onFill={(src, dst) => wb.fill(grid, src, dst)}
              onFilterButton={(col, x, y) => setFilterPop({ col, x, y })}
              onClipboard={onClipboard}
              onKeyCommand={(c) => (c === 'clear' ? wb.clearContents(grid, range) : setClip(null))}
              focusToken={focusToken}
            />
          ) : (
            <DashboardView
              key={tab.id}
              wb={wb}
              dash={tab}
              version={version}
              dark={dark}
              onAdd={(type) => {
                if (type === 'text') return wb.addWidget(tab.name, { type: 'text', w: 4, h: 3, text: '' })
                const rg = clampRange(range, wb.size(grid))
                if (type === 'chart') setDialog({ kind: 'chart', range: qualified(region()), dashId: tab.id })
                if (type === 'kpi') setDialog({ kind: 'kpi', formula: rg ? `=SUM(${qualified(rg)})` : '', dashId: tab.id })
                if (type === 'table') setDialog({ kind: 'table', range: qualified(region()), dashId: tab.id })
              }}
              onEdit={(w) => {
                if (w.type === 'chart') setDialog({ kind: 'chart', range: w.chart!.range, widget: w, dashId: tab.id })
                if (w.type === 'kpi') setDialog({ kind: 'kpi', formula: w.kpi!.formula, widget: w, dashId: tab.id })
                if (w.type === 'table') setDialog({ kind: 'table', range: w.table!.range, widget: w, dashId: tab.id })
              }}
            />
          )}
          {find && tab.kind === 'grid' && (
            <FindBar
              wb={wb}
              tab={grid}
              from={sel.anchor}
              onSelect={(c) => setSel({ anchor: c, focus: c })}
              onClose={() => {
                setFind(false)
                setFocusToken((x) => x + 1)
              }}
            />
          )}
          {filterPop && grid.filter && <FilterPopup wb={wb} tab={grid} col={filterPop.col} x={filterPop.x} y={filterPop.y} onClose={() => setFilterPop(null)} />}
          {loading && <div className="loading">{loading}</div>}
        </div>
        {showAI && (
          <AIPanel
            snap={snap}
            getContext={() => workbookContext(wbRef.current, gridRef.current, rangeRef.current)}
            onOpenSettings={() => setDialog({ kind: 'settings', section: 'providers' })}
            onClose={() => setShowAI(false)}
          />
        )}
      </div>

      <div>
        <SheetTabs
          tabs={wb.tabs}
          activeId={tab.id}
          onSelect={selectTab}
          onAdd={() => run('insert:sheet')}
          onRename={(id, name) => {
            try {
              wb.renameTab(id, name)
            } catch (err) {
              fail(err)
            }
          }}
          onMove={(id, idx) => {
            const from = wb.tabs.findIndex((t) => t.id === id)
            wb.moveTab(id, from < idx ? idx - 1 : idx)
          }}
          onContext={(id) => void tabMenu(id)}
        />
        <StatusBar wb={wb} tab={tab.kind === 'grid' ? grid : null} range={tab.kind === 'grid' ? range : null} message={edit ? (edit.text.startsWith('=') ? 'Point' : 'Enter') : undefined} />
      </div>

      {dialog?.kind === 'chart' && (
        <ChartDialog
          wb={wb}
          dark={dark}
          initialRange={dialog.range}
          widget={dialog.widget}
          preferredDashboard={dialog.dashId ? wb.tab(dialog.dashId)?.name : undefined}
          onSave={(s) => saveWidget(s, dialog.widget, dialog.dashId)}
          onClose={() => setDialog(null)}
        />
      )}
      {dialog?.kind === 'kpi' && (
        <KpiDialog
          wb={wb}
          widget={dialog.widget}
          defaultFormula={dialog.formula}
          preferredDashboard={dialog.dashId ? wb.tab(dialog.dashId)?.name : undefined}
          onSave={(s) => saveWidget(s, dialog.widget, dialog.dashId)}
          onClose={() => setDialog(null)}
        />
      )}
      {dialog?.kind === 'table' && (
        <TableDialog
          wb={wb}
          widget={dialog.widget}
          initialRange={dialog.range}
          preferredDashboard={dialog.dashId ? wb.tab(dialog.dashId)?.name : undefined}
          onSave={(s) => saveWidget(s, dialog.widget, dialog.dashId)}
          onClose={() => setDialog(null)}
        />
      )}
      {dialog?.kind === 'cond' && <CondFormatDialog wb={wb} tab={grid} selection={range} onClose={() => setDialog(null)} />}
      {dialog?.kind === 'function' && (
        <FunctionDialog
          functions={functions}
          onClose={() => setDialog(null)}
          onInsert={(name) => {
            setDialog(null)
            if (tab.kind !== 'grid') return
            const { r, c } = sel.anchor
            setEditState({ tabId: grid.id, r, c, text: `=${name}(`, original: wb.editText(grid, r, c), origin: 'cell', mode: 'edit' })
          }}
        />
      )}
      {dialog?.kind === 'numfmt' && (
        <PromptDialog
          title="Custom number format"
          label="Excel format code, e.g. #,##0.00;[Red]-#,##0.00 or yyyy-mm-dd"
          initial={dialog.initial}
          onClose={() => setDialog(null)}
          onOk={(v) => {
            setDialog(null)
            wb.applyStyle(grid, range, { fmt: v && v !== 'General' ? v : undefined }, 'Number Format')
          }}
        />
      )}
      {dialog?.kind === 'settings' && snap && <SettingsDialog snap={snap} initial={dialog.section} onClose={() => setDialog(null)} />}
      {toast && <div className="toast">{toast}</div>}
    </div>
  )
}
