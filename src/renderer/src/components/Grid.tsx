import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent, type MouseEvent as ReactMouseEvent } from 'react'
import { DetailedCellError } from 'hyperformula'
import { MAX_COLS, MAX_ROWS, colName, findRefs, inRange, normRange, quoteSheet, rangeName, type CellAddr, type Range } from '../engine/addr'
import { Axis } from '../engine/axis'
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, styleKey, type CellStyle, type GridSheet, type Workbook } from '../engine/workbook'
import { SERIES_LIGHT } from '../charts/options'
import { FormulaInput } from './FormulaInput'
import type { EditState, Sel } from './types'

const HH = 24

export interface GridProps {
  wb: Workbook
  tab: GridSheet
  version: number
  sel: Sel
  setSel: (s: Sel) => void
  edit: EditState | null
  setEdit: (e: EditState | null) => void
  startEdit: (initial?: string) => void
  commitEdit: (move?: [number, number]) => void
  cancelEdit: () => void
  showFormulas: boolean
  showGridlines: boolean
  copyRange: Range | null
  functions: string[]
  defaultFont: string
  defaultFontSize: number
  onContextMenu: (kind: 'cell' | 'row' | 'col' | 'corner') => void
  onFill: (src: Range, dst: Range) => void
  onFilterButton: (col: number, x: number, y: number) => void
  onClipboard: (kind: 'copy' | 'cut' | 'paste', e: ClipboardEvent) => void
  onKeyCommand: (cmd: 'clear' | 'escape') => void
  focusToken: number
}

type Drag =
  | { kind: 'cells' }
  | { kind: 'rows' }
  | { kind: 'cols' }
  | { kind: 'fill'; src: Range; target: Range | null }
  | { kind: 'pick' }
  | { kind: 'resize-col'; index: number; start: number; startSize: number }
  | { kind: 'resize-row'; index: number; start: number; startSize: number }

interface Region {
  r0: number
  r1: number
  c0: number
  c1: number
  sx: number
  sy: number
  clip: [number, number, number, number]
}

const NUMERIC_ALIGN = (v: unknown) => typeof v === 'number'

export function Grid(props: GridProps) {
  const { wb, tab, version, sel, edit, showFormulas, showGridlines, copyRange } = props
  const hostRef = useRef<HTMLDivElement>(null)
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const scrollerRef = useRef<HTMLDivElement>(null)
  const [view, setView] = useState({ w: 800, h: 600 })
  const [scroll, setScroll] = useState({ x: 0, y: 0 })
  const [extent, setExtent] = useState({ rows: 1000, cols: 52 })
  const [preview, setPreview] = useState<{ col?: [number, number]; row?: [number, number]; fill?: Range | null }>({})
  const dragRef = useRef<Drag | null>(null)
  const propsRef = useRef(props)
  propsRef.current = props

  const used = wb.size(tab)
  const rowCount = Math.min(MAX_ROWS, Math.max(extent.rows, used.rows + 100, sel.focus.r + 50, tab.freeze.rows + 50))
  const colCount = Math.min(MAX_COLS, Math.max(extent.cols, used.cols + 10, sel.focus.c + 10, tab.freeze.cols + 20))

  const colSizes = useMemo(() => {
    const s = { ...tab.colWidths }
    if (preview.col) s[preview.col[0]] = preview.col[1]
    return s
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab, version, preview.col])
  const rowSizes = useMemo(() => {
    const s = { ...tab.rowHeights }
    if (preview.row) s[preview.row[0]] = preview.row[1]
    return s
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tab, version, preview.row])
  const cols = useMemo(() => new Axis(DEFAULT_COL_WIDTH, colSizes, {}, colCount), [colSizes, colCount])
  const rows = useMemo(() => new Axis(DEFAULT_ROW_HEIGHT, rowSizes, tab.hiddenRows, rowCount), [rowSizes, tab.hiddenRows, rowCount, version])
  const HW = Math.max(40, String(rowCount).length * 8 + 18)
  const fr = Math.min(tab.freeze.rows, rowCount)
  const fc = Math.min(tab.freeze.cols, colCount)
  const fh = rows.pos(fr)
  const fw = cols.pos(fc)
  const range = normRange(sel.anchor, sel.focus)

  // ---------- geometry ----------
  const cellAt = useCallback(
    (x: number, y: number): CellAddr => {
      // read live scroll offsets so drags stay accurate while autoscrolling
      const sx = scrollerRef.current?.scrollLeft ?? scroll.x
      const sy = scrollerRef.current?.scrollTop ?? scroll.y
      const c = x < HW + fw ? cols.indexAt(x - HW) : cols.indexAt(x - HW + sx)
      const r = y < HH + fh ? rows.indexAt(y - HH) : rows.indexAt(y - HH + sy)
      return { r: Math.max(0, Math.min(r, rowCount - 1)), c: Math.max(0, Math.min(c, colCount - 1)) }
    },
    [HW, fw, fh, cols, rows, scroll, rowCount, colCount]
  )
  const cellRect = useCallback(
    (r: number, c: number, r2 = r, c2 = c) => {
      const sx = c < fc ? 0 : scroll.x
      const sy = r < fr ? 0 : scroll.y
      const x = HW + cols.pos(c) - sx
      const y = HH + rows.pos(r) - sy
      return { x, y, w: cols.pos(c2 + 1) - cols.pos(c), h: rows.pos(r2 + 1) - rows.pos(r) }
    },
    [HW, cols, rows, scroll, fc, fr]
  )

  // ---------- size & scroll ----------
  useLayoutEffect(() => {
    const el = hostRef.current!
    const ro = new ResizeObserver(() => {
      const sc = scrollerRef.current!
      setView({ w: sc.clientWidth, h: sc.clientHeight })
    })
    ro.observe(el)
    return () => ro.disconnect()
  }, [])

  const onScroll = () => {
    const sc = scrollerRef.current!
    setScroll({ x: sc.scrollLeft, y: sc.scrollTop })
    // grow the virtual sheet as the user scrolls toward its end
    if (sc.scrollTop + sc.clientHeight > sc.scrollHeight - 400) setExtent((e) => ({ ...e, rows: Math.min(MAX_ROWS, rowCount + 1000) }))
    if (sc.scrollLeft + sc.clientWidth > sc.scrollWidth - 300) setExtent((e) => ({ ...e, cols: Math.min(MAX_COLS, colCount + 26) }))
  }

  // keep the active cell visible
  useEffect(() => {
    const sc = scrollerRef.current
    if (!sc || dragRef.current) return
    const { r, c } = sel.focus
    if (c >= fc) {
      const left = cols.pos(c)
      const right = cols.pos(c + 1)
      const viewL = sc.scrollLeft + fw
      const viewR = sc.scrollLeft + sc.clientWidth - HW
      if (left < viewL) sc.scrollLeft = left - fw
      else if (right > viewR) sc.scrollLeft = right - (sc.clientWidth - HW)
    }
    if (r >= fr) {
      const top = rows.pos(r)
      const bottom = rows.pos(r + 1)
      const viewT = sc.scrollTop + fh
      const viewB = sc.scrollTop + sc.clientHeight - HH
      if (top < viewT) sc.scrollTop = top - fh
      else if (bottom > viewB) sc.scrollTop = bottom - (sc.clientHeight - HH)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [sel.focus.r, sel.focus.c, tab.id])

  useEffect(() => {
    if (!edit) scrollerRef.current?.focus({ preventScroll: true })
  }, [props.focusToken, tab.id]) // eslint-disable-line react-hooks/exhaustive-deps

  // reset scroll when switching sheets
  useLayoutEffect(() => {
    const sc = scrollerRef.current
    if (sc) {
      sc.scrollTop = 0
      sc.scrollLeft = 0
    }
    setScroll({ x: 0, y: 0 })
  }, [tab.id])

  // clipboard events while the grid has focus
  useEffect(() => {
    const handler = (kind: 'copy' | 'cut' | 'paste') => (e: Event) => {
      if (document.activeElement === scrollerRef.current) propsRef.current.onClipboard(kind, e as ClipboardEvent)
    }
    const copy = handler('copy')
    const cut = handler('cut')
    const paste = handler('paste')
    document.addEventListener('copy', copy)
    document.addEventListener('cut', cut)
    document.addEventListener('paste', paste)
    return () => {
      document.removeEventListener('copy', copy)
      document.removeEventListener('cut', cut)
      document.removeEventListener('paste', paste)
    }
  }, [])

  // ---------- drawing ----------
  const fontFor = (st: CellStyle | undefined) => {
    const size = st?.size ?? props.defaultFontSize
    return `${st?.italic ? 'italic ' : ''}${st?.bold ? '600 ' : ''}${size}px "${st?.font ?? props.defaultFont}", -apple-system, sans-serif`
  }

  useLayoutEffect(() => {
    const canvas = canvasRef.current
    if (!canvas) return
    const dpr = window.devicePixelRatio || 1
    const { w, h } = view
    if (canvas.width !== Math.round(w * dpr) || canvas.height !== Math.round(h * dpr)) {
      canvas.width = Math.round(w * dpr)
      canvas.height = Math.round(h * dpr)
      canvas.style.width = `${w}px`
      canvas.style.height = `${h}px`
    }
    const ctx = canvas.getContext('2d')!
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0)
    const css = getComputedStyle(hostRef.current!)
    const C = (n: string) => css.getPropertyValue(n).trim()
    const bg = C('--grid-bg')
    const line = C('--grid-line')
    const text = C('--grid-text')
    const accent = C('--grid-sel-line')

    ctx.fillStyle = bg
    ctx.fillRect(0, 0, w, h)

    const rStart = Math.max(fr, rows.indexAt(fh + scroll.y))
    const rEnd = rows.indexAt(h - HH + scroll.y)
    const cStart = Math.max(fc, cols.indexAt(fw + scroll.x))
    const cEnd = cols.indexAt(w - HW + scroll.x)

    const regions: Region[] = []
    if (fr && fc) regions.push({ r0: 0, r1: fr - 1, c0: 0, c1: fc - 1, sx: 0, sy: 0, clip: [HW, HH, fw, fh] })
    if (fr) regions.push({ r0: 0, r1: fr - 1, c0: cStart, c1: cEnd, sx: scroll.x, sy: 0, clip: [HW + fw, HH, w - HW - fw, fh] })
    if (fc) regions.push({ r0: rStart, r1: rEnd, c0: 0, c1: fc - 1, sx: 0, sy: scroll.y, clip: [HW, HH + fh, fw, h - HH - fh] })
    regions.push({ r0: rStart, r1: rEnd, c0: cStart, c1: cEnd, sx: scroll.x, sy: scroll.y, clip: [HW + fw, HH + fh, w - HW - fw, h - HH - fh] })

    const editingHere = edit && edit.tabId === tab.id
    const refs = editingHere && edit.text.startsWith('=') ? findRefs(edit.text).filter((r) => !r.sheet || r.sheet.toLowerCase() === tab.name.toLowerCase()) : []

    for (const R of regions) {
      if (R.clip[2] <= 0 || R.clip[3] <= 0 || R.r1 < R.r0 || R.c1 < R.c0) continue
      const X = (c: number) => HW + cols.pos(c) - R.sx
      const Y = (r: number) => HH + rows.pos(r) - R.sy
      ctx.save()
      ctx.beginPath()
      ctx.rect(...R.clip)
      ctx.clip()

      // merges touching this region
      const merges = tab.merges.filter((m) => m.r1 >= R.r0 && m.r0 <= R.r1 && m.c1 >= R.c0 && m.c0 <= R.c1)
      const inMerge = (r: number, c: number) => merges.some((m) => inRange(m, r, c))

      // 1. fills
      const styleCache = new Map<string, CellStyle | undefined>()
      const getStyle = (r: number, c: number) => {
        const k = styleKey(r, c)
        if (!styleCache.has(k)) {
          const base = tab.styles[k]
          const cond = wb.condStyle(tab, r, c)
          styleCache.set(k, cond ? { ...base, ...cond } : base)
        }
        return styleCache.get(k)
      }
      for (let r = R.r0; r <= R.r1; r++) {
        const rh = rows.size(r)
        if (!rh) continue
        for (let c = R.c0; c <= R.c1; c++) {
          const st = getStyle(r, c)
          if (st?.fill) {
            ctx.fillStyle = st.fill
            ctx.fillRect(X(c), Y(r), cols.size(c), rh)
          }
        }
      }

      // 2. gridlines
      if (showGridlines) {
        ctx.strokeStyle = line
        ctx.lineWidth = 1
        ctx.beginPath()
        for (let c = R.c0; c <= R.c1 + 1; c++) {
          const x = Math.round(X(c)) - 0.5
          ctx.moveTo(x, Y(R.r0))
          ctx.lineTo(x, Y(R.r1 + 1))
        }
        for (let r = R.r0; r <= R.r1 + 1; r++) {
          if (r <= R.r1 && !rows.size(r)) continue
          const y = Math.round(Y(r)) - 0.5
          ctx.moveTo(X(R.c0), y)
          ctx.lineTo(X(R.c1 + 1), y)
        }
        ctx.stroke()
      }

      // 3. merged cells cover gridlines
      for (const m of merges) {
        const st = getStyle(m.r0, m.c0)
        ctx.fillStyle = st?.fill ?? bg
        ctx.fillRect(X(m.c0), Y(m.r0), cols.pos(m.c1 + 1) - cols.pos(m.c0) - 1, rows.pos(m.r1 + 1) - rows.pos(m.r0) - 1)
      }

      // 4. text
      const drawText = (r: number, c: number, x: number, y: number, cw: number, ch: number, spillTo: number) => {
        const v = wb.value(tab, r, c)
        let s: string
        if (showFormulas && wb.isFormula(tab, r, c)) s = String(wb.raw(tab, r, c))
        else s = wb.display(tab, r, c)
        if (!s) return
        const st = getStyle(r, c)
        const size = st?.size ?? props.defaultFontSize
        ctx.font = fontFor(st)
        const isErr = v instanceof DetailedCellError
        ctx.fillStyle = st?.color ?? (isErr ? '#d03b3b' : text)
        const align = st?.align ?? (showFormulas && wb.isFormula(tab, r, c) ? 'left' : NUMERIC_ALIGN(v) ? 'right' : typeof v === 'boolean' || isErr ? 'center' : 'left')
        const pad = 4
        const lineH = Math.round(size * 1.3)
        let lines = [s]
        if (st?.wrap) {
          lines = []
          for (const para of s.split('\n')) {
            let cur = ''
            for (const word of para.split(/(\s+)/)) {
              const next = cur + word
              if (ctx.measureText(next).width > cw - pad * 2 && cur.trim()) {
                lines.push(cur.trimEnd())
                cur = word.trimStart()
              } else cur = next
            }
            lines.push(cur)
          }
        } else if (s.includes('\n')) lines = [s.split('\n')[0]]
        const tw = ctx.measureText(lines[0]).width
        if (!st?.wrap && typeof v === 'number' && !showFormulas && tw > cw - pad && st?.fmt !== '@') {
          lines = ['#'.repeat(Math.max(1, Math.floor((cw - pad) / ctx.measureText('#').width)))]
        }
        const valign = st?.valign ?? 'bottom'
        const blockH = lines.length * lineH
        let ty = valign === 'top' ? y + 3 : valign === 'middle' ? y + (ch - blockH) / 2 : y + ch - blockH - 2
        ctx.save()
        ctx.beginPath()
        const clipW = align === 'left' && !st?.wrap ? Math.max(cw, spillTo - x) : cw
        ctx.rect(x, y, clipW, ch)
        ctx.clip()
        ctx.textBaseline = 'middle'
        for (const ln of lines) {
          const lw = ctx.measureText(ln).width
          const tx = align === 'right' ? x + cw - pad - lw : align === 'center' ? x + (cw - lw) / 2 : x + pad
          const cy = ty + lineH / 2
          ctx.textAlign = 'left'
          ctx.fillText(ln, tx, cy)
          if (st?.underline || st?.strike) {
            ctx.fillRect(tx, st.underline ? cy + size * 0.45 : cy, lw, 1)
          }
          ty += lineH
        }
        ctx.restore()
      }

      for (let r = R.r0; r <= R.r1; r++) {
        const rh = rows.size(r)
        if (!rh) continue
        for (let c = R.c0; c <= R.c1; c++) {
          if (merges.length && inMerge(r, c)) continue
          const cw = cols.size(c)
          if (!cw) continue
          if (editingHere && edit.r === r && edit.c === c) continue
          // left-aligned text spills into empty neighbours, like Excel
          let spillTo = X(c) + cw
          const st = getStyle(r, c)
          if (!st?.wrap && (!st?.align || st.align === 'left')) {
            let k = c + 1
            while (k <= R.c1 + 1 && k < colCount && wb.raw(tab, r, k) === null && !inMerge(r, k)) {
              spillTo += cols.size(k)
              k++
              if (spillTo - X(c) > 2000) break
            }
          }
          drawText(r, c, X(c), Y(r), cw, rh, spillTo)
        }
      }
      for (const m of merges) {
        if (editingHere && edit.r === m.r0 && edit.c === m.c0) continue
        drawText(m.r0, m.c0, X(m.c0), Y(m.r0), cols.pos(m.c1 + 1) - cols.pos(m.c0), rows.pos(m.r1 + 1) - rows.pos(m.r0), 0)
      }

      // 5. borders
      for (let r = R.r0; r <= R.r1; r++) {
        for (let c = R.c0; c <= R.c1; c++) {
          const b = tab.styles[styleKey(r, c)]?.border
          if (!b) continue
          const x0 = Math.round(X(c)) - 0.5
          const y0 = Math.round(Y(r)) - 0.5
          const x1 = Math.round(X(c + 1)) - 0.5
          const y1 = Math.round(Y(r + 1)) - 0.5
          const seg = (color: string | undefined, ax: number, ay: number, bx: number, by: number) => {
            if (!color) return
            ctx.strokeStyle = color
            ctx.beginPath()
            ctx.moveTo(ax, ay)
            ctx.lineTo(bx, by)
            ctx.stroke()
          }
          seg(b.top, x0, y0, x1, y0)
          seg(b.bottom, x0, y1, x1, y1)
          seg(b.left, x0, y0, x0, y1)
          seg(b.right, x1, y0, x1, y1)
        }
      }

      // 6. filter buttons
      if (tab.filter) {
        const fr0 = tab.filter.range.r0
        if (fr0 >= R.r0 && fr0 <= R.r1) {
          for (let c = Math.max(R.c0, tab.filter.range.c0); c <= Math.min(R.c1, tab.filter.range.c1); c++) {
            const bx = X(c + 1) - 18
            const by = Y(fr0) + (rows.size(fr0) - 16) / 2
            const active = tab.filter.criteria[c] !== undefined
            ctx.fillStyle = active ? accent : C('--grid-header')
            ctx.strokeStyle = C('--border-strong')
            ctx.beginPath()
            ctx.roundRect(bx, by, 15, 16, 3)
            ctx.fill()
            ctx.stroke()
            ctx.fillStyle = active ? '#fff' : C('--grid-header-text')
            ctx.beginPath()
            ctx.moveTo(bx + 4, by + 6)
            ctx.lineTo(bx + 11, by + 6)
            ctx.lineTo(bx + 7.5, by + 11)
            ctx.fill()
          }
        }
      }

      // 7. formula reference highlights while editing
      refs.forEach((ref, i) => {
        const col = SERIES_LIGHT[i % SERIES_LIGHT.length]
        const x = X(ref.range.c0)
        const y = Y(ref.range.r0)
        ctx.strokeStyle = col
        ctx.lineWidth = 2
        ctx.strokeRect(x + 1, y + 1, X(ref.range.c1 + 1) - x - 2, Y(ref.range.r1 + 1) - y - 2)
        ctx.fillStyle = `${col}22`
        ctx.fillRect(x, y, X(ref.range.c1 + 1) - x, Y(ref.range.r1 + 1) - y)
        ctx.lineWidth = 1
      })

      // 8. selection
      const x0 = X(range.c0)
      const y0 = Y(range.r0)
      const x1 = X(range.c1 + 1)
      const y1 = Y(range.r1 + 1)
      const act = tab.merges.find((m) => inRange(m, sel.anchor.r, sel.anchor.c)) ?? { r0: sel.anchor.r, c0: sel.anchor.c, r1: sel.anchor.r, c1: sel.anchor.c }
      ctx.fillStyle = C('--grid-sel-fill')
      ctx.beginPath()
      ctx.rect(x0, y0, x1 - x0, y1 - y0)
      ctx.rect(X(act.c0), Y(act.r0), X(act.c1 + 1) - X(act.c0), Y(act.r1 + 1) - Y(act.r0))
      ctx.fill('evenodd')
      ctx.strokeStyle = accent
      ctx.lineWidth = 2
      ctx.strokeRect(x0, y0, x1 - x0 - 1, y1 - y0 - 1)
      ctx.lineWidth = 1
      if (!edit) {
        ctx.fillStyle = accent
        ctx.fillRect(x1 - 5, y1 - 5, 7, 7)
        ctx.strokeStyle = bg
        ctx.strokeRect(x1 - 5.5, y1 - 5.5, 8, 8)
      }
      if (preview.fill) {
        const f = preview.fill
        ctx.setLineDash([3, 3])
        ctx.strokeStyle = C('--text-3')
        ctx.strokeRect(X(f.c0) + 0.5, Y(f.r0) + 0.5, X(f.c1 + 1) - X(f.c0) - 1, Y(f.r1 + 1) - Y(f.r0) - 1)
        ctx.setLineDash([])
      }
      if (copyRange) {
        ctx.setLineDash([5, 3])
        ctx.strokeStyle = accent
        ctx.lineWidth = 2
        ctx.strokeRect(X(copyRange.c0) + 1, Y(copyRange.r0) + 1, X(copyRange.c1 + 1) - X(copyRange.c0) - 2, Y(copyRange.r1 + 1) - Y(copyRange.r0) - 2)
        ctx.setLineDash([])
        ctx.lineWidth = 1
      }
      ctx.restore()
    }

    // headers
    const headerBg = C('--grid-header')
    const headerText = C('--grid-header-text')
    const headerSel = C('--grid-header-sel')
    const headerSelText = C('--grid-header-sel-text')
    ctx.font = '11.5px -apple-system, sans-serif'
    ctx.textBaseline = 'middle'
    ctx.textAlign = 'center'
    const colParts: [number, number, number, [number, number]][] = []
    if (fc) colParts.push([0, fc - 1, 0, [HW, fw]])
    colParts.push([cStart, cEnd, scroll.x, [HW + fw, w - HW - fw]])
    ctx.fillStyle = headerBg
    ctx.fillRect(0, 0, w, HH)
    ctx.fillRect(0, 0, HW, h)
    const wholeRows = range.c0 === 0 && range.c1 >= MAX_COLS - 1
    const wholeCols = range.r0 === 0 && range.r1 >= MAX_ROWS - 1
    for (const [c0, c1, sx, [cx, cw]] of colParts) {
      ctx.save()
      ctx.beginPath()
      ctx.rect(cx, 0, cw, HH)
      ctx.clip()
      for (let c = c0; c <= c1; c++) {
        const x = HW + cols.pos(c) - sx
        const wdt = cols.size(c)
        const on = c >= range.c0 && c <= range.c1
        if (on) {
          ctx.fillStyle = wholeCols ? accent : headerSel
          ctx.fillRect(x, 0, wdt, HH)
        }
        ctx.fillStyle = on ? (wholeCols ? '#fff' : headerSelText) : headerText
        if (wdt > 12) ctx.fillText(colName(c), x + wdt / 2, HH / 2 + 1)
        ctx.fillStyle = line
        ctx.fillRect(Math.round(x + wdt) - 1, 4, 1, HH - 8)
      }
      ctx.restore()
    }
    const rowParts: [number, number, number, [number, number]][] = []
    if (fr) rowParts.push([0, fr - 1, 0, [HH, fh]])
    rowParts.push([rStart, rEnd, scroll.y, [HH + fh, h - HH - fh]])
    for (const [r0, r1, sy, [ry, rhh]] of rowParts) {
      ctx.save()
      ctx.beginPath()
      ctx.rect(0, ry, HW, rhh)
      ctx.clip()
      for (let r = r0; r <= r1; r++) {
        const hgt = rows.size(r)
        if (!hgt) continue
        const y = HH + rows.pos(r) - sy
        const on = r >= range.r0 && r <= range.r1
        if (on) {
          ctx.fillStyle = wholeRows ? accent : headerSel
          ctx.fillRect(0, y, HW, hgt)
        }
        ctx.fillStyle = on ? (wholeRows ? '#fff' : headerSelText) : headerText
        if (hgt > 10) ctx.fillText(String(r + 1), HW / 2, y + hgt / 2 + 1)
        ctx.fillStyle = line
        ctx.fillRect(4, Math.round(y + hgt) - 1, HW - 8, 1)
      }
      ctx.restore()
    }
    ctx.fillStyle = C('--border')
    ctx.fillRect(0, HH - 1, w, 1)
    ctx.fillRect(HW - 1, 0, 1, h)
    // corner
    ctx.fillStyle = headerBg
    ctx.fillRect(0, 0, HW - 1, HH - 1)
    ctx.fillStyle = headerText
    ctx.beginPath()
    ctx.moveTo(HW - 6, HH - 6)
    ctx.lineTo(HW - 6, HH - 14)
    ctx.lineTo(HW - 14, HH - 6)
    ctx.fill()
    // freeze lines
    ctx.fillStyle = C('--grid-freeze')
    if (fc) ctx.fillRect(HW + fw - 1, 0, 2, h)
    if (fr) ctx.fillRect(0, HH + fh - 1, w, 2)
  })

  // ---------- mouse ----------
  const local = (e: { clientX: number; clientY: number }) => {
    const rect = scrollerRef.current!.getBoundingClientRect()
    return { x: e.clientX - rect.left, y: e.clientY - rect.top }
  }

  const colBoundary = (x: number): number | null => {
    const c = x < HW + fw ? cols.indexAt(x - HW) : cols.indexAt(x - HW + scroll.x)
    const left = cellRect(0, c).x
    const right = left + cols.size(c)
    if (right - x <= 4) return c
    if (x - left <= 3 && c > 0) return c - 1
    return null
  }
  const rowBoundary = (y: number): number | null => {
    const r = y < HH + fh ? rows.indexAt(y - HH) : rows.indexAt(y - HH + scroll.y)
    const top = cellRect(r, 0).y
    const bottom = top + rows.size(r)
    if (bottom - y <= 3) return r
    if (y - top <= 2 && r > 0) return r - 1
    return null
  }

  const canPick = (e: EditState | null): boolean => {
    if (!e || !e.text.startsWith('=')) return false
    if (e.pick && e.text.length === e.pick.end) return true
    const last = e.text.trimEnd().slice(-1)
    return '=(,+-*/^&<>:;'.includes(last)
  }

  const insertPick = (a: CellAddr, b: CellAddr) => {
    const e = propsRef.current.edit
    if (!e) return
    const rg = normRange(a, b)
    const ref = (e.tabId !== tab.id ? `${quoteSheet(tab.name)}!` : '') + rangeName(rg)
    const replacing = e.pick && e.text.length === e.pick.end
    const start = replacing ? e.pick!.start : e.text.length
    const textOut = e.text.slice(0, start) + ref
    propsRef.current.setEdit({ ...e, text: textOut, mode: 'edit', pick: { start, end: textOut.length, anchor: a, tabId: tab.id } })
  }

  const autoFit = (c: number) => {
    const ctx = canvasRef.current!.getContext('2d')!
    let max = 30
    const limit = Math.min(used.rows, 5000)
    for (let r = 0; r < limit; r++) {
      const s = wb.display(tab, r, c)
      if (!s) continue
      ctx.font = fontFor(tab.styles[styleKey(r, c)])
      max = Math.max(max, ctx.measureText(s).width + 14)
    }
    const targets = range.r0 === 0 && range.r1 >= MAX_ROWS - 1 && c >= range.c0 && c <= range.c1 ? Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i) : [c]
    wb.setColWidth(tab, targets, Math.min(max, 600))
  }

  const onMouseDown = (e: ReactMouseEvent) => {
    const { x, y } = local(e)
    const sc = scrollerRef.current!
    if (x > sc.clientWidth || y > sc.clientHeight) return // scrollbar
    const p = propsRef.current
    const shift = e.shiftKey

    if (y < HH && x < HW) {
      if (p.edit) p.commitEdit()
      p.setSel({ anchor: { r: 0, c: 0 }, focus: { r: MAX_ROWS - 1, c: MAX_COLS - 1 } })
      return
    }
    if (y < HH) {
      const b = colBoundary(x)
      if (b !== null && e.button === 0) {
        e.preventDefault()
        if (e.detail === 2) {
          autoFit(b)
          return
        }
        dragRef.current = { kind: 'resize-col', index: b, start: x, startSize: cols.size(b) }
      } else {
        const { c } = cellAt(x, HH + 1)
        if (e.button === 2 && c >= range.c0 && c <= range.c1 && range.r0 === 0 && range.r1 >= MAX_ROWS - 1) return
        if (p.edit) p.commitEdit()
        p.setSel({ anchor: { r: 0, c: shift ? sel.anchor.c : c }, focus: { r: MAX_ROWS - 1, c } })
        dragRef.current = { kind: 'cols' }
      }
    } else if (x < HW) {
      const b = rowBoundary(y)
      if (b !== null && e.button === 0) {
        e.preventDefault()
        dragRef.current = { kind: 'resize-row', index: b, start: y, startSize: rows.size(b) }
      } else {
        const { r } = cellAt(HW + 1, y)
        if (e.button === 2 && r >= range.r0 && r <= range.r1 && range.c0 === 0 && range.c1 >= MAX_COLS - 1) return
        if (p.edit) p.commitEdit()
        p.setSel({ anchor: { r: shift ? sel.anchor.r : r, c: 0 }, focus: { r, c: MAX_COLS - 1 } })
        dragRef.current = { kind: 'rows' }
      }
    } else {
      const cell = cellAt(x, y)
      // fill handle
      const br = cellRect(range.r1, range.c1)
      if (!p.edit && e.button === 0 && Math.abs(x - (br.x + br.w)) <= 5 && Math.abs(y - (br.y + br.h)) <= 5) {
        e.preventDefault()
        dragRef.current = { kind: 'fill', src: range, target: null }
        return
      }
      // filter dropdown button
      if (tab.filter && cell.r === tab.filter.range.r0 && cell.c >= tab.filter.range.c0 && cell.c <= tab.filter.range.c1) {
        const cr = cellRect(cell.r, cell.c)
        if (x >= cr.x + cr.w - 19) {
          e.preventDefault()
          p.onFilterButton(cell.c, e.clientX, e.clientY)
          return
        }
      }
      if (e.button === 2) {
        if (!inRange(range, cell.r, cell.c)) p.setSel({ anchor: cell, focus: cell })
        return
      }
      if (canPick(p.edit)) {
        e.preventDefault()
        insertPick(cell, cell)
        dragRef.current = { kind: 'pick' }
        return
      }
      if (p.edit) p.commitEdit()
      if (e.detail === 2 && !shift) {
        p.setSel({ anchor: cell, focus: cell })
        p.startEdit()
        return
      }
      const m = wb.mergeAt(tab, cell.r, cell.c)
      if (shift) p.setSel({ anchor: sel.anchor, focus: cell })
      else p.setSel(m ? { anchor: { r: m.r0, c: m.c0 }, focus: { r: m.r1, c: m.c1 } } : { anchor: cell, focus: cell })
      dragRef.current = { kind: 'cells' }
    }

    const move = (ev: MouseEvent) => {
      const d = dragRef.current
      if (!d) return
      const pt = local(ev)
      // autoscroll when dragging past the edges
      if (d.kind !== 'resize-col' && d.kind !== 'resize-row') {
        if (pt.y > sc.clientHeight) sc.scrollTop += pt.y - sc.clientHeight
        if (pt.y < HH && d.kind !== 'cols') sc.scrollTop -= HH - pt.y
        if (pt.x > sc.clientWidth) sc.scrollLeft += pt.x - sc.clientWidth
        if (pt.x < HW && d.kind !== 'rows') sc.scrollLeft -= HW - pt.x
      }
      const P = propsRef.current
      const cur = cellAt(Math.max(HW + 1, pt.x), Math.max(HH + 1, pt.y))
      switch (d.kind) {
        case 'cells': {
          const a = P.sel.anchor
          P.setSel({ anchor: a, focus: cur })
          break
        }
        case 'cols':
          P.setSel({ anchor: P.sel.anchor, focus: { r: MAX_ROWS - 1, c: cur.c } })
          break
        case 'rows':
          P.setSel({ anchor: P.sel.anchor, focus: { r: cur.r, c: MAX_COLS - 1 } })
          break
        case 'pick': {
          const pk = P.edit?.pick
          if (pk) insertPick(pk.anchor, cur)
          break
        }
        case 'fill': {
          const s = d.src
          const down = cur.r > s.r1 ? cur.r - s.r1 : cur.r < s.r0 ? s.r0 - cur.r : 0
          const right = cur.c > s.c1 ? cur.c - s.c1 : cur.c < s.c0 ? s.c0 - cur.c : 0
          let target: Range | null = null
          if (down >= right && down > 0) target = cur.r > s.r1 ? { ...s, r1: cur.r } : { ...s, r0: cur.r }
          else if (right > 0) target = cur.c > s.c1 ? { ...s, c1: cur.c } : { ...s, c0: cur.c }
          d.target = target
          setPreview({ fill: target })
          break
        }
        case 'resize-col':
          setPreview({ col: [d.index, Math.max(4, d.startSize + pt.x - d.start)] })
          break
        case 'resize-row':
          setPreview({ row: [d.index, Math.max(4, d.startSize + pt.y - d.start)] })
          break
      }
    }
    const up = (ev: MouseEvent) => {
      const d = dragRef.current
      dragRef.current = null
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      const P = propsRef.current
      const pt = local(ev)
      if (d?.kind === 'fill' && d.target) {
        P.onFill(d.src, d.target)
        P.setSel({ anchor: { r: d.target.r0, c: d.target.c0 }, focus: { r: d.target.r1, c: d.target.c1 } })
      } else if (d?.kind === 'resize-col') {
        const size = Math.max(4, d.startSize + pt.x - d.start)
        const whole = range.r0 === 0 && range.r1 >= MAX_ROWS - 1 && d.index >= range.c0 && d.index <= range.c1
        wb.setColWidth(tab, whole ? Array.from({ length: range.c1 - range.c0 + 1 }, (_, i) => range.c0 + i) : [d.index], size)
      } else if (d?.kind === 'resize-row') {
        const size = Math.max(4, d.startSize + pt.y - d.start)
        const whole = range.c0 === 0 && range.c1 >= MAX_COLS - 1 && d.index >= range.r0 && d.index <= range.r1
        wb.setRowHeight(tab, whole ? Array.from({ length: range.r1 - range.r0 + 1 }, (_, i) => range.r0 + i) : [d.index], size)
      }
      setPreview({})
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }

  const [cursor, setCursor] = useState('cell')
  const onMouseMove = (e: ReactMouseEvent) => {
    if (dragRef.current) return
    const { x, y } = local(e)
    let cur = 'cell'
    if (y < HH && x >= HW) cur = colBoundary(x) !== null ? 'col-resize' : 's-resize'
    else if (x < HW && y >= HH) cur = rowBoundary(y) !== null ? 'row-resize' : 'e-resize'
    else if (y >= HH && x >= HW) {
      const br = cellRect(range.r1, range.c1)
      if (Math.abs(x - (br.x + br.w)) <= 5 && Math.abs(y - (br.y + br.h)) <= 5) cur = 'crosshair'
    } else cur = 'default'
    if (cur !== cursor) setCursor(cur)
  }

  const onContextMenu = (e: ReactMouseEvent) => {
    e.preventDefault()
    const { x, y } = local(e)
    props.onContextMenu(y < HH && x < HW ? 'corner' : y < HH ? 'col' : x < HW ? 'row' : 'cell')
  }

  // ---------- keyboard ----------
  const moveSel = (dr: number, dc: number, extend: boolean, jump: boolean) => {
    const p = propsRef.current
    let { r, c } = extend ? p.sel.focus : p.sel.anchor
    const m = !extend ? wb.mergeAt(tab, r, c) : undefined
    if (m) {
      if (dr > 0) r = m.r1
      if (dc > 0) c = m.c1
    }
    if (jump) {
      // Cmd+Arrow: jump to the edge of the data region
      const filled = (rr: number, cc: number) => wb.raw(tab, rr, cc) !== null
      const inBounds = (rr: number, cc: number) => rr >= 0 && cc >= 0 && rr < rowCount && cc < colCount
      const startFilled = filled(r, c)
      const nextFilled = inBounds(r + dr, c + dc) && filled(r + dr, c + dc)
      if (startFilled && nextFilled) {
        while (inBounds(r + dr, c + dc) && filled(r + dr, c + dc)) {
          r += dr
          c += dc
        }
      } else {
        r += dr
        c += dc
        while (inBounds(r, c) && !filled(r, c) && inBounds(r + dr, c + dc)) {
          r += dr
          c += dc
        }
      }
      r = Math.max(0, Math.min(r, rowCount - 1))
      c = Math.max(0, Math.min(c, colCount - 1))
    } else {
      if (dr) r = rows.step(r, dr > 0 ? 1 : -1)
      if (dc) c = Math.max(0, Math.min(c + dc, MAX_COLS - 1))
    }
    const target = { r, c }
    const tm = wb.mergeAt(tab, r, c)
    if (extend) p.setSel({ anchor: p.sel.anchor, focus: target })
    else p.setSel(tm ? { anchor: { r: tm.r0, c: tm.c0 }, focus: { r: tm.r1, c: tm.c1 } } : { anchor: target, focus: target })
  }

  const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
    const p = propsRef.current
    if (p.edit) return
    const mod = e.metaKey || e.ctrlKey
    const pageRows = Math.max(1, Math.floor((view.h - HH) / DEFAULT_ROW_HEIGHT) - 1)
    switch (e.key) {
      case 'ArrowDown':
        e.preventDefault()
        return moveSel(1, 0, e.shiftKey, mod)
      case 'ArrowUp':
        e.preventDefault()
        return moveSel(-1, 0, e.shiftKey, mod)
      case 'ArrowRight':
        e.preventDefault()
        return moveSel(0, 1, e.shiftKey, mod)
      case 'ArrowLeft':
        e.preventDefault()
        return moveSel(0, -1, e.shiftKey, mod)
      case 'Enter':
        e.preventDefault()
        return moveSel(e.shiftKey ? -1 : 1, 0, false, false)
      case 'Tab':
        e.preventDefault()
        return moveSel(0, e.shiftKey ? -1 : 1, false, false)
      case 'PageDown':
        e.preventDefault()
        return moveSel(pageRows, 0, e.shiftKey, false)
      case 'PageUp':
        e.preventDefault()
        return moveSel(-pageRows, 0, e.shiftKey, false)
      case 'Home':
        e.preventDefault()
        return p.setSel(mod ? { anchor: { r: 0, c: 0 }, focus: { r: 0, c: 0 } } : { anchor: { r: p.sel.anchor.r, c: 0 }, focus: { r: p.sel.anchor.r, c: 0 } })
      case 'F2':
        e.preventDefault()
        return p.startEdit()
      case 'Backspace':
      case 'Delete':
        e.preventDefault()
        return p.onKeyCommand('clear')
      case 'Escape':
        return p.onKeyCommand('escape')
    }
    if (!mod && !e.altKey && e.key.length === 1) {
      e.preventDefault()
      p.startEdit(e.key)
    }
  }

  // ---------- in-cell editor ----------
  let editor: React.ReactNode = null
  if (edit && edit.tabId === tab.id) {
    const m = wb.mergeAt(tab, edit.r, edit.c)
    const rc = m ? cellRect(m.r0, m.c0, m.r1, m.c1) : cellRect(edit.r, edit.c)
    const st = tab.styles[styleKey(edit.r, edit.c)]
    const lines = edit.text.split('\n')
    const longest = Math.max(...lines.map((l) => l.length))
    const width = Math.min(Math.max(rc.w + 4, longest * ((st?.size ?? props.defaultFontSize) * 0.62) + 20), Math.max(rc.w + 4, view.w - rc.x - 8))
    const height = Math.max(rc.h + 4, lines.length * 18 + 8)
    editor = (
      <FormulaInput
        className="cell-editor"
        style={{
          left: rc.x - 2,
          top: rc.y - 2,
          width,
          height,
          font: fontFor(st),
          lineHeight: '18px',
          textAlign: st?.align ?? 'left',
          background: st?.fill ?? undefined,
          color: st?.color ?? undefined
        }}
        value={edit.text}
        functions={props.functions}
        autoFocus={edit.origin === 'cell'}
        caretToEnd
        onFocus={() => {
          const cur = propsRef.current.edit
          if (cur && cur.origin !== 'cell') propsRef.current.setEdit({ ...cur, origin: 'cell' })
        }}
        onChange={(text) => {
          const cur = propsRef.current.edit
          if (cur) propsRef.current.setEdit({ ...cur, text, pick: cur.pick && text.length >= cur.pick.end ? cur.pick : undefined })
        }}
        onKeyDown={(e) => {
          const cur = propsRef.current.edit
          if (!cur) return
          if (e.key === 'Enter' && !e.altKey) {
            e.preventDefault()
            props.commitEdit([e.shiftKey ? -1 : 1, 0])
          } else if (e.key === 'Enter' && e.altKey) {
            e.preventDefault()
            const ta = e.currentTarget
            const pos = ta.selectionStart
            const next = `${cur.text.slice(0, pos)}\n${cur.text.slice(ta.selectionEnd)}`
            propsRef.current.setEdit({ ...cur, text: next, mode: 'edit' })
            requestAnimationFrame(() => ta.setSelectionRange(pos + 1, pos + 1))
          } else if (e.key === 'Tab') {
            e.preventDefault()
            props.commitEdit([0, e.shiftKey ? -1 : 1])
          } else if (e.key === 'Escape') {
            e.preventDefault()
            props.cancelEdit()
          } else if (cur.mode === 'enter' && !cur.text.startsWith('=') && e.key.startsWith('Arrow')) {
            e.preventDefault()
            const d: Record<string, [number, number]> = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }
            props.commitEdit(d[e.key])
          }
        }}
      />
    )
  }

  const spacerW = HW + cols.total()
  const spacerH = Math.min(HH + rows.total(), 30_000_000)

  return (
    <div className="grid-host" ref={hostRef}>
      <canvas ref={canvasRef} />
      <div
        className="grid-scroller"
        ref={scrollerRef}
        tabIndex={0}
        style={{ cursor }}
        onScroll={onScroll}
        onMouseDown={onMouseDown}
        onMouseMove={onMouseMove}
        onContextMenu={onContextMenu}
        onKeyDown={onKeyDown}
      >
        <div style={{ width: spacerW, height: spacerH }} />
      </div>
      {editor}
    </div>
  )
}
