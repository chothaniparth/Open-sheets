import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import { ChartArea, ChartBar, ChartColumn, ChartColumnStacked, ChartLine, ChartPie, ChartScatter, Donut, Funnel, Radar, X } from 'lucide-react'
import { buildChartOption } from '../charts/options'
import { cellName, rangeName, type CellAddr, type Range } from '../engine/addr'
import { parseInput } from '../engine/format'
import { clampRange, newId, type ChartSpec, type ChartType, type CondRule, type CondType, type GridSheet, type Widget, type Workbook } from '../engine/workbook'
import { formatKpi } from './Dashboard'
import { EChart } from './EChart'

export function Modal({ title, onClose, children, footer, wide }: { title: string; onClose: () => void; children: ReactNode; footer?: ReactNode; wide?: boolean }) {
  useEffect(() => {
    const k = (e: KeyboardEvent) => e.key === 'Escape' && onClose()
    window.addEventListener('keydown', k)
    return () => window.removeEventListener('keydown', k)
  }, [onClose])
  return (
    <div className="overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div className={`dialog${wide ? ' wide' : ''}`} role="dialog" aria-label={title}>
        <div className="dialog-head">
          {title}
          <button className="tbtn x" onClick={onClose} title="Close">
            <X size={16} />
          </button>
        </div>
        <div className="dialog-body">{children}</div>
        {footer && <div className="dialog-foot">{footer}</div>}
      </div>
    </div>
  )
}

// ---------------------------------------------------------------- shared destination picker

function DashboardPicker({ wb, value, onChange }: { wb: Workbook; value: string; onChange: (v: string) => void }) {
  const dashboards = wb.tabs.filter((t) => t.kind === 'dashboard')
  const [isNew, setIsNew] = useState(dashboards.length === 0)
  return (
    <div className="row">
      <label className="field">
        <span>Place on dashboard</span>
        <select
          value={isNew ? '__new' : value}
          onChange={(e) => {
            if (e.target.value === '__new') {
              setIsNew(true)
              onChange(wb.uniqueName('Dashboard'))
            } else {
              setIsNew(false)
              onChange(e.target.value)
            }
          }}
        >
          {dashboards.map((d) => (
            <option key={d.id} value={d.name}>
              {d.name}
            </option>
          ))}
          <option value="__new">New dashboard…</option>
        </select>
      </label>
      {isNew && (
        <label className="field">
          <span>New dashboard name</span>
          <input value={value} onChange={(e) => onChange(e.target.value)} />
        </label>
      )}
    </div>
  )
}

function defaultDashboard(wb: Workbook, preferred?: string) {
  return preferred ?? wb.tabs.find((t) => t.kind === 'dashboard')?.name ?? 'Dashboard'
}

// ---------------------------------------------------------------- chart

const CHART_TYPES: { type: ChartType; label: string; icon: ReactNode }[] = [
  { type: 'column', label: 'Column', icon: <ChartColumn size={20} /> },
  { type: 'stackedColumn', label: 'Stacked', icon: <ChartColumnStacked size={20} /> },
  { type: 'bar', label: 'Bar', icon: <ChartBar size={20} /> },
  { type: 'line', label: 'Line', icon: <ChartLine size={20} /> },
  { type: 'area', label: 'Area', icon: <ChartArea size={20} /> },
  { type: 'pie', label: 'Pie', icon: <ChartPie size={20} /> },
  { type: 'doughnut', label: 'Doughnut', icon: <Donut size={20} /> },
  { type: 'scatter', label: 'Scatter', icon: <ChartScatter size={20} /> },
  { type: 'radar', label: 'Radar', icon: <Radar size={20} /> },
  { type: 'funnel', label: 'Funnel', icon: <Funnel size={20} /> }
]

export interface WidgetSave {
  dashboard: string
  widget: Omit<Widget, 'id' | 'x' | 'y'>
}

export function ChartDialog({ wb, dark, initialRange, widget, preferredDashboard, onSave, onClose }: { wb: Workbook; dark: boolean; initialRange: string; widget?: Widget; preferredDashboard?: string; onSave: (s: WidgetSave) => void; onClose: () => void }) {
  const [spec, setSpec] = useState<ChartSpec>(widget?.chart ?? { type: 'column', range: initialRange, seriesBy: 'columns', headers: true, legend: true })
  const [title, setTitle] = useState(widget?.title ?? '')
  const [dash, setDash] = useState(defaultDashboard(wb, preferredDashboard))
  const res = wb.refValues(spec.range)
  const option = useMemo(() => buildChartOption(spec, res?.values ?? [], dark, title || undefined), [spec, title, dark, res?.values])
  const save = () => onSave({ dashboard: dash, widget: { type: 'chart', w: widget?.w ?? 6, h: widget?.h ?? 5, title: title || undefined, chart: spec } })
  return (
    <Modal
      title={widget ? 'Edit chart' : 'Insert chart'}
      wide
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            Cancel
          </button>
          <button className="btn primary" disabled={!res} onClick={save}>
            {widget ? 'Save' : 'Insert'}
          </button>
        </>
      }
    >
      <div className="chart-types">
        {CHART_TYPES.map((t) => (
          <button key={t.type} className={`chart-type${spec.type === t.type ? ' on' : ''}`} onClick={() => setSpec({ ...spec, type: t.type })}>
            {t.icon}
            {t.label}
          </button>
        ))}
      </div>
      <div className="row">
        <label className="field">
          <span>Data range</span>
          <input value={spec.range} onChange={(e) => setSpec({ ...spec, range: e.target.value })} spellCheck={false} />
        </label>
        <label className="field">
          <span>Title</span>
          <input value={title} onChange={(e) => setTitle(e.target.value)} placeholder="Optional" />
        </label>
      </div>
      <div className="row" style={{ gap: 18 }}>
        <label className="check">
          <input type="checkbox" checked={spec.headers} onChange={(e) => setSpec({ ...spec, headers: e.target.checked })} /> First row is series names
        </label>
        <label className="check">
          <input type="checkbox" checked={spec.seriesBy === 'rows'} onChange={(e) => setSpec({ ...spec, seriesBy: e.target.checked ? 'rows' : 'columns' })} /> Series in rows
        </label>
        <label className="check">
          <input type="checkbox" checked={spec.legend} onChange={(e) => setSpec({ ...spec, legend: e.target.checked })} /> Legend
        </label>
        {(spec.type === 'line' || spec.type === 'area') && (
          <label className="check">
            <input type="checkbox" checked={!!spec.smooth} onChange={(e) => setSpec({ ...spec, smooth: e.target.checked })} /> Smooth
          </label>
        )}
      </div>
      <div className="chart-preview">{res ? <EChart option={option} /> : <div className="muted">Enter a valid range such as Sheet1!A1:C12</div>}</div>
      {!widget && <DashboardPicker wb={wb} value={dash} onChange={setDash} />}
    </Modal>
  )
}

// ---------------------------------------------------------------- KPI

export function KpiDialog({ wb, widget, defaultFormula, preferredDashboard, onSave, onClose }: { wb: Workbook; widget?: Widget; defaultFormula: string; preferredDashboard?: string; onSave: (s: WidgetSave) => void; onClose: () => void }) {
  const [title, setTitle] = useState(widget?.title ?? 'Total')
  const [formula, setFormula] = useState(widget?.kpi?.formula ?? defaultFormula)
  const [fmt, setFmt] = useState(widget?.kpi?.fmt ?? '')
  const [compare, setCompare] = useState(widget?.kpi?.compare ?? '')
  const [compareLabel, setCompareLabel] = useState(widget?.kpi?.compareLabel ?? '')
  const [dash, setDash] = useState(defaultDashboard(wb, preferredDashboard))
  const value = wb.evaluate(formula)
  return (
    <Modal
      title={widget ? 'Edit KPI card' : 'Add KPI card'}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            Cancel
          </button>
          <button
            className="btn primary"
            onClick={() =>
              onSave({ dashboard: dash, widget: { type: 'kpi', w: widget?.w ?? 3, h: widget?.h ?? 2, title, kpi: { formula, fmt: fmt || undefined, compare: compare || undefined, compareLabel: compareLabel || undefined } } })
            }
          >
            {widget ? 'Save' : 'Add'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Label</span>
        <input value={title} onChange={(e) => setTitle(e.target.value)} />
      </label>
      <label className="field">
        <span>Value formula</span>
        <input value={formula} onChange={(e) => setFormula(e.target.value)} spellCheck={false} placeholder="=SUM(Sheet1!B2:B100)" />
      </label>
      <div className="row">
        <label className="field">
          <span>Number format</span>
          <select value={fmt} onChange={(e) => setFmt(e.target.value)}>
            <option value="">Automatic</option>
            <option value="#,##0">1,234</option>
            <option value="#,##0.00">1,234.56</option>
            <option value="$#,##0">$1,234</option>
            <option value="$#,##0.00">$1,234.56</option>
            <option value="0%">12%</option>
            <option value="0.0%">12.3%</option>
          </select>
        </label>
        <div className="field">
          <span>Preview</span>
          <div style={{ fontSize: 22, fontWeight: 650 }}>{formatKpi(value, fmt || undefined)}</div>
        </div>
      </div>
      <div className="row">
        <label className="field">
          <span>Compare with (optional)</span>
          <input value={compare} onChange={(e) => setCompare(e.target.value)} placeholder="=Sheet1!C20" spellCheck={false} />
        </label>
        <label className="field">
          <span>Comparison label</span>
          <input value={compareLabel} onChange={(e) => setCompareLabel(e.target.value)} placeholder="vs. last month" />
        </label>
      </div>
      {!widget && <DashboardPicker wb={wb} value={dash} onChange={setDash} />}
    </Modal>
  )
}

export function TableDialog({ wb, widget, initialRange, preferredDashboard, onSave, onClose }: { wb: Workbook; widget?: Widget; initialRange: string; preferredDashboard?: string; onSave: (s: WidgetSave) => void; onClose: () => void }) {
  const [title, setTitle] = useState(widget?.title ?? '')
  const [range, setRange] = useState(widget?.table?.range ?? initialRange)
  const [dash, setDash] = useState(defaultDashboard(wb, preferredDashboard))
  const ok = !!wb.resolveRef(range, wb.grids()[0])
  return (
    <Modal
      title={widget ? 'Edit table' : 'Add table'}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            Cancel
          </button>
          <button className="btn primary" disabled={!ok} onClick={() => onSave({ dashboard: dash, widget: { type: 'table', w: widget?.w ?? 6, h: widget?.h ?? 5, title: title || undefined, table: { range } } })}>
            {widget ? 'Save' : 'Add'}
          </button>
        </>
      }
    >
      <label className="field">
        <span>Title</span>
        <input value={title} onChange={(e) => setTitle(e.target.value)} />
      </label>
      <label className="field">
        <span>Range (first row = headers)</span>
        <input value={range} onChange={(e) => setRange(e.target.value)} spellCheck={false} />
      </label>
      {!widget && <DashboardPicker wb={wb} value={dash} onChange={setDash} />}
    </Modal>
  )
}

// ---------------------------------------------------------------- prompt

export function PromptDialog({ title, label, initial, onOk, onClose }: { title: string; label: string; initial: string; onOk: (v: string) => void; onClose: () => void }) {
  const [v, setV] = useState(initial)
  return (
    <Modal
      title={title}
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            Cancel
          </button>
          <button className="btn primary" onClick={() => onOk(v)}>
            OK
          </button>
        </>
      }
    >
      <label className="field">
        <span>{label}</span>
        <input autoFocus value={v} onChange={(e) => setV(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && onOk(v)} spellCheck={false} />
      </label>
    </Modal>
  )
}

// ---------------------------------------------------------------- functions

export function FunctionDialog({ functions, onInsert, onClose }: { functions: string[]; onInsert: (name: string) => void; onClose: () => void }) {
  const [q, setQ] = useState('')
  const [sel, setSel] = useState<string | null>(null)
  const list = functions.filter((f) => f.includes(q.toUpperCase()))
  return (
    <Modal
      title={`Insert function (${functions.length} available)`}
      wide
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={onClose}>
            Cancel
          </button>
          <button className="btn primary" disabled={!sel} onClick={() => sel && onInsert(sel)}>
            Insert
          </button>
        </>
      }
    >
      <label className="field">
        <span>Search</span>
        <input autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="e.g. LOOKUP, SUM, DATE" />
      </label>
      <div className="fn-list">
        {list.map((f) => (
          <div key={f} className={sel === f ? 'sel' : ''} onClick={() => setSel(f)} onDoubleClick={() => onInsert(f)}>
            {f}
          </div>
        ))}
      </div>
      <div className="hint">Formulas follow Excel syntax. Double-click a function to insert it in the active cell.</div>
    </Modal>
  )
}

// ---------------------------------------------------------------- conditional formatting

const COND_LABELS: Record<CondType, string> = {
  gt: 'Greater than',
  lt: 'Less than',
  between: 'Between',
  eq: 'Equal to',
  contains: 'Text contains',
  blank: 'Is blank',
  duplicate: 'Duplicate values',
  scale: 'Color scale (min → max)'
}

export function CondFormatDialog({ wb, tab, selection, onClose }: { wb: Workbook; tab: GridSheet; selection: Range; onClose: () => void }) {
  const [rules, setRules] = useState<CondRule[]>(tab.cond)
  const add = () => setRules([...rules, { id: newId('cf'), range: rangeName(selection), type: 'gt', a: '0', fill: '#c6efce', color: '#006100' }])
  const update = (i: number, patch: Partial<CondRule>) => setRules(rules.map((r, j) => (j === i ? { ...r, ...patch } : r)))
  return (
    <Modal
      title="Conditional formatting"
      wide
      onClose={onClose}
      footer={
        <>
          <button className="btn" onClick={add}>
            Add rule
          </button>
          <span style={{ flex: 1 }} />
          <button className="btn" onClick={onClose}>
            Cancel
          </button>
          <button
            className="btn primary"
            onClick={() => {
              wb.setCondRules(tab, rules)
              onClose()
            }}
          >
            Apply
          </button>
        </>
      }
    >
      {rules.length === 0 && <div className="muted">No rules yet. Rules highlight cells based on their values and update as data changes.</div>}
      {rules.map((r, i) => (
        <div key={r.id} className="row" style={{ alignItems: 'flex-end', flexWrap: 'wrap', borderBottom: '1px solid var(--border)', paddingBottom: 10 }}>
          <label className="field" style={{ maxWidth: 110 }}>
            <span>Range</span>
            <input value={r.range} onChange={(e) => update(i, { range: e.target.value })} />
          </label>
          <label className="field" style={{ maxWidth: 190 }}>
            <span>Rule</span>
            <select value={r.type} onChange={(e) => update(i, { type: e.target.value as CondType, scale: e.target.value === 'scale' ? ['#f8696b', '#63be7b'] : r.scale })}>
              {Object.entries(COND_LABELS).map(([k, v]) => (
                <option key={k} value={k}>
                  {v}
                </option>
              ))}
            </select>
          </label>
          {['gt', 'lt', 'between', 'eq', 'contains'].includes(r.type) && (
            <label className="field" style={{ maxWidth: 100 }}>
              <span>Value</span>
              <input value={r.a ?? ''} onChange={(e) => update(i, { a: e.target.value })} />
            </label>
          )}
          {r.type === 'between' && (
            <label className="field" style={{ maxWidth: 100 }}>
              <span>And</span>
              <input value={r.b ?? ''} onChange={(e) => update(i, { b: e.target.value })} />
            </label>
          )}
          {r.type === 'scale' ? (
            <>
              <label className="field" style={{ maxWidth: 70 }}>
                <span>Min</span>
                <input type="color" value={r.scale?.[0] ?? '#f8696b'} onChange={(e) => update(i, { scale: [e.target.value, r.scale?.[1] ?? '#63be7b'] })} />
              </label>
              <label className="field" style={{ maxWidth: 70 }}>
                <span>Max</span>
                <input type="color" value={r.scale?.[1] ?? '#63be7b'} onChange={(e) => update(i, { scale: [r.scale?.[0] ?? '#f8696b', e.target.value] })} />
              </label>
            </>
          ) : (
            <>
              <label className="field" style={{ maxWidth: 70 }}>
                <span>Fill</span>
                <input type="color" value={r.fill ?? '#ffffff'} onChange={(e) => update(i, { fill: e.target.value })} />
              </label>
              <label className="field" style={{ maxWidth: 70 }}>
                <span>Text</span>
                <input type="color" value={r.color ?? '#000000'} onChange={(e) => update(i, { color: e.target.value })} />
              </label>
              <label className="check" style={{ paddingBottom: 6 }}>
                <input type="checkbox" checked={!!r.bold} onChange={(e) => update(i, { bold: e.target.checked })} /> Bold
              </label>
            </>
          )}
          <button className="btn small danger" onClick={() => setRules(rules.filter((_, j) => j !== i))}>
            Remove
          </button>
        </div>
      ))}
    </Modal>
  )
}

// ---------------------------------------------------------------- find & replace

export function FindBar({ wb, tab, from, onSelect, onClose }: { wb: Workbook; tab: GridSheet; from: CellAddr; onSelect: (c: CellAddr) => void; onClose: () => void }) {
  const [q, setQ] = useState('')
  const [rep, setRep] = useState('')
  const [matchCase, setMatchCase] = useState(false)
  const [status, setStatus] = useState('')
  const cursor = useRef(from)
  const input = useRef<HTMLInputElement>(null)
  useEffect(() => input.current?.focus(), [])

  const matches = (r: number, c: number) => {
    const s = wb.display(tab, r, c)
    return matchCase ? s.includes(q) : s.toLowerCase().includes(q.toLowerCase())
  }
  const findNext = (dir: 1 | -1 = 1) => {
    if (!q) return
    const { rows, cols } = wb.size(tab)
    const total = rows * cols
    if (!total) return setStatus('No matches')
    let idx = cursor.current.r * cols + cursor.current.c
    for (let i = 0; i < total; i++) {
      idx = (idx + dir + total) % total
      const r = Math.floor(idx / cols)
      const c = idx % cols
      if (matches(r, c)) {
        cursor.current = { r, c }
        onSelect({ r, c })
        setStatus(cellName(r, c))
        return
      }
    }
    setStatus('No matches')
  }
  const replaceIn = (raw: unknown): string => {
    const s = raw === null || raw === undefined ? '' : String(raw)
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), matchCase ? 'g' : 'gi')
    return s.replace(re, rep)
  }
  const replaceAll = () => {
    if (!q) return
    const { rows, cols } = wb.size(tab)
    const changes: { r: number; c: number; raw: ReturnType<typeof parseInput>['raw'] }[] = []
    for (let r = 0; r < rows; r++)
      for (let c = 0; c < cols; c++) {
        const raw = wb.raw(tab, r, c)
        if (raw === null || !matches(r, c)) continue
        const next = replaceIn(raw)
        if (next !== String(raw)) changes.push({ r, c, raw: parseInput(next).raw })
      }
    if (!changes.length) return setStatus('No matches')
    wb.transact('Replace All', () => changes.forEach((ch) => wb.hf.setCellContents({ sheet: wb.sid(tab), row: ch.r, col: ch.c }, ch.raw)), { hf: true })
    setStatus(`Replaced ${changes.length} cell${changes.length === 1 ? '' : 's'}`)
  }
  const replaceOne = () => {
    const { r, c } = cursor.current
    if (q && matches(r, c)) wb.setInput(tab, r, c, replaceIn(wb.raw(tab, r, c)))
    findNext()
  }
  return (
    <div className="floating findbar" onKeyDown={(e) => e.key === 'Escape' && onClose()}>
      <div className="row">
        <input ref={input} placeholder="Find" value={q} onChange={(e) => setQ(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && findNext(e.shiftKey ? -1 : 1)} />
        <button className="tbtn" onClick={onClose} title="Close">
          <X size={14} />
        </button>
      </div>
      <div className="row">
        <input placeholder="Replace with" value={rep} onChange={(e) => setRep(e.target.value)} />
      </div>
      <div className="row" style={{ justifyContent: 'space-between' }}>
        <label className="check muted">
          <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} /> Match case
        </label>
        <span className="muted">{status}</span>
      </div>
      <div className="row">
        <button className="btn small" onClick={() => findNext(-1)}>
          Previous
        </button>
        <button className="btn small" onClick={() => findNext(1)}>
          Next
        </button>
        <span style={{ flex: 1 }} />
        <button className="btn small" onClick={replaceOne}>
          Replace
        </button>
        <button className="btn small" onClick={replaceAll}>
          All
        </button>
      </div>
    </div>
  )
}

// ---------------------------------------------------------------- filter popup

export function FilterPopup({ wb, tab, col, x, y, onClose }: { wb: Workbook; tab: GridSheet; col: number; x: number; y: number; onClose: () => void }) {
  const filter = tab.filter!
  const values = useMemo(() => wb.distinctValues(tab, filter.range, col), [wb, tab, filter, col])
  const [checked, setChecked] = useState<Set<string>>(new Set(filter.criteria[col] ?? values))
  const [q, setQ] = useState('')
  const ref = useRef<HTMLDivElement>(null)
  useEffect(() => {
    const away = (e: MouseEvent) => !ref.current?.contains(e.target as Node) && onClose()
    setTimeout(() => window.addEventListener('mousedown', away))
    return () => window.removeEventListener('mousedown', away)
  }, [onClose])
  const shown = values.filter((v) => v.toLowerCase().includes(q.toLowerCase()))
  const apply = () => {
    const criteria = { ...filter.criteria }
    if (checked.size === values.length) delete criteria[col]
    else criteria[col] = [...checked]
    wb.setFilter(tab, { ...filter, criteria })
    onClose()
  }
  const sort = (desc: boolean) => {
    const rg = clampRange(filter.range, wb.size(tab))
    if (rg) wb.sortRange(tab, rg, col, desc, true)
    onClose()
  }
  return (
    <div ref={ref} className="floating filter-pop" style={{ position: 'fixed', left: Math.min(x, window.innerWidth - 260), top: y + 8 }}>
      <div className="row">
        <button className="btn small" onClick={() => sort(false)}>
          Sort A→Z
        </button>
        <button className="btn small" onClick={() => sort(true)}>
          Sort Z→A
        </button>
      </div>
      <input className="tselect" placeholder="Search values" value={q} onChange={(e) => setQ(e.target.value)} />
      <div className="values">
        <label>
          <input type="checkbox" checked={checked.size === values.length} onChange={(e) => setChecked(new Set(e.target.checked ? values : []))} /> (Select all)
        </label>
        {shown.map((v) => (
          <label key={v} title={v}>
            <input
              type="checkbox"
              checked={checked.has(v)}
              onChange={(e) => {
                const next = new Set(checked)
                if (e.target.checked) next.add(v)
                else next.delete(v)
                setChecked(next)
              }}
            />
            {v === '' ? '(Blanks)' : v}
          </label>
        ))}
      </div>
      <div className="row" style={{ justifyContent: 'flex-end' }}>
        <button
          className="btn small"
          onClick={() => {
            const criteria = { ...filter.criteria }
            delete criteria[col]
            wb.setFilter(tab, { ...filter, criteria })
            onClose()
          }}
        >
          Clear
        </button>
        <button className="btn small primary" onClick={apply}>
          OK
        </button>
      </div>
    </div>
  )
}
