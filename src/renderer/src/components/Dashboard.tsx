import { useMemo, useState } from 'react'
import GridLayout, { useContainerWidth, type Layout } from 'react-grid-layout'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import { ChartColumn, Check, Gauge, GripVertical, Pencil, Table, Trash, Type } from 'lucide-react'
import { DetailedCellError, type CellValue } from 'hyperformula'
import { buildChartOption } from '../charts/options'
import { formatValue } from '../engine/format'
import type { Dashboard, Widget, Workbook } from '../engine/workbook'
import { EChart } from './EChart'

interface Props {
  wb: Workbook
  dash: Dashboard
  version: number
  dark: boolean
  onAdd: (type: Widget['type']) => void
  onEdit: (w: Widget) => void
}

export function DashboardView({ wb, dash, version, dark, onAdd, onEdit }: Props) {
  const [editing, setEditing] = useState(dash.widgets.length === 0)
  const { width, containerRef, mounted } = useContainerWidth()

  const layout: Layout = useMemo(() => dash.widgets.map((w) => ({ i: w.id, x: w.x, y: w.y, w: w.w, h: w.h, minW: 2, minH: 2 })), [dash.widgets])

  const commitLayout = (next: Layout) => {
    const changed = dash.widgets.some((w) => {
      const l = next.find((x) => x.i === w.id)
      return l && (l.x !== w.x || l.y !== w.y || l.w !== w.w || l.h !== w.h)
    })
    if (!changed) return
    wb.updateDashboard(
      dash.id,
      (d) => {
        d.widgets = d.widgets.map((w) => {
          const l = next.find((x) => x.i === w.id)
          return l ? { ...w, x: l.x, y: l.y, w: l.w, h: l.h } : w
        })
      },
      'Arrange Dashboard'
    )
  }

  return (
    <div className="dashboard">
      <div className="dash-head">
        <h2>{dash.name}</h2>
        <button className="tbtn" onClick={() => onAdd('chart')}>
          <ChartColumn size={15} />
          <span className="lbl">Chart</span>
        </button>
        <button className="tbtn" onClick={() => onAdd('kpi')}>
          <Gauge size={15} />
          <span className="lbl">KPI</span>
        </button>
        <button className="tbtn" onClick={() => onAdd('table')}>
          <Table size={15} />
          <span className="lbl">Table</span>
        </button>
        <button className="tbtn" onClick={() => onAdd('text')}>
          <Type size={15} />
          <span className="lbl">Text</span>
        </button>
        <span style={{ flex: 1 }} />
        <button className={`tbtn${editing ? ' primary' : ''}`} onClick={() => setEditing(!editing)}>
          {editing ? <Check size={15} /> : <Pencil size={15} />}
          <span className="lbl">{editing ? 'Done' : 'Edit layout'}</span>
        </button>
      </div>
      <div ref={containerRef}>
        {dash.widgets.length === 0 && (
          <div className="dash-empty">
            <p style={{ fontSize: 15, color: 'var(--text-2)', margin: '0 0 6px' }}>This dashboard is empty.</p>
            <p style={{ margin: 0 }}>Add charts, KPI cards, tables and notes that stay in sync with your sheets — or ask the AI assistant to build one for you.</p>
          </div>
        )}
        {mounted && dash.widgets.length > 0 && (
          <GridLayout
            width={width}
            layout={layout}
            gridConfig={{ cols: 12, rowHeight: 56, margin: [14, 14], containerPadding: [0, 0] }}
            dragConfig={{ enabled: editing, handle: '.widget-drag' }}
            resizeConfig={{ enabled: editing }}
            onDragStop={(l) => commitLayout(l)}
            onResizeStop={(l) => commitLayout(l)}
          >
            {dash.widgets.map((w) => (
              <div key={w.id}>
                <WidgetCard
                  wb={wb}
                  w={w}
                  version={version}
                  dark={dark}
                  editing={editing}
                  onEdit={() => onEdit(w)}
                  onDelete={() => wb.updateDashboard(dash.id, (d) => (d.widgets = d.widgets.filter((x) => x.id !== w.id)), 'Delete Widget')}
                  onText={(text) =>
                    wb.updateDashboard(dash.id, (d) => (d.widgets = d.widgets.map((x) => (x.id === w.id ? { ...x, text } : x))), 'Edit Text')
                  }
                />
              </div>
            ))}
          </GridLayout>
        )}
      </div>
    </div>
  )
}

function WidgetCard({
  wb,
  w,
  version,
  dark,
  editing,
  onEdit,
  onDelete,
  onText
}: {
  wb: Workbook
  w: Widget
  version: number
  dark: boolean
  editing: boolean
  onEdit: () => void
  onDelete: () => void
  onText: (t: string) => void
}) {
  return (
    <div className="widget">
      <div className={`widget-head${editing ? ' widget-drag' : ''}`}>
        {editing && <GripVertical size={14} style={{ color: 'var(--text-3)', marginLeft: -6 }} />}
        <span className="wtitle">{w.type === 'chart' ? '' : (w.title ?? '')}</span>
        {editing && (
          <>
            {w.type !== 'text' && (
              <button className="tbtn" title="Edit" onMouseDown={(e) => e.stopPropagation()} onClick={onEdit}>
                <Pencil size={13} />
              </button>
            )}
            <button className="tbtn" title="Delete" onMouseDown={(e) => e.stopPropagation()} onClick={onDelete}>
              <Trash size={13} />
            </button>
          </>
        )}
      </div>
      <div className="widget-body">
        {w.type === 'chart' && w.chart && <ChartBody wb={wb} w={w} version={version} dark={dark} />}
        {w.type === 'kpi' && w.kpi && <KpiBody wb={wb} kpi={w.kpi} />}
        {w.type === 'table' && w.table && <TableBody wb={wb} range={w.table.range} />}
        {w.type === 'text' &&
          (editing ? (
            <div className="widget-text">
              <textarea defaultValue={w.text ?? ''} onBlur={(e) => e.target.value !== (w.text ?? '') && onText(e.target.value)} placeholder="Write notes in Markdown…" />
            </div>
          ) : (
            <div className="widget-text">
              <ReactMarkdown remarkPlugins={[remarkGfm]}>{w.text || '*Empty note*'}</ReactMarkdown>
            </div>
          ))}
      </div>
    </div>
  )
}

function ChartBody({ wb, w, version, dark }: { wb: Workbook; w: Widget; version: number; dark: boolean }) {
  const option = useMemo(() => {
    const res = wb.refValues(w.chart!.range)
    return buildChartOption(w.chart!, res?.values ?? [], dark, w.title)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [w, version, dark])
  if (!wb.refValues(w.chart!.range)) return <div className="muted">Range “{w.chart!.range}” not found.</div>
  return <EChart option={option} />
}

export function formatKpi(v: CellValue, fmt?: string): string {
  if (v instanceof DetailedCellError) return v.value
  if (typeof v === 'number' && !fmt) return v.toLocaleString(undefined, { maximumFractionDigits: 2 })
  return formatValue(v, fmt)
}

function KpiBody({ wb, kpi }: { wb: Workbook; kpi: NonNullable<Widget['kpi']> }) {
  const v = wb.evaluate(kpi.formula)
  const cmp = kpi.compare ? wb.evaluate(kpi.compare) : null
  let delta: number | null = null
  if (typeof v === 'number' && typeof cmp === 'number' && cmp !== 0) delta = (v - cmp) / Math.abs(cmp)
  return (
    <div className="kpi">
      <div className="kpi-value">{formatKpi(v, kpi.fmt)}</div>
      {delta !== null && (
        <div className="kpi-sub">
          <span className={delta >= 0 ? 'delta-up' : 'delta-down'}>
            {delta >= 0 ? '▲' : '▼'} {Math.abs(delta * 100).toFixed(1)}%
          </span>{' '}
          {kpi.compareLabel ?? 'vs. comparison'}
        </div>
      )}
    </div>
  )
}

function TableBody({ wb, range }: { wb: Workbook; range: string }) {
  const res = wb.resolveRef(range, wb.grids()[0])
  if (!res) return <div className="muted">Range “{range}” not found.</div>
  const { tab, range: rg } = res
  const size = wb.size(tab)
  const r1 = Math.min(rg.r1, size.rows - 1, rg.r0 + 500)
  const c1 = Math.min(rg.c1, size.cols - 1, rg.c0 + 30)
  const rows: { text: string; num: boolean }[][] = []
  for (let r = rg.r0; r <= r1; r++) {
    const row = []
    for (let c = rg.c0; c <= c1; c++) row.push({ text: wb.display(tab, r, c), num: typeof wb.value(tab, r, c) === 'number' })
    rows.push(row)
  }
  const [head, ...body] = rows
  return (
    <div className="widget-table">
      <table>
        {head && (
          <thead>
            <tr>
              {head.map((c, i) => (
                <th key={i}>{c.text}</th>
              ))}
            </tr>
          </thead>
        )}
        <tbody>
          {body.map((row, i) => (
            <tr key={i}>
              {row.map((c, j) => (
                <td key={j} className={c.num ? 'num' : ''}>
                  {c.text}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}
