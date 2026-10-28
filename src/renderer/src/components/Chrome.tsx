import { useEffect, useState, type DragEvent } from 'react'
import { LayoutDashboard, Plus } from 'lucide-react'
import { DetailedCellError } from 'hyperformula'
import { MAX_COLS, MAX_ROWS, rangeName, type Range } from '../engine/addr'
import { clampRange, type GridSheet, type Tab, type Workbook } from '../engine/workbook'
import { FormulaInput } from './FormulaInput'
import type { EditState } from './types'

// ---------------------------------------------------------------- formula bar

export function FormulaBar(props: {
  nameBox: string
  value: string
  edit: EditState | null
  editingElsewhere: boolean
  functions: string[]
  disabled: boolean
  onGoto: (ref: string) => void
  onFocus: () => void
  onChange: (text: string) => void
  onCommit: (move?: [number, number]) => void
  onCancel: () => void
}) {
  const [name, setName] = useState(props.nameBox)
  useEffect(() => setName(props.nameBox), [props.nameBox])
  return (
    <div className="formula-bar">
      <input
        className="name-box"
        value={name}
        disabled={props.disabled}
        onChange={(e) => setName(e.target.value)}
        onFocus={(e) => e.target.select()}
        onKeyDown={(e) => {
          if (e.key === 'Enter') props.onGoto(name)
          if (e.key === 'Escape') setName(props.nameBox)
        }}
        spellCheck={false}
      />
      <div className="fx">fx</div>
      <div className="formula-input-wrap">
        <FormulaInput
          className="formula-input"
          value={props.edit ? props.edit.text : props.value}
          functions={props.functions}
          autoFocus={props.editingElsewhere || props.edit?.origin === 'bar'}
          caretToEnd
          onFocus={props.onFocus}
          onChange={props.onChange}
          onKeyDown={(e) => {
            if (e.key === 'Enter' && !e.altKey) {
              e.preventDefault()
              props.onCommit([0, 0])
            } else if (e.key === 'Escape') {
              e.preventDefault()
              props.onCancel()
            } else if (e.key === 'Tab') {
              e.preventDefault()
              props.onCommit([0, 1])
            }
          }}
        />
      </div>
    </div>
  )
}

// ---------------------------------------------------------------- sheet tabs

export function SheetTabs(props: {
  tabs: Tab[]
  activeId: string
  onSelect: (id: string) => void
  onAdd: () => void
  onRename: (id: string, name: string) => void
  onMove: (id: string, index: number) => void
  onContext: (id: string) => void
}) {
  const [renaming, setRenaming] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const [dropAt, setDropAt] = useState<number | null>(null)

  return (
    <div className="tabbar">
      <button className="tbtn" title="New sheet" onClick={props.onAdd}>
        <Plus size={15} />
      </button>
      {props.tabs.map((t, i) => (
        <button
          key={t.id}
          className={`sheet-tab${t.id === props.activeId ? ' active' : ''}${dropAt === i ? ' drop-before' : ''}`}
          draggable={renaming !== t.id}
          onDragStart={(e: DragEvent) => e.dataTransfer.setData('text/x-tab', t.id)}
          onDragOver={(e) => {
            e.preventDefault()
            setDropAt(i)
          }}
          onDragLeave={() => setDropAt(null)}
          onDrop={(e) => {
            const id = e.dataTransfer.getData('text/x-tab')
            setDropAt(null)
            if (id) props.onMove(id, i)
          }}
          onClick={() => props.onSelect(t.id)}
          onDoubleClick={() => {
            setRenaming(t.id)
            setDraft(t.name)
          }}
          onContextMenu={(e) => {
            e.preventDefault()
            props.onSelect(t.id)
            props.onContext(t.id)
          }}
        >
          {t.kind === 'dashboard' && <LayoutDashboard size={13} />}
          {renaming === t.id ? (
            <input
              autoFocus
              value={draft}
              onChange={(e) => setDraft(e.target.value)}
              onFocus={(e) => e.target.select()}
              onBlur={() => {
                props.onRename(t.id, draft)
                setRenaming(null)
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') (e.target as HTMLInputElement).blur()
                if (e.key === 'Escape') setRenaming(null)
              }}
            />
          ) : (
            t.name
          )}
          {t.kind === 'grid' && t.tabColor && <span className="tab-color" style={{ background: t.tabColor }} />}
        </button>
      ))}
      <div
        style={{ flex: 1, alignSelf: 'stretch' }}
        onDragOver={(e) => {
          e.preventDefault()
          setDropAt(props.tabs.length)
        }}
        onDrop={(e) => {
          const id = e.dataTransfer.getData('text/x-tab')
          setDropAt(null)
          if (id) props.onMove(id, props.tabs.length)
        }}
      />
    </div>
  )
}

// ---------------------------------------------------------------- status bar

export function StatusBar({ wb, tab, range, message }: { wb: Workbook; tab: GridSheet | null; range: Range | null; message?: string }) {
  let stats: { count: number; numCount: number; sum: number; min: number; max: number } | null = null
  if (tab && range && (range.r0 !== range.r1 || range.c0 !== range.c1)) {
    const rg = clampRange(range, wb.size(tab))
    if (rg && (rg.r1 - rg.r0 + 1) * (rg.c1 - rg.c0 + 1) <= 500_000) {
      stats = { count: 0, numCount: 0, sum: 0, min: Infinity, max: -Infinity }
      for (const row of wb.rangeValues(tab, rg)) {
        for (const v of row) {
          if (v === null || v === '') continue
          stats.count++
          if (typeof v === 'number') {
            stats.numCount++
            stats.sum += v
            stats.min = Math.min(stats.min, v)
            stats.max = Math.max(stats.max, v)
          } else if (v instanceof DetailedCellError) continue
        }
      }
    }
  }
  const n = (v: number) => v.toLocaleString(undefined, { maximumFractionDigits: 4 })
  const whole = range && range.r1 >= MAX_ROWS - 1 && range.c1 >= MAX_COLS - 1
  return (
    <div className="statusbar">
      <span>{message ?? 'Ready'}</span>
      <span className="spacer" />
      {range && !whole && <span>{rangeName(range)}</span>}
      {stats && stats.numCount > 0 && (
        <>
          <span>
            Average <b>{n(stats.sum / stats.numCount)}</b>
          </span>
          <span>
            Min <b>{n(stats.min)}</b>
          </span>
          <span>
            Max <b>{n(stats.max)}</b>
          </span>
          <span>
            Sum <b>{n(stats.sum)}</b>
          </span>
        </>
      )}
      {stats && (
        <span>
          Count <b>{stats.count}</b>
        </span>
      )}
    </div>
  )
}
