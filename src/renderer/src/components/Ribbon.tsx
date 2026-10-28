import type { ReactNode } from 'react'
import {
  ArrowDownAZ,
  ArrowDownZA,
  Baseline,
  BetweenHorizontalStart,
  BetweenVerticalStart,
  Bold,
  ChartColumn,
  ClipboardPaste,
  Columns3,
  Copy,
  Eraser,
  Eye,
  FilePlus,
  Funnel,
  Gauge,
  Grid3x3,
  Highlighter,
  Italic,
  LayoutDashboard,
  ListX,
  PaintBucket,
  Percent,
  Printer,
  Redo2,
  Rows3,
  Scissors,
  Search,
  Sigma,
  Snowflake,
  SquareFunction,
  Strikethrough,
  TableCellsMerge,
  TextAlignCenter,
  TextAlignEnd,
  TextAlignStart,
  TextWrap,
  Type,
  Underline,
  Undo2,
  DollarSign
} from 'lucide-react'
import { NUMBER_FORMATS } from '../engine/format'
import type { CellStyle } from '../engine/workbook'

export type RibbonTab = 'home' | 'insert' | 'data' | 'view'

interface Props {
  tab: RibbonTab
  style: CellStyle | undefined
  canUndo: boolean
  canRedo: boolean
  isGrid: boolean
  showFormulas: boolean
  showGridlines: boolean
  run: (cmd: string, arg?: string) => void
}

const FONTS = ['Helvetica Neue', 'Arial', 'Avenir Next', 'Calibri', 'Cambria', 'Courier New', 'Georgia', 'Gill Sans', 'Menlo', 'Times New Roman', 'Verdana']
const SIZES = [9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 36, 48]

function B({ icon, title, onClick, on, disabled, label }: { icon: ReactNode; title: string; onClick: () => void; on?: boolean; disabled?: boolean; label?: string }) {
  return (
    <button className={`tbtn${on ? ' on' : ''}`} title={title} onMouseDown={(e) => e.preventDefault()} onClick={onClick} disabled={disabled}>
      {icon}
      {label && <span className="lbl">{label}</span>}
    </button>
  )
}

function ColorButton({ icon, title, color, onPick }: { icon: ReactNode; title: string; color: string; onPick: (c: string) => void }) {
  return (
    <span className="tbtn color-btn" title={title}>
      {icon}
      <span className="swatch" style={{ background: color }} />
      <input type="color" value={color} onChange={(e) => onPick(e.target.value)} />
    </span>
  )
}

export function Ribbon({ tab, style: st, canUndo, canRedo, isGrid, showFormulas, showGridlines, run }: Props) {
  const s = 16
  if (tab === 'home') {
    const fmtValue = NUMBER_FORMATS.find((f) => f.fmt === st?.fmt)?.fmt ?? (st?.fmt ? 'custom' : 'General')
    return (
      <div className="ribbon">
        <div className="rgroup">
          <B icon={<Undo2 size={s} />} title="Undo (⌘Z)" onClick={() => run('undo')} disabled={!canUndo} />
          <B icon={<Redo2 size={s} />} title="Redo (⇧⌘Z)" onClick={() => run('redo')} disabled={!canRedo} />
        </div>
        <div className="rgroup">
          <B icon={<ClipboardPaste size={s} />} title="Paste (⌘V)" onClick={() => run('paste')} disabled={!isGrid} />
          <B icon={<Scissors size={s} />} title="Cut (⌘X)" onClick={() => run('cut')} disabled={!isGrid} />
          <B icon={<Copy size={s} />} title="Copy (⌘C)" onClick={() => run('copy')} disabled={!isGrid} />
        </div>
        <div className="rgroup">
          <select className="tselect" style={{ width: 132 }} value={st?.font ?? ''} disabled={!isGrid} onChange={(e) => run('format:font', e.target.value)} title="Font">
            <option value="">Default font</option>
            {FONTS.map((f) => (
              <option key={f} value={f}>
                {f}
              </option>
            ))}
          </select>
          <select className="tselect" style={{ width: 58 }} value={st?.size ?? ''} disabled={!isGrid} onChange={(e) => run('format:size', e.target.value)} title="Font size">
            <option value="">13</option>
            {SIZES.map((n) => (
              <option key={n} value={n}>
                {n}
              </option>
            ))}
          </select>
        </div>
        <div className="rgroup">
          <B icon={<Bold size={s} />} title="Bold (⌘B)" on={st?.bold} onClick={() => run('format:bold')} disabled={!isGrid} />
          <B icon={<Italic size={s} />} title="Italic (⌘I)" on={st?.italic} onClick={() => run('format:italic')} disabled={!isGrid} />
          <B icon={<Underline size={s} />} title="Underline (⌘U)" on={st?.underline} onClick={() => run('format:underline')} disabled={!isGrid} />
          <B icon={<Strikethrough size={s} />} title="Strikethrough" on={st?.strike} onClick={() => run('format:strike')} disabled={!isGrid} />
          <ColorButton icon={<Baseline size={s} />} title="Text color" color={st?.color ?? '#1d1d1f'} onPick={(c) => run('format:color', c)} />
          <ColorButton icon={<PaintBucket size={s} />} title="Fill color" color={st?.fill ?? '#ffffff'} onPick={(c) => run('format:fill', c)} />
          <B icon={<Type size={s} />} title="Remove fill" onClick={() => run('format:fill', '')} disabled={!isGrid} label="No fill" />
        </div>
        <div className="rgroup">
          <B icon={<TextAlignStart size={s} />} title="Align left" on={st?.align === 'left'} onClick={() => run('format:align', 'left')} disabled={!isGrid} />
          <B icon={<TextAlignCenter size={s} />} title="Center" on={st?.align === 'center'} onClick={() => run('format:align', 'center')} disabled={!isGrid} />
          <B icon={<TextAlignEnd size={s} />} title="Align right" on={st?.align === 'right'} onClick={() => run('format:align', 'right')} disabled={!isGrid} />
          <B icon={<TextWrap size={s} />} title="Wrap text" on={st?.wrap} onClick={() => run('format:wrap')} disabled={!isGrid} />
          <B icon={<TableCellsMerge size={s} />} title="Merge & center / unmerge" onClick={() => run('format:merge')} disabled={!isGrid} />
        </div>
        <div className="rgroup">
          <select className="tselect" style={{ width: 118 }} value={fmtValue} disabled={!isGrid} onChange={(e) => run('format:numfmt', e.target.value)} title="Number format">
            {NUMBER_FORMATS.map((f) => (
              <option key={f.label} value={f.fmt}>
                {f.label}
              </option>
            ))}
            <option value="custom">Custom…</option>
          </select>
          <B icon={<DollarSign size={s} />} title="Currency" onClick={() => run('format:numfmt', '$#,##0.00')} disabled={!isGrid} />
          <B icon={<Percent size={s} />} title="Percent" onClick={() => run('format:numfmt', '0%')} disabled={!isGrid} />
          <B icon={<span style={{ fontWeight: 700 }}>,</span>} title="Thousands separator" onClick={() => run('format:numfmt', '#,##0')} disabled={!isGrid} />
          <B icon={<span style={{ fontSize: 11 }}>.0→.00</span>} title="Increase decimals" onClick={() => run('format:decimals', '1')} disabled={!isGrid} />
          <B icon={<span style={{ fontSize: 11 }}>.00→.0</span>} title="Decrease decimals" onClick={() => run('format:decimals', '-1')} disabled={!isGrid} />
        </div>
        <div className="rgroup">
          <select className="tselect" style={{ width: 104 }} value="" disabled={!isGrid} onChange={(e) => e.target.value && run('format:borders', e.target.value)} title="Borders">
            <option value="">Borders…</option>
            <option value="all">All borders</option>
            <option value="outside">Outside</option>
            <option value="bottom">Bottom</option>
            <option value="none">No borders</option>
          </select>
          <B icon={<Eraser size={s} />} title="Clear formats" onClick={() => run('clearFormats')} disabled={!isGrid} />
        </div>
        <div className="rgroup">
          <B icon={<Sigma size={s} />} title="AutoSum" onClick={() => run('autosum')} disabled={!isGrid} label="AutoSum" />
        </div>
      </div>
    )
  }
  if (tab === 'insert') {
    return (
      <div className="ribbon">
        <div className="rgroup">
          <B icon={<ChartColumn size={s} />} title="Insert chart from selection" onClick={() => run('insert:chart')} disabled={!isGrid} label="Chart" />
          <B icon={<Gauge size={s} />} title="KPI card" onClick={() => run('insert:kpi')} label="KPI card" />
          <B icon={<LayoutDashboard size={s} />} title="New dashboard" onClick={() => run('insert:dashboard')} label="Dashboard" />
        </div>
        <div className="rgroup">
          <B icon={<FilePlus size={s} />} title="New sheet" onClick={() => run('insert:sheet')} label="Sheet" />
        </div>
        <div className="rgroup">
          <B icon={<BetweenVerticalStart size={s} />} title="Insert row above" onClick={() => run('insert:rowAbove')} disabled={!isGrid} label="Row" />
          <B icon={<BetweenHorizontalStart size={s} />} title="Insert column left" onClick={() => run('insert:colLeft')} disabled={!isGrid} label="Column" />
          <B icon={<Rows3 size={s} />} title="Delete selected rows" onClick={() => run('delete:rows')} disabled={!isGrid} label="Delete rows" />
          <B icon={<Columns3 size={s} />} title="Delete selected columns" onClick={() => run('delete:cols')} disabled={!isGrid} label="Delete columns" />
        </div>
        <div className="rgroup">
          <B icon={<SquareFunction size={s} />} title="Insert function" onClick={() => run('insert:function')} disabled={!isGrid} label="Function" />
        </div>
      </div>
    )
  }
  if (tab === 'data') {
    return (
      <div className="ribbon">
        <div className="rgroup">
          <B icon={<ArrowDownAZ size={s} />} title="Sort ascending" onClick={() => run('data:sortAsc')} disabled={!isGrid} label="Sort A→Z" />
          <B icon={<ArrowDownZA size={s} />} title="Sort descending" onClick={() => run('data:sortDesc')} disabled={!isGrid} label="Sort Z→A" />
          <B icon={<Funnel size={s} />} title="Filter (⇧⌘L)" onClick={() => run('data:filter')} disabled={!isGrid} label="Filter" />
          <B icon={<ListX size={s} />} title="Remove duplicate rows" onClick={() => run('data:removeDuplicates')} disabled={!isGrid} label="Remove duplicates" />
        </div>
        <div className="rgroup">
          <B icon={<Highlighter size={s} />} title="Conditional formatting" onClick={() => run('format:conditional')} disabled={!isGrid} label="Conditional formatting" />
          <B icon={<Search size={s} />} title="Find & replace (⌘F)" onClick={() => run('find')} disabled={!isGrid} label="Find" />
        </div>
      </div>
    )
  }
  return (
    <div className="ribbon">
      <div className="rgroup">
        <B icon={<Snowflake size={s} />} title="Freeze top row" onClick={() => run('view:freezeTop')} disabled={!isGrid} label="Top row" />
        <B icon={<Snowflake size={s} />} title="Freeze first column" onClick={() => run('view:freezeFirst')} disabled={!isGrid} label="First column" />
        <B icon={<Snowflake size={s} />} title="Freeze panes above and left of the selection" onClick={() => run('view:freezeSelection')} disabled={!isGrid} label="At selection" />
        <B icon={<Snowflake size={s} />} title="Unfreeze" onClick={() => run('view:unfreeze')} disabled={!isGrid} label="Unfreeze" />
      </div>
      <div className="rgroup">
        <B icon={<Grid3x3 size={s} />} title="Gridlines" on={showGridlines} onClick={() => run('view:gridlines')} label="Gridlines" />
        <B icon={<Eye size={s} />} title="Show formulas" on={showFormulas} onClick={() => run('view:formulas')} label="Formulas" />
      </div>
      <div className="rgroup">
        <B icon={<Printer size={s} />} title="Print (⌘P)" onClick={() => run('print')} disabled={!isGrid} label="Print" />
      </div>
    </div>
  )
}
