import type { CellAddr, Range } from '../engine/addr'

export interface Sel {
  anchor: CellAddr
  focus: CellAddr
}

export interface EditState {
  tabId: string
  r: number
  c: number
  text: string
  original: string
  origin: 'cell' | 'bar'
  /** 'enter' = started by typing (arrow keys commit), 'edit' = F2/double-click (arrows move caret) */
  mode: 'enter' | 'edit'
  /** Reference most recently inserted by clicking cells while editing a formula. */
  pick?: { start: number; end: number; anchor: CellAddr; tabId: string }
}

export interface ClipboardData {
  tabId: string
  range: Range
  tsv: string
  cut: boolean
}
