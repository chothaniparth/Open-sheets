import * as XLSX from 'xlsx'
import { DetailedCellError, type CellValue } from 'hyperformula'

const SSF = XLSX.SSF

export function isDateFormat(fmt?: string): boolean {
  if (!fmt) return false
  return SSF.is_date(fmt)
}

/** Format a computed cell value for display using an Excel number-format code. */
export function formatValue(value: CellValue | undefined, fmt?: string): string {
  if (value === null || value === undefined) return ''
  if (value instanceof DetailedCellError) return value.value
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'string') return value
  if (!Number.isFinite(value)) return '#NUM!'
  try {
    return SSF.format(fmt && fmt !== '@' ? fmt : 'General', value)
  } catch {
    return String(value)
  }
}

export const NUMBER_FORMATS: { label: string; fmt: string }[] = [
  { label: 'General', fmt: 'General' },
  { label: 'Number', fmt: '#,##0.00' },
  { label: 'Currency', fmt: '$#,##0.00' },
  { label: 'Accounting', fmt: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
  { label: 'Percentage', fmt: '0.00%' },
  { label: 'Scientific', fmt: '0.00E+00' },
  { label: 'Fraction', fmt: '# ?/?' },
  { label: 'Short Date', fmt: 'm/d/yyyy' },
  { label: 'ISO Date', fmt: 'yyyy-mm-dd' },
  { label: 'Long Date', fmt: 'dddd, mmmm d, yyyy' },
  { label: 'Time', fmt: 'h:mm:ss AM/PM' },
  { label: 'Date & Time', fmt: 'yyyy-mm-dd h:mm' },
  { label: 'Text', fmt: '@' }
]

/** Increase/decrease the number of decimals shown by a number format. */
export function adjustDecimals(fmt: string | undefined, delta: 1 | -1, sample?: number): string {
  let base = !fmt || fmt === 'General' ? '' : fmt
  if (!base) {
    const decimals = sample !== undefined && !Number.isInteger(sample) ? Math.min(String(sample).split('.')[1]?.length ?? 0, 10) : 0
    base = decimals ? `0.${'0'.repeat(decimals)}` : '0'
  }
  return base
    .split(';')
    .map((sec) => {
      const m = /\.(0+)/.exec(sec)
      if (m) {
        const n = m[1].length + delta
        return sec.slice(0, m.index) + (n > 0 ? `.${'0'.repeat(n)}` : '') + sec.slice(m.index + m[0].length)
      }
      const i = sec.lastIndexOf('0')
      if (delta < 0 || i < 0) return sec
      return `${sec.slice(0, i + 1)}.0${sec.slice(i + 1)}`
    })
    .join(';')
}

const DAY_MS = 86_400_000
const EPOCH = Date.UTC(1899, 11, 30)

export function dateToSerial(y: number, m: number, d: number, hh = 0, mm = 0, ss = 0): number {
  return (Date.UTC(y, m - 1, d, hh, mm, ss) - EPOCH) / DAY_MS
}

export function jsDateToSerial(date: Date): number {
  return (date.getTime() - EPOCH) / DAY_MS
}

export interface ParsedInput {
  raw: string | number | boolean | null
  fmt?: string
}

const CURRENCY = '$€£¥₹'

/** Interpret what the user typed, Excel-style (numbers with separators, %, currency, dates, times). */
export function parseInput(text: string): ParsedInput {
  const t = text.trim()
  if (text === '') return { raw: null }
  if (text.startsWith('=')) return { raw: text }
  if (text.startsWith("'")) return { raw: text }
  if (/^(true|false)$/i.test(t)) return { raw: t.toUpperCase() === 'TRUE' }

  const num = new RegExp(`^([-+])?([${CURRENCY}])?\\s?(\\d{1,3}(?:,\\d{3})+|\\d+)?(\\.\\d+)?\\s?(%)?$`).exec(t)
  if (num && (num[3] || num[4])) {
    const [, sign, cur, int = '0', dec = '', pct] = num
    let v = Number(`${int.replace(/,/g, '')}${dec}`)
    if (sign === '-') v = -v
    const decimals = dec ? dec.length - 1 : 0
    const d = decimals ? `.${'0'.repeat(decimals)}` : ''
    if (pct) return { raw: v / 100, fmt: `0${d}%` }
    if (cur) return { raw: v, fmt: `${cur}#,##0${d || '.00'}` }
    if (int.includes(',')) return { raw: v, fmt: `#,##0${d}` }
    return { raw: v }
  }
  if (/^[-+]?\d+(\.\d+)?e[-+]?\d+$/i.test(t)) return { raw: Number(t) }

  let m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t)
  if (m && validDate(+m[1], +m[2], +m[3])) {
    const hasTime = m[4] !== undefined
    return {
      raw: dateToSerial(+m[1], +m[2], +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0)),
      fmt: hasTime ? 'yyyy-mm-dd h:mm' : 'yyyy-mm-dd'
    }
  }
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(t)
  if (m) {
    const y = m[3].length === 2 ? 2000 + +m[3] - (+m[3] > 50 ? 100 : 0) : +m[3]
    if (validDate(y, +m[1], +m[2])) return { raw: dateToSerial(y, +m[1], +m[2]), fmt: 'm/d/yyyy' }
  }
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?$/i.exec(t)
  if (m) {
    let h = +m[1]
    const ampm = m[4]?.toLowerCase()
    if (ampm === 'pm' && h < 12) h += 12
    if (ampm === 'am' && h === 12) h = 0
    if (h < 24 && +m[2] < 60) {
      return {
        raw: (h * 3600 + +m[2] * 60 + +(m[3] ?? 0)) / 86400,
        fmt: ampm ? 'h:mm AM/PM' : m[3] ? 'h:mm:ss' : 'h:mm'
      }
    }
  }
  return { raw: text }
}

function validDate(y: number, m: number, d: number) {
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return false
  return d <= new Date(Date.UTC(y, m, 0)).getUTCDate()
}

/** Text shown in the editor for a stored value (keeps dates/percents human-readable). */
export function editTextFor(raw: unknown, fmt?: string): string {
  if (raw === null || raw === undefined) return ''
  if (typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE'
  if (typeof raw === 'number') {
    if (fmt && isDateFormat(fmt)) {
      const hasTime = /h|s/i.test(fmt.replace(/"[^"]*"/g, ''))
      const hasDate = /[dy]/i.test(fmt.replace(/"[^"]*"/g, ''))
      if (hasDate) return SSF.format(hasTime ? 'yyyy-mm-dd h:mm:ss' : 'yyyy-mm-dd', raw)
      return SSF.format('h:mm:ss', raw)
    }
    if (fmt && fmt.includes('%')) return `${+(raw * 100).toPrecision(15)}%`
    return String(+raw.toPrecision(15))
  }
  return String(raw)
}
