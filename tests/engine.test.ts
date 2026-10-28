import { describe, expect, it } from 'vitest'
import { findRefs, parseRange, shiftFormula } from '../src/renderer/src/engine/addr'
import { adjustDecimals, formatValue, parseInput } from '../src/renderer/src/engine/format'
import { Workbook, type GridSheet } from '../src/renderer/src/engine/workbook'

describe('addr', () => {
  it('shifts relative references only', () => {
    expect(shiftFormula('=A1+$B$2+C$3+$D4', 1, 1)).toBe('=B2+$B$2+D$3+$D5')
    expect(shiftFormula('=SUM(A1:A10)*LOG10(B1)', 2, 0)).toBe('=SUM(A3:A12)*LOG10(B3)')
    expect(shiftFormula('="A1"&Sheet1!A1', 1, 0)).toBe('="A1"&Sheet1!A2')
    expect(shiftFormula("='My Sheet'!B2", 0, -1)).toBe("='My Sheet'!A2")
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!')
  })
  it('parses ranges', () => {
    expect(parseRange('b2:a1')).toEqual({ r0: 0, c0: 0, r1: 1, c1: 1 })
    expect(parseRange('C:D', { rows: 10, cols: 5 })).toEqual({ r0: 0, c0: 2, r1: 9, c1: 3 })
  })
  it('finds refs in formulas', () => {
    expect(findRefs('=SUM(A1:B3)+Data!C4+"D5"').length).toBe(2)
  })
})

describe('format', () => {
  it('parses typed input like Excel', () => {
    expect(parseInput('1,234.50')).toEqual({ raw: 1234.5, fmt: '#,##0.00' })
    expect(parseInput('12%')).toEqual({ raw: 0.12, fmt: '0%' })
    expect(parseInput('$9')).toEqual({ raw: 9, fmt: '$#,##0.00' })
    expect(parseInput('2026-01-31').fmt).toBe('yyyy-mm-dd')
    expect(parseInput('2026-01-31').raw).toBe(46053)
    expect(parseInput('hello')).toEqual({ raw: 'hello' })
    expect(parseInput('TRUE')).toEqual({ raw: true })
  })
  it('formats numbers', () => {
    expect(formatValue(1234.5, '#,##0.00')).toBe('1,234.50')
    expect(formatValue(0.125, '0.0%')).toBe('12.5%')
    expect(formatValue(46053, 'yyyy-mm-dd')).toBe('2026-01-31')
    expect(adjustDecimals('#,##0.00', 1)).toBe('#,##0.000')
    expect(adjustDecimals('0%', 1)).toBe('0.0%')
    expect(adjustDecimals('0.0', -1)).toBe('0')
  })
})

describe('workbook', () => {
  const setup = () => {
    const wb = Workbook.create()
    const tab = wb.grids()[0] as GridSheet
    return { wb, tab }
  }

  it('computes formulas and supports undo/redo across content and formatting', () => {
    const { wb, tab } = setup()
    wb.writeTexts(tab, 0, 0, [['10'], ['20'], ['=SUM(A1:A2)']])
    expect(wb.value(tab, 2, 0)).toBe(30)
    wb.applyStyle(tab, { r0: 0, c0: 0, r1: 0, c1: 0 }, { bold: true })
    expect(wb.style(tab, 0, 0)?.bold).toBe(true)
    wb.setInput(tab, 0, 0, '15')
    expect(wb.value(tab, 2, 0)).toBe(35)
    wb.undo()
    expect(wb.value(tab, 2, 0)).toBe(30)
    wb.undo()
    expect(wb.tab(tab.id)!.kind === 'grid' && (wb.tab(tab.id) as GridSheet).styles['0:0']).toBeFalsy()
    wb.redo()
    expect((wb.tab(tab.id) as GridSheet).styles['0:0']?.bold).toBe(true)
    wb.redo()
    expect(wb.value(tab, 2, 0)).toBe(35)
  })

  it('fills series and shifts formulas', () => {
    const { wb, tab } = setup()
    wb.writeTexts(tab, 0, 0, [['1', '=A1*2'], ['2', '=A2*2']])
    wb.fill(tab, { r0: 0, c0: 0, r1: 1, c1: 1 }, { r0: 0, c0: 0, r1: 4, c1: 1 })
    expect(wb.value(tab, 4, 0)).toBe(5)
    expect(wb.raw(tab, 4, 1)).toBe('=A5*2')
    expect(wb.value(tab, 4, 1)).toBe(10)
  })

  it('sorts with header and keeps formulas pointing at their row', () => {
    const { wb, tab } = setup()
    wb.writeTexts(tab, 0, 0, [['Name', 'Score', 'Double'], ['b', '2', '=B2*2'], ['a', '9', '=B3*2'], ['c', '5', '=B4*2']])
    wb.sortRange(tab, { r0: 0, c0: 0, r1: 3, c1: 2 }, 1, true, true)
    expect(wb.value(tab, 1, 0)).toBe('a')
    expect(wb.value(tab, 1, 2)).toBe(18)
    expect(wb.value(tab, 3, 0)).toBe('b')
  })

  it('inserts rows and moves formatting with HF references', () => {
    const { wb, tab } = setup()
    wb.writeTexts(tab, 0, 0, [['5'], ['=A1+1']])
    wb.applyStyle(tab, { r0: 1, c0: 0, r1: 1, c1: 0 }, { fill: '#ff0000' })
    wb.insertRows(tab, 0, 2)
    const t = wb.tab(tab.id) as GridSheet
    expect(t.styles['3:0']?.fill).toBe('#ff0000')
    expect(wb.raw(t, 3, 0)).toBe('=A3+1')
    wb.undo()
    expect((wb.tab(tab.id) as GridSheet).styles['1:0']?.fill).toBe('#ff0000')
    expect(wb.raw(t, 1, 0)).toBe('=A1+1')
  })

  it('handles sheets and dashboards', () => {
    const { wb, tab } = setup()
    wb.writeTexts(tab, 0, 0, [['x', 'y'], ['a', '3']])
    const { dashboard } = wb.addWidget('Overview', { type: 'kpi', w: 3, h: 2, title: 'Total', kpi: { formula: '=SUM(Sheet1!B:B)' } })
    expect(wb.evaluate(dashboard.widgets[0].kpi!.formula)).toBe(3)
    wb.renameTab(tab.id, 'Data')
    expect((wb.findTab('Overview') as any).widgets[0].kpi.formula).toBe('=SUM(Data!B:B)')
    expect(wb.evaluate('=SUM(Data!B:B)')).toBe(3)
  })
})
