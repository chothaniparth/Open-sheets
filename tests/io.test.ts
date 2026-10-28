import { describe, expect, it } from 'vitest'
import { exportWorkbook } from '../src/renderer/src/io/exporter'
import { importFile } from '../src/renderer/src/io/importer'
import { Workbook, type GridSheet } from '../src/renderer/src/engine/workbook'

function sample() {
  const wb = Workbook.create()
  const tab = wb.grids()[0] as GridSheet
  wb.writeTexts(tab, 0, 0, [['Region', 'Sales', 'Date'], ['North', '1,200', '2026-01-31'], ['South', '800', '2026-02-28'], ['Total', '=SUM(B2:B3)', '']])
  wb.applyStyle(tab, { r0: 0, c0: 0, r1: 0, c1: 2 }, { bold: true, fill: '#dbeafe' })
  wb.merge(tab, { r0: 5, c0: 0, r1: 5, c1: 2 })
  wb.setColWidth(tab, [0], 160)
  wb.setFreeze(tab, 1, 0)
  wb.setCondRules(tab, [{ id: 'r1', range: 'B2:B3', type: 'gt', a: '1000', fill: '#c6efce' }])
  wb.addWidget('Overview', { type: 'chart', w: 6, h: 4, title: 'Sales', chart: { type: 'column', range: 'Sheet1!A1:B3', seriesBy: 'columns', headers: true, legend: true } })
  return { wb, tab }
}

describe('file formats', () => {
  it('round-trips xlsx with styles, merges, freeze and dashboards', async () => {
    const { wb, tab } = sample()
    const bytes = await exportWorkbook(wb, 'xlsx', tab)
    const data = await importFile('book.xlsx', bytes)
    const back = Workbook.fromData(data)
    const t = back.grids()[0]
    expect(back.value(t, 3, 1)).toBe(2000)
    expect(back.raw(t, 3, 1)).toBe('=SUM(B2:B3)')
    expect(back.display(t, 1, 2)).toBe('2026-01-31')
    expect(back.display(t, 1, 1)).toBe('1,200')
    expect(t.styles['0:0']).toMatchObject({ bold: true, fill: '#dbeafe' })
    expect(t.merges).toEqual([{ r0: 5, c0: 0, r1: 5, c1: 2 }])
    expect(t.colWidths[0]).toBe(160)
    expect(t.freeze).toEqual({ rows: 1, cols: 0 })
    expect(t.cond[0].type).toBe('gt')
    expect(back.tabs.map((x) => x.name)).toEqual(['Sheet1', 'Overview'])
  })

  it.each(['xls', 'xlsb', 'ods', 'fods', 'xml', 'numbers'])('round-trips values and formulas through .%s', async (ext) => {
    const { wb, tab } = sample()
    const bytes = await exportWorkbook(wb, ext, tab)
    const back = Workbook.fromData(await importFile(`book.${ext}`, bytes))
    const t = back.grids()[0]
    expect(back.value(t, 1, 0)).toBe('North')
    expect(back.value(t, 3, 1)).toBe(2000)
  })

  it('exports and re-imports CSV', async () => {
    const { wb, tab } = sample()
    const bytes = await exportWorkbook(wb, 'csv', tab)
    const text = new TextDecoder().decode(bytes)
    expect(text).toContain('North,"1,200",2026-01-31')
    const back = Workbook.fromData(await importFile('book.csv', bytes))
    const t = back.grids()[0]
    expect(back.value(t, 3, 1)).toBe(2000)
  })
})
