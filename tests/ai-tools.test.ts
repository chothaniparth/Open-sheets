import { describe, expect, it } from 'vitest'
import { runTool, workbookContext, type ToolHost } from '../src/renderer/src/ai/tools'
import { Workbook, type GridSheet } from '../src/renderer/src/engine/workbook'

function host() {
  const wb = Workbook.create()
  let active = wb.grids()[0] as GridSheet
  const h: ToolHost = {
    wb: () => wb,
    activeGrid: () => active,
    selection: () => ({ r0: 0, c0: 0, r1: 0, c1: 0 }),
    setActive: (id) => (active = wb.tab(id) as GridSheet)
  }
  return { wb, h }
}

describe('AI tools', () => {
  it('writes, reads, formats and charts data like the model would', async () => {
    const { wb, h } = host()
    const res = (await runTool(h, 'write_range', {
      start: 'A1',
      values: [
        ['Region', 'Sales'],
        ['North', '1200'],
        ['South', '800'],
        ['Total', '=SUM(B2:B3)']
      ]
    })) as { written: string }
    expect(res.written).toBe('A1:B4')

    const read = (await runTool(h, 'read_range', { range: 'A1:B4' })) as { values: unknown[][]; formulas: Record<string, string> }
    expect(read.values[3][1]).toBe(2000)
    expect(read.formulas.B4).toBe('=SUM(B2:B3)')

    await runTool(h, 'format_range', { range: 'A1:B1', bold: true, fillColor: '#dbeafe', borders: 'bottom' })
    expect(wb.style(wb.grids()[0], 0, 0)).toMatchObject({ bold: true, fill: '#dbeafe' })

    const chart = (await runTool(h, 'create_chart', { chartType: 'column', range: 'A1:B3', title: 'Sales' })) as { dashboard: string }
    expect(chart.dashboard).toBe('Dashboard')
    const kpi = (await runTool(h, 'add_kpi_card', { label: 'Total', formula: '=SUM(Sheet1!B2:B3)' })) as { value: number }
    expect(kpi.value).toBe(2000)

    await runTool(h, 'sort_range', { range: 'A1:B3', column: 'B', descending: false, hasHeader: true })
    expect(wb.value(wb.grids()[0], 1, 0)).toBe('South')

    // every AI edit is undoable
    wb.undo()
    expect(wb.value(wb.grids()[0], 1, 0)).toBe('North')
  })

  it('reports formula errors and bad sheet names back to the model', async () => {
    const { h } = host()
    const res = (await runTool(h, 'write_range', { start: 'A1', values: [['=1/0']] })) as { errors?: string[] }
    expect(res.errors?.[0]).toContain('#DIV/0!')
    await expect(runTool(h, 'read_range', { sheet: 'Nope', range: 'A1' })).rejects.toThrow(/No sheet named/)
  })

  it('creates sheets and builds context', async () => {
    const { wb, h } = host()
    await runTool(h, 'create_sheet', { name: 'Report' })
    expect(h.activeGrid().name).toBe('Report')
    const info = (await runTool(h, 'get_workbook_info', {})) as { tabs: { name: string }[] }
    expect(info.tabs.map((t) => t.name)).toEqual(['Sheet1', 'Report'])
    expect(workbookContext(wb, h.activeGrid(), null)).toContain('Active sheet: Report')
  })
})
