import type { EChartsOption, SeriesOption } from 'echarts'
import { DetailedCellError, type CellValue } from 'hyperformula'
import type { ChartSpec } from '../engine/workbook'

// Validated categorical order (CVD-checked for adjacent pairs); dark steps chosen for the dark surface.
export const SERIES_LIGHT = ['#2a78d6', '#eb6834', '#1baf7a', '#eda100', '#e87ba4', '#008300', '#6250d6', '#e34948']
export const SERIES_DARK = ['#3987e5', '#d95926', '#199e70', '#c98500', '#d55181', '#008300', '#9085e9', '#e66767']
const MAX_SERIES = 8

interface Ink {
  primary: string
  secondary: string
  muted: string
  grid: string
  surface: string
}

export function inkFor(dark: boolean): Ink {
  return dark
    ? { primary: '#f0efec', secondary: '#c3c2b7', muted: '#8c8b84', grid: '#2e2e2c', surface: '#1a1a19' }
    : { primary: '#0b0b0b', secondary: '#52514e', muted: '#8a8984', grid: '#ebeae6', surface: '#fcfcfb' }
}

const toNum = (v: CellValue): number | null => (typeof v === 'number' && Number.isFinite(v) ? v : typeof v === 'boolean' ? Number(v) : null)
const toLabel = (v: CellValue): string => (v instanceof DetailedCellError ? v.value : v === null ? '' : String(v))

export interface ChartData {
  categories: string[]
  series: { name: string; data: (number | null)[] }[]
}

/** Interpret a block of cells as categories × series (first row = names, first column = labels). */
export function shapeData(values: CellValue[][], spec: Pick<ChartSpec, 'seriesBy' | 'headers'>): ChartData {
  let m = values.map((row) => [...row])
  if (spec.seriesBy === 'rows' && m.length) m = m[0].map((_, c) => m.map((row) => row[c]))
  if (!m.length || !m[0].length) return { categories: [], series: [] }

  const header = spec.headers ? m[0] : null
  const body = spec.headers ? m.slice(1) : m
  // First column holds labels if it has any non-numeric content.
  const labelCol = m[0].length > 1 && body.some((row) => row[0] !== null && toNum(row[0]) === null)
  const startCol = labelCol ? 1 : 0
  const categories = body.map((row, i) => (labelCol ? toLabel(row[0]) : String(i + 1)))
  const series: ChartData['series'] = []
  for (let c = startCol; c < m[0].length && series.length < MAX_SERIES; c++) {
    series.push({
      name: header ? toLabel(header[c]) || `Series ${series.length + 1}` : `Series ${series.length + 1}`,
      data: body.map((row) => toNum(row[c]))
    })
  }
  return { categories, series }
}

const fmtNum = (v: number) => (Math.abs(v) >= 1e6 ? `${+(v / 1e6).toFixed(1)}M` : Math.abs(v) >= 1e3 ? `${+(v / 1e3).toFixed(1)}K` : `${+v.toFixed(2)}`)

export function buildChartOption(spec: ChartSpec, values: CellValue[][], dark: boolean, title?: string): EChartsOption {
  const { categories, series } = shapeData(values, spec)
  const ink = inkFor(dark)
  const colors = dark ? SERIES_DARK : SERIES_LIGHT
  const showLegend = spec.legend && series.length >= 2
  const base: EChartsOption = {
    color: colors,
    backgroundColor: 'transparent',
    textStyle: { color: ink.secondary, fontFamily: '-apple-system, "Helvetica Neue", sans-serif' },
    title: title ? { text: title, left: 0, top: 0, textStyle: { color: ink.primary, fontSize: 13, fontWeight: 600 } } : undefined,
    legend: showLegend ? { top: title ? 22 : 0, left: 0, icon: 'roundRect', itemWidth: 10, itemHeight: 10, textStyle: { color: ink.secondary } } : undefined,
    tooltip: {
      backgroundColor: dark ? '#262624' : '#ffffff',
      borderColor: ink.grid,
      textStyle: { color: ink.primary, fontSize: 12 },
      valueFormatter: (v) => (typeof v === 'number' ? v.toLocaleString() : String(v ?? ''))
    },
    animationDuration: 300
  }
  const top = (title ? 24 : 4) + (showLegend ? 26 : 0) + 8
  const grid = { left: 8, right: 16, top, bottom: 8, containLabel: true }
  const axisCommon = {
    axisLine: { lineStyle: { color: ink.grid } },
    axisTick: { show: false },
    axisLabel: { color: ink.muted, fontSize: 11 },
    splitLine: { lineStyle: { color: ink.grid } }
  }
  const valueAxis = { type: 'value' as const, ...axisCommon, axisLine: { show: false }, axisLabel: { ...axisCommon.axisLabel, formatter: (v: number) => fmtNum(v) } }
  const catAxis = { type: 'category' as const, data: categories, ...axisCommon, splitLine: { show: false } }

  switch (spec.type) {
    case 'column':
    case 'stackedColumn':
    case 'bar': {
      const horizontal = spec.type === 'bar'
      const stacked = spec.type === 'stackedColumn'
      const radius = horizontal ? [0, 4, 4, 0] : [4, 4, 0, 0]
      return {
        ...base,
        tooltip: { ...base.tooltip, trigger: 'axis', axisPointer: { type: 'shadow' } },
        grid,
        xAxis: horizontal ? valueAxis : catAxis,
        yAxis: horizontal ? { ...catAxis, inverse: true } : valueAxis,
        series: series.map(
          (s, i): SeriesOption => ({
            type: 'bar',
            name: s.name,
            data: s.data,
            stack: stacked ? 'total' : undefined,
            barMaxWidth: 36,
            barGap: '8%',
            itemStyle: {
              borderRadius: stacked && i < series.length - 1 ? 0 : radius,
              borderColor: stacked ? ink.surface : undefined,
              borderWidth: stacked ? 1 : 0
            },
            emphasis: { focus: 'series' }
          })
        )
      }
    }
    case 'line':
    case 'area':
      return {
        ...base,
        tooltip: { ...base.tooltip, trigger: 'axis', axisPointer: { type: 'line', lineStyle: { color: ink.muted } } },
        grid,
        xAxis: { ...catAxis, boundaryGap: false },
        yAxis: valueAxis,
        series: series.map(
          (s): SeriesOption => ({
            type: 'line',
            name: s.name,
            data: s.data,
            smooth: spec.smooth ?? false,
            showSymbol: categories.length <= 24,
            symbolSize: 8,
            lineStyle: { width: 2 },
            areaStyle: spec.type === 'area' ? { opacity: 0.18 } : undefined,
            emphasis: { focus: 'series' }
          })
        )
      }
    case 'pie':
    case 'doughnut':
    case 'funnel': {
      const s = series[0]
      const data = categories.map((name, i) => ({ name, value: s?.data[i] ?? 0 })).filter((d) => d.value !== null && d.value > 0)
      // Past eight slices the tail folds into "Other" rather than inventing new hues.
      const sorted = [...data].sort((a, b) => b.value - a.value)
      const folded = sorted.length > MAX_SERIES ? [...sorted.slice(0, MAX_SERIES - 1), { name: 'Other', value: sorted.slice(MAX_SERIES - 1).reduce((t, d) => t + d.value, 0) }] : sorted
      if (spec.type === 'funnel') {
        return {
          ...base,
          tooltip: { ...base.tooltip, trigger: 'item' },
          series: [{ type: 'funnel', top: title ? 32 : 8, bottom: 8, left: '10%', width: '80%', gap: 2, data: folded, label: { color: ink.primary }, itemStyle: { borderColor: ink.surface, borderWidth: 2 } }]
        }
      }
      return {
        ...base,
        legend: spec.legend ? { orient: 'vertical', right: 0, top: 'middle', icon: 'roundRect', itemWidth: 10, itemHeight: 10, textStyle: { color: ink.secondary } } : undefined,
        tooltip: { ...base.tooltip, trigger: 'item', valueFormatter: undefined, formatter: '{b}: {c} ({d}%)' },
        series: [
          {
            type: 'pie',
            radius: spec.type === 'doughnut' ? ['48%', '72%'] : [0, '72%'],
            center: [spec.legend ? '40%' : '50%', title ? '55%' : '50%'],
            data: folded,
            itemStyle: { borderColor: ink.surface, borderWidth: 2, borderRadius: 4 },
            label: { show: !spec.legend, color: ink.secondary },
            emphasis: { scale: true, scaleSize: 4 }
          }
        ]
      }
    }
    case 'scatter': {
      // first numeric series = x; the rest = y series (capped at three for legibility)
      const [xs, ...ys] = series
      return {
        ...base,
        tooltip: { ...base.tooltip, trigger: 'item' },
        grid,
        xAxis: { ...valueAxis, name: xs?.name, nameLocation: 'middle', nameGap: 24, splitLine: { show: false } },
        yAxis: valueAxis,
        legend: ys.length >= 2 && spec.legend ? base.legend : undefined,
        series: ys.slice(0, 3).map(
          (s): SeriesOption => ({
            type: 'scatter',
            name: s.name,
            symbolSize: 9,
            itemStyle: { borderColor: ink.surface, borderWidth: 2 },
            data: s.data.map((y, i) => [xs.data[i], y])
          })
        )
      }
    }
    case 'radar': {
      const max = Math.max(1, ...series.flatMap((s) => s.data.map((v) => v ?? 0)))
      return {
        ...base,
        tooltip: { ...base.tooltip, trigger: 'item' },
        radar: {
          indicator: categories.map((name) => ({ name, max: max * 1.1 })),
          center: ['50%', title ? '58%' : '52%'],
          radius: '62%',
          axisName: { color: ink.secondary },
          splitLine: { lineStyle: { color: ink.grid } },
          splitArea: { show: false },
          axisLine: { lineStyle: { color: ink.grid } }
        },
        series: [{ type: 'radar', data: series.map((s) => ({ name: s.name, value: s.data.map((v) => v ?? 0) })), lineStyle: { width: 2 }, areaStyle: { opacity: 0.12 }, symbolSize: 6 }]
      }
    }
  }
}
