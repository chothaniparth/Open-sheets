import { ipcMain, type WebContents } from 'electron'
import { randomUUID } from 'node:crypto'
import { stepCountIs, streamText, tool, type ModelMessage } from 'ai'
import { z } from 'zod'
import type { ChatEvent, ChatRequest, ToolRequest } from '../shared/types'
import { createModel } from './providers'
import { resolveProvider } from './store'

const SYSTEM_PROMPT = `You are the AI assistant built into Open Sheets, a spreadsheet app for macOS.
You can read and change the user's open workbook with the tools provided.

How to work:
- Look before you write: call read_range or get_workbook_info when you need to know what the data looks like.
- Prefer live formulas (e.g. =SUM(B2:B20)) over hard-coded results so the sheet stays correct when data changes. The formula engine is Excel-compatible (SUM, AVERAGE, IF, XLOOKUP, VLOOKUP, INDEX/MATCH, SUMIFS, COUNTIFS, TEXT, DATE, etc.).
- When writing a table, put headers in the first row, then format the header (bold, fill) and set sensible number formats.
- For "dashboard" requests, build summary tables with formulas first, then add charts and KPI cards that reference them.
- Keep replies short. Say what you changed and where (sheet + range). Use Markdown; put formulas in backticks.
- If the request is ambiguous or would overwrite existing data, ask first.`

const sheet = z.string().optional().describe('Sheet name. Omit to use the active sheet.')
const range = z.string().describe('A1-style range such as "A1:D20" or a single cell "B3". May include a sheet prefix like "Sales!A1:C10".')

type Sender = WebContents
const pending = new Map<string, { resolve: (v: unknown) => void; reject: (e: Error) => void; timer: NodeJS.Timeout }>()

/** Run a tool inside the renderer (which owns the workbook) and wait for its answer. */
function callRenderer(sender: Sender, name: string, input: unknown): Promise<unknown> {
  const reqId = randomUUID()
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pending.delete(reqId)
      reject(new Error(`Tool ${name} timed out.`))
    }, 60_000)
    pending.set(reqId, { resolve, reject, timer })
    const req: ToolRequest = { reqId, name, input }
    sender.send('ai:tool-request', req)
  })
}

function buildTools(sender: Sender) {
  const run = (name: string) => (input: unknown) => callRenderer(sender, name, input)
  return {
    get_workbook_info: tool({
      description: 'List sheets and dashboards with their used ranges, plus the active sheet and current selection.',
      inputSchema: z.object({}),
      execute: run('get_workbook_info')
    }),
    read_range: tool({
      description: 'Read cell values (and formulas where present) from a range. Large ranges are truncated.',
      inputSchema: z.object({ sheet, range }),
      execute: run('read_range')
    }),
    write_range: tool({
      description:
        'Write a 2D block of cells starting at a top-left cell. Each cell is text: numbers like "42" or "3.5", booleans "TRUE"/"FALSE", dates "2026-01-31", formulas starting with "=", or "" to clear.',
      inputSchema: z.object({
        sheet,
        start: z.string().describe('Top-left cell, e.g. "A1".'),
        values: z.array(z.array(z.string())).describe('Rows of cells.')
      }),
      execute: run('write_range')
    }),
    clear_range: tool({
      description: 'Clear the contents (not formatting) of a range.',
      inputSchema: z.object({ sheet, range }),
      execute: run('clear_range')
    }),
    format_range: tool({
      description: 'Apply formatting to a range. Only the fields you pass are changed.',
      inputSchema: z.object({
        sheet,
        range,
        bold: z.boolean().optional(),
        italic: z.boolean().optional(),
        underline: z.boolean().optional(),
        fontSize: z.number().optional(),
        fontColor: z.string().optional().describe('CSS hex color, e.g. "#1f2937".'),
        fillColor: z.string().optional().describe('CSS hex background color, or "none".'),
        horizontalAlign: z.enum(['left', 'center', 'right']).optional(),
        wrap: z.boolean().optional(),
        numberFormat: z
          .string()
          .optional()
          .describe('Excel number format code, e.g. "#,##0.00", "$#,##0", "0.0%", "yyyy-mm-dd", "General".'),
        borders: z.enum(['all', 'outside', 'bottom', 'none']).optional()
      }),
      execute: run('format_range')
    }),
    set_column_width: tool({
      description: 'Set the width in pixels of one or more columns, or "auto" to fit contents.',
      inputSchema: z.object({
        sheet,
        columns: z.string().describe('Column letters, e.g. "A" or "B:E".'),
        width: z.union([z.number(), z.literal('auto')])
      }),
      execute: run('set_column_width')
    }),
    create_sheet: tool({
      description: 'Add a new empty worksheet and make it active.',
      inputSchema: z.object({ name: z.string() }),
      execute: run('create_sheet')
    }),
    sort_range: tool({
      description: 'Sort the rows of a range by one column.',
      inputSchema: z.object({
        sheet,
        range,
        column: z.string().describe('Column letter to sort by, e.g. "C".'),
        descending: z.boolean().optional(),
        hasHeader: z.boolean().optional().describe('True if the first row is a header row.')
      }),
      execute: run('sort_range')
    }),
    freeze_panes: tool({
      description: 'Freeze the top N rows and left M columns of a sheet (0 to unfreeze).',
      inputSchema: z.object({ sheet, rows: z.number().int().min(0), columns: z.number().int().min(0) }),
      execute: run('freeze_panes')
    }),
    create_chart: tool({
      description:
        'Create a chart from a data range and place it on a dashboard. The first row should hold series names and the first column category labels.',
      inputSchema: z.object({
        chartType: z.enum(['column', 'stackedColumn', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'radar', 'funnel']),
        range,
        title: z.string().optional(),
        dashboard: z.string().optional().describe('Dashboard name. Created if it does not exist. Defaults to "Dashboard".')
      }),
      execute: run('create_chart')
    }),
    add_kpi_card: tool({
      description: 'Add a KPI card (a big single number) to a dashboard. The value comes from a formula that is kept live.',
      inputSchema: z.object({
        label: z.string(),
        formula: z.string().describe('Formula with explicit sheet references, e.g. "=SUM(Sales!C2:C200)".'),
        numberFormat: z.string().optional().describe('e.g. "$#,##0", "0.0%", "#,##0".'),
        dashboard: z.string().optional()
      }),
      execute: run('add_kpi_card')
    })
  }
}

const sessions = new Map<string, ModelMessage[]>()
const running = new Map<string, AbortController>()

function errorMessage(err: unknown): string {
  if (err instanceof Error) {
    const anyErr = err as any
    const body = typeof anyErr.responseBody === 'string' ? anyErr.responseBody : ''
    const fromBody = body.match(/"message"\s*:\s*"([^"]+)"/)?.[1]
    return fromBody ? `${err.message} — ${fromBody}` : err.message
  }
  return String(err)
}

async function runChat(sender: Sender, req: ChatRequest) {
  const key = `${sender.id}:${req.chatId}`
  const emit = (ev: ChatEvent) => {
    if (!sender.isDestroyed()) sender.send('ai:event', ev)
  }
  const [providerId, ...rest] = req.model.split('::')
  const modelId = rest.join('::')
  const controller = new AbortController()
  running.set(key, controller)

  try {
    const provider = resolveProvider(providerId)
    if (!provider.apiKey && !provider.keyOptional) {
      throw new Error(`No API key for ${provider.name}. Add one in Settings → AI Providers.`)
    }
    const history = sessions.get(key) ?? []
    const userMessage: ModelMessage = {
      role: 'user',
      content: `<workbook_context>\n${req.context}\n</workbook_context>\n\n${req.text}`
    }
    const messages = [...history, userMessage]

    const result = streamText({
      model: createModel(provider, modelId),
      system: SYSTEM_PROMPT,
      messages,
      tools: buildTools(sender),
      stopWhen: stepCountIs(25),
      maxOutputTokens: provider.kind === 'anthropic' ? 32000 : undefined,
      abortSignal: controller.signal
    })

    for await (const part of result.fullStream) {
      switch (part.type) {
        case 'text-delta':
          emit({ chatId: req.chatId, type: 'text', text: part.text })
          break
        case 'reasoning-delta':
          emit({ chatId: req.chatId, type: 'reasoning', text: part.text })
          break
        case 'tool-call':
          emit({ chatId: req.chatId, type: 'tool-call', id: part.toolCallId, name: part.toolName, input: part.input })
          break
        case 'tool-result':
          emit({ chatId: req.chatId, type: 'tool-result', id: part.toolCallId, output: part.output })
          break
        case 'tool-error':
          emit({ chatId: req.chatId, type: 'tool-error', id: part.toolCallId, error: errorMessage(part.error) })
          break
        case 'error':
          throw part.error
      }
    }
    const responseMessages = await result.responseMessages
    sessions.set(key, [...messages, ...responseMessages])
    const usage = await result.totalUsage
    emit({ chatId: req.chatId, type: 'done', usage: { input: usage.inputTokens, output: usage.outputTokens } })
  } catch (err) {
    if (controller.signal.aborted) {
      emit({ chatId: req.chatId, type: 'done' })
    } else {
      emit({ chatId: req.chatId, type: 'error', message: errorMessage(err) })
    }
  } finally {
    running.delete(key)
  }
}

export function registerAiIpc() {
  ipcMain.handle('ai:chat', (e, req: ChatRequest) => {
    void runChat(e.sender, req)
  })
  ipcMain.handle('ai:stop', (e, chatId: string) => {
    running.get(`${e.sender.id}:${chatId}`)?.abort()
  })
  ipcMain.handle('ai:reset', (e, chatId: string) => {
    sessions.delete(`${e.sender.id}:${chatId}`)
  })
  ipcMain.on('ai:tool-response', (_e, msg: { reqId: string; ok: boolean; result?: unknown; error?: string }) => {
    const p = pending.get(msg.reqId)
    if (!p) return
    clearTimeout(p.timer)
    pending.delete(msg.reqId)
    if (msg.ok) p.resolve(msg.result)
    else p.reject(new Error(msg.error ?? 'Tool failed'))
  })
}
