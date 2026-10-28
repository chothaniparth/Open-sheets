import { useEffect, useRef, useState } from 'react'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import { CircleCheck, CircleX, MessageSquarePlus, Send, Settings, Sparkles, Square, Wrench, X } from 'lucide-react'
import type { ChatEvent, SettingsSnapshot } from '@shared/types'

type Part = { kind: 'text'; text: string } | { kind: 'tool'; id: string; name: string; input: unknown; output?: unknown; error?: string }

interface Message {
  id: string
  role: 'user' | 'assistant' | 'error'
  text?: string
  parts?: Part[]
}

const TOOL_LABELS: Record<string, string> = {
  get_workbook_info: 'Looked at the workbook',
  read_range: 'Read cells',
  write_range: 'Wrote cells',
  clear_range: 'Cleared cells',
  format_range: 'Formatted cells',
  set_column_width: 'Resized columns',
  create_sheet: 'Created a sheet',
  sort_range: 'Sorted data',
  freeze_panes: 'Froze panes',
  create_chart: 'Created a chart',
  add_kpi_card: 'Added a KPI card'
}

const SUGGESTIONS = [
  'Summarize what is in this sheet',
  'Build a dashboard with charts and KPIs from this data',
  'Add a totals row with formulas and format it',
  'Clean up this table: header style, number formats, column widths'
]

const newChatId = () => crypto.randomUUID()

export function AIPanel({ snap, getContext, onOpenSettings, onClose }: { snap: SettingsSnapshot | null; getContext: () => string; onOpenSettings: () => void; onClose: () => void }) {
  const [messages, setMessages] = useState<Message[]>([])
  const [input, setInput] = useState('')
  const [running, setRunning] = useState(false)
  const [width, setWidth] = useState(() => {
    try {
      return Number(localStorage.getItem('ai-width')) || 380
    } catch {
      return 380
    }
  })
  const chatId = useRef(newChatId())
  const listRef = useRef<HTMLDivElement>(null)

  // models from enabled providers that have a key (or need none)
  const options = (snap?.providers ?? [])
    .filter((p) => p.enabled && (p.hasKey || p.keyOptional))
    .map((p) => ({ provider: p, models: p.pinnedModels.length ? p.pinnedModels : p.models }))
    .filter((g) => g.models.length)
  const allValues = options.flatMap((g) => g.models.map((m) => `${g.provider.id}::${m}`))
  const selected = snap?.settings.selectedModel && allValues.includes(snap.settings.selectedModel) ? snap.settings.selectedModel : allValues[0]

  useEffect(() => {
    return window.api.ai.onEvent((ev: ChatEvent) => {
      if (ev.chatId !== chatId.current) return
      setMessages((msgs) => {
        const out = [...msgs]
        let last = out[out.length - 1]
        if (ev.type === 'error') {
          out.push({ id: crypto.randomUUID(), role: 'error', text: ev.message })
          return out
        }
        if (ev.type === 'done') return out
        if (!last || last.role !== 'assistant') {
          last = { id: crypto.randomUUID(), role: 'assistant', parts: [] }
          out.push(last)
        } else {
          last = { ...last, parts: [...(last.parts ?? [])] }
          out[out.length - 1] = last
        }
        const parts = last.parts!
        if (ev.type === 'text') {
          const tail = parts[parts.length - 1]
          if (tail?.kind === 'text') parts[parts.length - 1] = { ...tail, text: tail.text + ev.text }
          else parts.push({ kind: 'text', text: ev.text })
        } else if (ev.type === 'tool-call') {
          parts.push({ kind: 'tool', id: ev.id, name: ev.name, input: ev.input })
        } else if (ev.type === 'tool-result' || ev.type === 'tool-error') {
          const i = parts.findIndex((p) => p.kind === 'tool' && p.id === ev.id)
          if (i >= 0) {
            const p = parts[i] as Extract<Part, { kind: 'tool' }>
            parts[i] = ev.type === 'tool-result' ? { ...p, output: ev.output } : { ...p, error: ev.error }
          }
        }
        return out
      })
      if (ev.type === 'done' || ev.type === 'error') setRunning(false)
    })
  }, [])

  useEffect(() => {
    listRef.current?.scrollTo({ top: listRef.current.scrollHeight })
  }, [messages])

  const send = (text: string) => {
    const t = text.trim()
    if (!t || running || !selected) return
    setMessages((m) => [...m, { id: crypto.randomUUID(), role: 'user', text: t }])
    setInput('')
    setRunning(true)
    void window.api.ai.chat({ chatId: chatId.current, model: selected, text: t, context: getContext() })
  }

  const reset = () => {
    void window.api.ai.stop(chatId.current)
    void window.api.ai.reset(chatId.current)
    chatId.current = newChatId()
    setMessages([])
    setRunning(false)
  }

  const startResize = (e: React.MouseEvent) => {
    const startX = e.clientX
    const startW = width
    const move = (ev: MouseEvent) => setWidth(Math.max(300, Math.min(700, startW + startX - ev.clientX)))
    const up = (ev: MouseEvent) => {
      window.removeEventListener('mousemove', move)
      window.removeEventListener('mouseup', up)
      try {
        localStorage.setItem('ai-width', String(Math.max(300, Math.min(700, startW + startX - ev.clientX))))
      } catch {
        /* storage unavailable */
      }
    }
    window.addEventListener('mousemove', move)
    window.addEventListener('mouseup', up)
  }

  return (
    <aside className="ai-panel" style={{ width }}>
      <div className="ai-resize" onMouseDown={startResize} />
      <div className="ai-head">
        <span className="title">
          <Sparkles size={15} /> AI
        </span>
        {options.length > 0 ? (
          <select className="model-select" value={selected} onChange={(e) => window.api.settings.update({ selectedModel: e.target.value })} title="Model">
            {options.map((g) => (
              <optgroup key={g.provider.id} label={g.provider.name}>
                {g.models.map((m) => (
                  <option key={m} value={`${g.provider.id}::${m}`}>
                    {m}
                  </option>
                ))}
              </optgroup>
            ))}
          </select>
        ) : (
          <span style={{ flex: 1 }} />
        )}
        <button className="tbtn" title="New chat" onClick={reset}>
          <MessageSquarePlus size={15} />
        </button>
        <button className="tbtn" title="AI settings" onClick={onOpenSettings}>
          <Settings size={15} />
        </button>
        <button className="tbtn" title="Close (⌘J)" onClick={onClose}>
          <X size={15} />
        </button>
      </div>

      <div className="ai-messages" ref={listRef}>
        {messages.length === 0 && (
          <div className="ai-empty">
            <Sparkles size={26} />
            <h3>Ask about your data</h3>
            {options.length === 0 ? (
              <>
                <div>Connect an AI provider to get started. Any provider works: Anthropic, OpenAI, Google, Mistral, Groq, OpenRouter, a local Ollama model and more.</div>
                <button className="btn primary" onClick={onOpenSettings}>
                  Add an API key
                </button>
              </>
            ) : (
              <>
                <div>The assistant can read and edit your sheets, write formulas, and build dashboards.</div>
                <div className="suggestions">
                  {SUGGESTIONS.map((s) => (
                    <button key={s} className="suggestion" onClick={() => send(s)}>
                      {s}
                    </button>
                  ))}
                </div>
              </>
            )}
          </div>
        )}
        {messages.map((m) =>
          m.role === 'user' ? (
            <div key={m.id} className="msg user">
              {m.text}
            </div>
          ) : m.role === 'error' ? (
            <div key={m.id} className="msg error">
              {m.text}
            </div>
          ) : (
            <div key={m.id} className="msg assistant">
              {m.parts?.map((p, i) =>
                p.kind === 'text' ? (
                  <ReactMarkdown key={i} remarkPlugins={[remarkGfm]} components={{ a: ({ href, children }) => <a href={href} target="_blank" rel="noreferrer">{children}</a> }}>
                    {p.text}
                  </ReactMarkdown>
                ) : (
                  <details key={p.id} className="tool-chip" style={{ margin: '6px 0' }}>
                    <summary>
                      {p.error ? <CircleX size={13} className="err" /> : p.output !== undefined ? <CircleCheck size={13} className="ok" /> : <Wrench size={13} />}
                      {TOOL_LABELS[p.name] ?? p.name}
                      {summarize(p.input)}
                    </summary>
                    <pre>{JSON.stringify({ input: p.input, ...(p.error ? { error: p.error } : { output: p.output }) }, null, 2)}</pre>
                  </details>
                )
              )}
            </div>
          )
        )}
        {running && (
          <span className="typing" aria-label="Working">
            <i />
            <i />
            <i />
          </span>
        )}
      </div>

      <div className="ai-input">
        <div className="ai-input-box">
          <textarea
            value={input}
            placeholder={options.length ? 'Ask anything, e.g. "Total sales by region in a new sheet"' : 'Add an API key in Settings to chat'}
            disabled={!options.length}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && !e.shiftKey && !e.nativeEvent.isComposing) {
                e.preventDefault()
                send(input)
              }
            }}
          />
          <div className="ai-input-row">
            <span>Enter to send · Shift+Enter for a new line</span>
            {running ? (
              <button className="tbtn" title="Stop" onClick={() => window.api.ai.stop(chatId.current)}>
                <Square size={14} />
              </button>
            ) : (
              <button className="tbtn primary" title="Send" disabled={!input.trim() || !selected} onClick={() => send(input)}>
                <Send size={14} />
              </button>
            )}
          </div>
        </div>
      </div>
    </aside>
  )
}

function summarize(input: unknown): string {
  if (!input || typeof input !== 'object') return ''
  const o = input as Record<string, unknown>
  const where = o.range ?? o.start ?? o.name ?? o.label ?? o.columns
  return where ? ` · ${o.sheet ? `${o.sheet}!` : ''}${String(where)}` : ''
}
