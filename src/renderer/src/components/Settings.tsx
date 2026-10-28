import { useState } from 'react'
import { Bot, ChevronDown, ChevronRight, ExternalLink, Info, KeyRound, Palette, RefreshCw } from 'lucide-react'
import type { ProviderState, SettingsSnapshot } from '@shared/types'
import { Modal } from './Dialogs'

type Section = 'providers' | 'general' | 'about'

export function SettingsDialog({ snap, initial = 'providers', onClose }: { snap: SettingsSnapshot; initial?: Section; onClose: () => void }) {
  const [section, setSection] = useState<Section>(initial)
  return (
    <Modal title="Settings" wide onClose={onClose}>
      <div className="settings" style={{ margin: -16 }}>
        <nav className="settings-nav">
          <button className={section === 'providers' ? 'on' : ''} onClick={() => setSection('providers')}>
            <KeyRound size={15} /> AI Providers
          </button>
          <button className={section === 'general' ? 'on' : ''} onClick={() => setSection('general')}>
            <Palette size={15} /> General
          </button>
          <button className={section === 'about' ? 'on' : ''} onClick={() => setSection('about')}>
            <Info size={15} /> About
          </button>
        </nav>
        <div className="settings-main">
          {section === 'providers' && <Providers snap={snap} />}
          {section === 'general' && <General snap={snap} />}
          {section === 'about' && <About snap={snap} />}
        </div>
      </div>
    </Modal>
  )
}

function Providers({ snap }: { snap: SettingsSnapshot }) {
  const [open, setOpen] = useState<string | null>(snap.providers.find((p) => !p.hasKey && !p.keyOptional)?.id ?? null)
  const [name, setName] = useState('')
  const [url, setUrl] = useState('')
  const configured = snap.providers.filter((p) => p.enabled && (p.hasKey || p.keyOptional))
  return (
    <>
      <p className="hint" style={{ marginTop: 0 }}>
        Bring your own API keys. Keys are encrypted with your macOS Keychain and stored only on this Mac — Open Sheets sends them to nothing but the provider you
        chose. {configured.length ? `${configured.length} provider${configured.length > 1 ? 's' : ''} ready.` : ''}
      </p>
      {snap.providers.map((p) => (
        <ProviderRow key={p.id} p={p} open={open === p.id} onToggle={() => setOpen(open === p.id ? null : p.id)} />
      ))}
      <div className="provider" style={{ padding: 12 }}>
        <div style={{ fontWeight: 600, marginBottom: 8 }}>Add an OpenAI-compatible provider</div>
        <div className="row" style={{ alignItems: 'flex-end' }}>
          <label className="field">
            <span>Name</span>
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="My vLLM server" />
          </label>
          <label className="field">
            <span>Base URL</span>
            <input value={url} onChange={(e) => setUrl(e.target.value)} placeholder="https://host/v1" spellCheck={false} />
          </label>
          <button
            className="btn"
            disabled={!name.trim() || !/^https?:\/\//.test(url.trim())}
            onClick={async () => {
              const id = await window.api.settings.addCustomProvider(name.trim(), url.trim())
              setName('')
              setUrl('')
              setOpen(id)
            }}
          >
            Add
          </button>
        </div>
      </div>
    </>
  )
}

function ProviderRow({ p, open, onToggle }: { p: ProviderState; open: boolean; onToggle: () => void }) {
  const [key, setKey] = useState('')
  const [base, setBase] = useState(p.baseURLOverride ?? '')
  const [busy, setBusy] = useState(false)
  const [msg, setMsg] = useState<{ ok: boolean; text: string } | null>(null)
  const [q, setQ] = useState('')
  const [custom, setCustom] = useState('')
  const ready = p.enabled && (p.hasKey || !!p.keyOptional)

  const run = async (fn: () => Promise<unknown>, okText?: string) => {
    setBusy(true)
    setMsg(null)
    try {
      await fn()
      if (okText) setMsg({ ok: true, text: okText })
    } catch (err) {
      setMsg({ ok: false, text: (err as Error).message.replace(/^Error invoking remote method '[^']+': (Error: )?/, '') })
    } finally {
      setBusy(false)
    }
  }

  const pinned = new Set(p.pinnedModels)
  const shown = p.models.filter((m) => m.toLowerCase().includes(q.toLowerCase()))
  const togglePin = (m: string) => {
    const next = new Set(pinned)
    if (next.has(m)) next.delete(m)
    else next.add(m)
    void window.api.settings.updateProvider(p.id, { pinnedModels: [...next] })
  }

  return (
    <div className="provider">
      <div className="provider-head" onClick={onToggle}>
        {open ? <ChevronDown size={15} /> : <ChevronRight size={15} />}
        <span className={`dot${ready ? ' on' : ''}`} />
        <span className="name">{p.name}</span>
        {p.hasKey && <span className="badge ok">Key saved</span>}
        {p.keyOptional && !p.hasKey && <span className="badge">No key needed</span>}
        <label className="check" onClick={(e) => e.stopPropagation()} title="Show this provider's models in the chat model picker">
          <input type="checkbox" checked={p.enabled} onChange={(e) => window.api.settings.updateProvider(p.id, { enabled: e.target.checked })} /> Enabled
        </label>
      </div>
      {open && (
        <div className="provider-body">
          {p.note && <div className="hint">{p.note}</div>}
          <div className="row" style={{ alignItems: 'flex-end' }}>
            <label className="field">
              <span>API key {p.hasKey && <span className="muted">— saved ({p.keyHint})</span>}</span>
              <input type="password" value={key} onChange={(e) => setKey(e.target.value)} placeholder={p.hasKey ? 'Enter a new key to replace' : p.keyOptional ? 'Optional' : 'Paste your API key'} autoComplete="off" spellCheck={false} />
            </label>
            <button className="btn primary" disabled={!key.trim() || busy} onClick={() => run(async () => { await window.api.settings.setKey(p.id, key.trim()); setKey('') }, 'Key saved to Keychain.')}>
              Save
            </button>
            {p.hasKey && (
              <button className="btn danger" disabled={busy} onClick={() => run(() => window.api.settings.deleteKey(p.id), 'Key removed.')}>
                Remove
              </button>
            )}
          </div>
          {p.keyUrl && (
            <button className="btn small" style={{ alignSelf: 'flex-start' }} onClick={() => window.api.app.openExternal(p.keyUrl!)}>
              Get an API key <ExternalLink size={12} />
            </button>
          )}
          <label className="field">
            <span>Base URL</span>
            <input
              value={base}
              onChange={(e) => setBase(e.target.value)}
              onBlur={() => base !== (p.baseURLOverride ?? '') && window.api.settings.updateProvider(p.id, { baseURLOverride: base.trim() || undefined })}
              placeholder={p.baseURL ?? 'https://…'}
              spellCheck={false}
            />
          </label>
          <div className="row">
            <strong style={{ fontSize: 12 }}>Models</strong>
            <span className="muted">{p.models.length} available · {pinned.size ? `${pinned.size} shown in picker` : 'all shown in picker'}</span>
            <span style={{ flex: 1 }} />
            <button className="btn small" disabled={busy || (!p.hasKey && !p.keyOptional)} onClick={() => run(async () => { const m = await window.api.settings.fetchModels(p.id); setMsg({ ok: true, text: `Found ${m.length} models.` }) })}>
              <RefreshCw size={12} /> Fetch models
            </button>
          </div>
          {p.models.length > 12 && <input className="tselect" placeholder="Filter models" value={q} onChange={(e) => setQ(e.target.value)} />}
          {p.models.length > 0 && (
            <div className="model-list">
              {shown.map((m) => (
                <label key={m}>
                  <input type="checkbox" checked={pinned.has(m)} onChange={() => togglePin(m)} /> {m}
                </label>
              ))}
            </div>
          )}
          <div className="row">
            <input className="tselect" style={{ flex: 1 }} placeholder="Add a model id manually (e.g. a new release or Azure deployment)" value={custom} onChange={(e) => setCustom(e.target.value)} spellCheck={false} />
            <button
              className="btn small"
              disabled={!custom.trim()}
              onClick={() => {
                const id = custom.trim()
                void window.api.settings.updateProvider(p.id, { models: [...new Set([...p.models, id])], pinnedModels: pinned.size ? [...pinned, id] : [] })
                setCustom('')
              }}
            >
              Add model
            </button>
          </div>
          {msg && <div className={msg.ok ? 'hint' : 'msg error'}>{msg.text}</div>}
          {p.custom && (
            <button className="btn small danger" style={{ alignSelf: 'flex-start' }} onClick={() => window.api.settings.removeCustomProvider(p.id)}>
              Remove provider
            </button>
          )}
        </div>
      )}
    </div>
  )
}

function General({ snap }: { snap: SettingsSnapshot }) {
  const s = snap.settings
  const set = (patch: Partial<typeof s>) => window.api.settings.update(patch)
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 14, maxWidth: 420 }}>
      <label className="field">
        <span>Appearance</span>
        <select value={s.theme} onChange={(e) => set({ theme: e.target.value as typeof s.theme })}>
          <option value="system">Match macOS</option>
          <option value="light">Light</option>
          <option value="dark">Dark</option>
        </select>
      </label>
      <div className="row">
        <label className="field">
          <span>Default font</span>
          <select value={s.defaultFont} onChange={(e) => set({ defaultFont: e.target.value })}>
            {['Helvetica Neue', 'SF Pro Text', 'Arial', 'Avenir Next', 'Calibri', 'Menlo', 'Georgia'].map((f) => (
              <option key={f}>{f}</option>
            ))}
          </select>
        </label>
        <label className="field" style={{ maxWidth: 110 }}>
          <span>Size</span>
          <select value={s.defaultFontSize} onChange={(e) => set({ defaultFontSize: Number(e.target.value) })}>
            {[11, 12, 13, 14, 15, 16].map((n) => (
              <option key={n}>{n}</option>
            ))}
          </select>
        </label>
      </div>
      <label className="check">
        <input type="checkbox" checked={s.showGridlines} onChange={(e) => set({ showGridlines: e.target.checked })} /> Show gridlines in new windows
      </label>
    </div>
  )
}

function About({ snap }: { snap: SettingsSnapshot }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 10, maxWidth: 480 }}>
      <div className="row">
        <Bot size={28} />
        <div>
          <div style={{ fontSize: 16, fontWeight: 650 }}>Open Sheets {snap.version}</div>
          <div className="muted">Free and open-source spreadsheets for macOS</div>
        </div>
      </div>
      <p className="hint">
        Licensed under the GNU General Public License v3.0. Built with HyperFormula (formula engine), ExcelJS and SheetJS (file formats), Apache ECharts (charts) and the
        Vercel AI SDK (AI providers). Open Sheets is an independent project and is not affiliated with or endorsed by Microsoft.
      </p>
      <div className="row">
        <button className="btn small" onClick={() => window.api.app.openExternal('https://github.com/OWNER/open-sheets')}>
          Source code <ExternalLink size={12} />
        </button>
        <button className="btn small" onClick={() => window.api.app.openExternal('https://github.com/OWNER/open-sheets/issues')}>
          Report an issue <ExternalLink size={12} />
        </button>
      </div>
    </div>
  )
}
