import { useEffect, useLayoutEffect, useMemo, useRef, useState, type CSSProperties, type KeyboardEvent, type Ref } from 'react'

interface Props {
  value: string
  onChange: (text: string) => void
  onKeyDown?: (e: KeyboardEvent<HTMLTextAreaElement>) => void
  onFocus?: () => void
  functions: string[]
  className?: string
  style?: CSSProperties
  autoFocus?: boolean
  placeholder?: string
  inputRef?: Ref<HTMLTextAreaElement>
  /** Place the caret at the end when focusing. */
  caretToEnd?: boolean
}

const TOKEN_RE = /(?:^=|[=(,+\-*/^&<>:;\s])([A-Za-z][A-Za-z0-9.]*)$/

/** Textarea used for cell and formula-bar editing, with function-name autocomplete. */
export function FormulaInput({ value, onChange, onKeyDown, onFocus, functions, className, style, autoFocus, placeholder, inputRef, caretToEnd }: Props) {
  const local = useRef<HTMLTextAreaElement | null>(null)
  const [caret, setCaret] = useState(value.length)
  const [sel, setSel] = useState(0)
  const [dismissed, setDismissed] = useState(false)
  const [rect, setRect] = useState<DOMRect | null>(null)

  const setRefs = (el: HTMLTextAreaElement | null) => {
    local.current = el
    if (typeof inputRef === 'function') inputRef(el)
    else if (inputRef) (inputRef as { current: HTMLTextAreaElement | null }).current = el
  }

  useLayoutEffect(() => {
    if (autoFocus && local.current) {
      local.current.focus()
      if (caretToEnd) {
        const n = local.current.value.length
        local.current.setSelectionRange(n, n)
      }
    }
  }, [autoFocus, caretToEnd])

  const token = useMemo(() => {
    if (!value.startsWith('=')) return null
    const m = TOKEN_RE.exec(value.slice(0, caret))
    return m ? m[1] : null
  }, [value, caret])

  const matches = useMemo(() => {
    if (!token || dismissed) return []
    const t = token.toUpperCase()
    return functions.filter((f) => f.startsWith(t) && f !== t).slice(0, 12)
  }, [token, functions, dismissed])

  useEffect(() => {
    setSel(0)
    if (matches.length && local.current) setRect(local.current.getBoundingClientRect())
  }, [matches.length, token])

  const accept = (name: string) => {
    if (!token) return
    const start = caret - token.length
    const next = `${value.slice(0, start)}${name}(${value.slice(caret)}`
    onChange(next)
    const pos = start + name.length + 1
    requestAnimationFrame(() => {
      local.current?.setSelectionRange(pos, pos)
      setCaret(pos)
    })
  }

  const handleKey = (e: KeyboardEvent<HTMLTextAreaElement>) => {
    if (matches.length) {
      if (e.key === 'ArrowDown') {
        e.preventDefault()
        setSel((s) => (s + 1) % matches.length)
        return
      }
      if (e.key === 'ArrowUp') {
        e.preventDefault()
        setSel((s) => (s - 1 + matches.length) % matches.length)
        return
      }
      if (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey)) {
        e.preventDefault()
        accept(matches[sel])
        return
      }
      if (e.key === 'Escape') {
        e.preventDefault()
        setDismissed(true)
        return
      }
    }
    onKeyDown?.(e)
  }

  return (
    <>
      <textarea
        ref={setRefs}
        className={className}
        style={style}
        value={value}
        placeholder={placeholder}
        spellCheck={false}
        rows={1}
        onFocus={onFocus}
        onChange={(e) => {
          setDismissed(false)
          onChange(e.target.value)
          setCaret(e.target.selectionStart ?? e.target.value.length)
        }}
        onSelect={(e) => setCaret((e.target as HTMLTextAreaElement).selectionStart ?? 0)}
        onKeyDown={handleKey}
        onBlur={() => setDismissed(true)}
      />
      {matches.length > 0 && rect && (
        <div className="autocomplete" style={{ left: rect.left, top: rect.bottom + 4 }}>
          {matches.map((m, i) => (
            <div
              key={m}
              className={i === sel ? 'sel' : ''}
              onMouseDown={(e) => {
                e.preventDefault()
                accept(m)
              }}
            >
              {m}
            </div>
          ))}
        </div>
      )}
    </>
  )
}
