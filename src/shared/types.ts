// Types shared between the Electron main process, preload and renderer.

export type ProviderKind =
  | 'anthropic'
  | 'openai'
  | 'google'
  | 'mistral'
  | 'groq'
  | 'xai'
  | 'deepseek'
  | 'cohere'
  | 'azure'
  | 'openai-compatible'

export interface ProviderInfo {
  id: string
  name: string
  kind: ProviderKind
  /** Default API base URL (editable for most providers). */
  baseURL?: string
  /** Local servers such as Ollama / LM Studio need no key. */
  keyOptional?: boolean
  /** Where users create an API key. */
  keyUrl?: string
  /** Shown in settings to explain provider specifics. */
  note?: string
  custom?: boolean
}

export interface ProviderState extends ProviderInfo {
  enabled: boolean
  hasKey: boolean
  /** Masked key preview, e.g. "sk-…4f2a". Never the full key. */
  keyHint?: string
  baseURLOverride?: string
  /** Models fetched from the provider or entered by the user. */
  models: string[]
  /** Subset of models shown in the chat model picker (empty = all). */
  pinnedModels: string[]
  modelsFetchedAt?: number
}

export interface AppSettings {
  theme: 'system' | 'light' | 'dark'
  /** "providerId::modelId" */
  selectedModel?: string
  showGridlines: boolean
  defaultFont: string
  defaultFontSize: number
  aiAutoApply: boolean
}

export interface SettingsSnapshot {
  settings: AppSettings
  providers: ProviderState[]
  version: string
}

export interface OpenedFile {
  path?: string
  name: string
  data: Uint8Array
}

export type ChatEvent =
  | { chatId: string; type: 'text'; text: string }
  | { chatId: string; type: 'reasoning'; text: string }
  | { chatId: string; type: 'tool-call'; id: string; name: string; input: unknown }
  | { chatId: string; type: 'tool-result'; id: string; output: unknown }
  | { chatId: string; type: 'tool-error'; id: string; error: string }
  | { chatId: string; type: 'error'; message: string }
  | { chatId: string; type: 'done'; usage?: { input?: number; output?: number } }

export interface ChatRequest {
  chatId: string
  model: string
  text: string
  /** Live workbook context (active sheet, selection, sheet list). */
  context: string
}

export interface ToolRequest {
  reqId: string
  name: string
  input: unknown
}

export interface ContextMenuItem {
  id?: string
  label?: string
  type?: 'separator'
  enabled?: boolean
  accelerator?: string
}

export const FILE_EXTENSIONS = {
  open: [
    'xlsx', 'xlsm', 'xltx', 'xltm', 'xls', 'xlt', 'xlsb', 'xlw', 'xml', 'ods', 'fods',
    'csv', 'tsv', 'tab', 'txt', 'prn', 'dif', 'slk', 'sylk', 'dbf', 'htm', 'html', 'numbers'
  ],
  save: ['xlsx', 'xls', 'xlsb', 'ods', 'fods', 'csv', 'tsv', 'txt', 'html', 'xml', 'slk', 'dif', 'dbf', 'numbers']
} as const
