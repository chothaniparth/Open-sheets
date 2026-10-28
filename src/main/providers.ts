import { createAnthropic } from '@ai-sdk/anthropic'
import { createAzure } from '@ai-sdk/azure'
import { createCohere } from '@ai-sdk/cohere'
import { createDeepSeek } from '@ai-sdk/deepseek'
import { createGoogleGenerativeAI } from '@ai-sdk/google'
import { createGroq } from '@ai-sdk/groq'
import { createMistral } from '@ai-sdk/mistral'
import { createOpenAI } from '@ai-sdk/openai'
import { createOpenAICompatible } from '@ai-sdk/openai-compatible'
import { createXai } from '@ai-sdk/xai'
import type { ProviderInfo } from '../shared/types'

/**
 * Built-in provider catalog. Any service with an OpenAI-compatible API can also
 * be added by the user as a custom provider from Settings.
 */
export const BUILTIN_PROVIDERS: ProviderInfo[] = [
  { id: 'anthropic', name: 'Anthropic (Claude)', kind: 'anthropic', baseURL: 'https://api.anthropic.com/v1', keyUrl: 'https://console.anthropic.com/settings/keys' },
  { id: 'openai', name: 'OpenAI', kind: 'openai', baseURL: 'https://api.openai.com/v1', keyUrl: 'https://platform.openai.com/api-keys' },
  { id: 'google', name: 'Google Gemini', kind: 'google', baseURL: 'https://generativelanguage.googleapis.com/v1beta', keyUrl: 'https://aistudio.google.com/apikey' },
  { id: 'mistral', name: 'Mistral', kind: 'mistral', baseURL: 'https://api.mistral.ai/v1', keyUrl: 'https://console.mistral.ai/api-keys' },
  { id: 'groq', name: 'Groq', kind: 'groq', baseURL: 'https://api.groq.com/openai/v1', keyUrl: 'https://console.groq.com/keys' },
  { id: 'xai', name: 'xAI (Grok)', kind: 'xai', baseURL: 'https://api.x.ai/v1', keyUrl: 'https://console.x.ai' },
  { id: 'deepseek', name: 'DeepSeek', kind: 'deepseek', baseURL: 'https://api.deepseek.com/v1', keyUrl: 'https://platform.deepseek.com/api_keys' },
  { id: 'cohere', name: 'Cohere', kind: 'cohere', baseURL: 'https://api.cohere.com/v2', keyUrl: 'https://dashboard.cohere.com/api-keys' },
  { id: 'openrouter', name: 'OpenRouter', kind: 'openai-compatible', baseURL: 'https://openrouter.ai/api/v1', keyUrl: 'https://openrouter.ai/keys', note: 'One key for hundreds of models from many vendors.' },
  { id: 'together', name: 'Together AI', kind: 'openai-compatible', baseURL: 'https://api.together.xyz/v1', keyUrl: 'https://api.together.ai/settings/api-keys' },
  { id: 'fireworks', name: 'Fireworks AI', kind: 'openai-compatible', baseURL: 'https://api.fireworks.ai/inference/v1', keyUrl: 'https://fireworks.ai/account/api-keys' },
  { id: 'perplexity', name: 'Perplexity', kind: 'openai-compatible', baseURL: 'https://api.perplexity.ai', keyUrl: 'https://www.perplexity.ai/settings/api' },
  { id: 'azure', name: 'Azure OpenAI', kind: 'azure', keyUrl: 'https://portal.azure.com', note: 'Set Base URL to https://<resource>.openai.azure.com/openai and add your deployment names as models.' },
  { id: 'ollama', name: 'Ollama (local)', kind: 'openai-compatible', baseURL: 'http://localhost:11434/v1', keyOptional: true, note: 'Runs models on your Mac. Install from ollama.com, then click Fetch models.' },
  { id: 'lmstudio', name: 'LM Studio (local)', kind: 'openai-compatible', baseURL: 'http://localhost:1234/v1', keyOptional: true, note: 'Start the LM Studio local server, then click Fetch models.' }
]

/** Starter model ids so the picker is usable before "Fetch models" is clicked. */
export const DEFAULT_MODELS: Record<string, string[]> = {
  anthropic: ['claude-opus-5-5', 'claude-sonnet-5-5', 'claude-haiku-5-5'],
  deepseek: ['deepseek-chat', 'deepseek-reasoner'],
  perplexity: ['sonar', 'sonar-pro']
}

export interface ResolvedProvider extends ProviderInfo {
  apiKey?: string
  effectiveBaseURL?: string
}

export function createModel(p: ResolvedProvider, modelId: string) {
  const apiKey = p.apiKey || (p.keyOptional ? 'not-needed' : undefined)
  const baseURL = p.effectiveBaseURL
  switch (p.kind) {
    case 'anthropic':
      return createAnthropic({ apiKey, baseURL })(modelId)
    case 'openai':
      return createOpenAI({ apiKey, baseURL })(modelId)
    case 'google':
      return createGoogleGenerativeAI({ apiKey, baseURL })(modelId)
    case 'mistral':
      return createMistral({ apiKey, baseURL })(modelId)
    case 'groq':
      return createGroq({ apiKey, baseURL })(modelId)
    case 'xai':
      return createXai({ apiKey, baseURL })(modelId)
    case 'deepseek':
      return createDeepSeek({ apiKey, baseURL })(modelId)
    case 'cohere':
      return createCohere({ apiKey, baseURL })(modelId)
    case 'azure':
      return createAzure({ apiKey, baseURL })(modelId)
    case 'openai-compatible':
      if (!baseURL) throw new Error(`${p.name} needs a Base URL.`)
      return createOpenAICompatible({ name: p.id, apiKey, baseURL })(modelId)
  }
}

/** Ask the provider which models the key can use. */
export async function fetchModelList(p: ResolvedProvider): Promise<string[]> {
  const base = (p.effectiveBaseURL || '').replace(/\/+$/, '')
  const get = async (url: string, headers: Record<string, string>) => {
    const res = await fetch(url, { headers, signal: AbortSignal.timeout(15000) })
    if (!res.ok) {
      const body = await res.text().catch(() => '')
      throw new Error(`HTTP ${res.status}: ${body.slice(0, 300)}`)
    }
    return res.json() as Promise<any>
  }
  let ids: string[] = []
  switch (p.kind) {
    case 'anthropic': {
      const j = await get(`${base}/models?limit=1000`, {
        'x-api-key': p.apiKey ?? '',
        'anthropic-version': '2023-06-01'
      })
      ids = (j.data ?? []).map((m: any) => m.id)
      break
    }
    case 'google': {
      const j = await get(`${base}/models?pageSize=1000`, { 'x-goog-api-key': p.apiKey ?? '' })
      ids = (j.models ?? [])
        .filter((m: any) => (m.supportedGenerationMethods ?? []).includes('generateContent'))
        .map((m: any) => String(m.name).replace(/^models\//, ''))
      break
    }
    case 'cohere': {
      const j = await get('https://api.cohere.com/v1/models?endpoint=chat&page_size=1000', {
        Authorization: `Bearer ${p.apiKey}`
      })
      ids = (j.models ?? []).map((m: any) => m.name)
      break
    }
    case 'azure':
      throw new Error('Azure does not list deployments over this API. Add your deployment names manually.')
    default: {
      const headers: Record<string, string> = {}
      if (p.apiKey) headers.Authorization = `Bearer ${p.apiKey}`
      const j = await get(`${base}/models`, headers)
      const list = Array.isArray(j) ? j : (j.data ?? j.models ?? [])
      ids = list.map((m: any) => m.id ?? m.name).filter(Boolean)
    }
  }
  return [...new Set(ids)].sort((a, b) => a.localeCompare(b))
}
