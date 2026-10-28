import { app, safeStorage } from 'electron'
import { existsSync, readFileSync, writeFileSync, renameSync } from 'node:fs'
import { join } from 'node:path'
import type { AppSettings, ProviderInfo, ProviderState, SettingsSnapshot } from '../shared/types'
import { BUILTIN_PROVIDERS, DEFAULT_MODELS, type ResolvedProvider } from './providers'

interface ProviderPrefs {
  enabled?: boolean
  baseURLOverride?: string
  models?: string[]
  pinnedModels?: string[]
  modelsFetchedAt?: number
}

interface StoredSettings {
  settings: AppSettings
  providers: Record<string, ProviderPrefs>
  customProviders: ProviderInfo[]
}

const DEFAULT_SETTINGS: AppSettings = {
  theme: 'system',
  showGridlines: true,
  defaultFont: 'Helvetica Neue',
  defaultFontSize: 13,
  aiAutoApply: true
}

const settingsPath = () => join(app.getPath('userData'), 'settings.json')
const secretsPath = () => join(app.getPath('userData'), 'secrets.json')

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf8')) as T
  } catch (err) {
    console.error(`Could not read ${path}`, err)
  }
  return fallback
}

function writeJson(path: string, value: unknown) {
  const tmp = `${path}.tmp`
  writeFileSync(tmp, JSON.stringify(value, null, 2), { mode: 0o600 })
  renameSync(tmp, path)
}

let cache: StoredSettings | undefined

function load(): StoredSettings {
  if (!cache) {
    const raw = readJson<Partial<StoredSettings>>(settingsPath(), {})
    cache = {
      settings: { ...DEFAULT_SETTINGS, ...(raw.settings ?? {}) },
      providers: raw.providers ?? {},
      customProviders: raw.customProviders ?? []
    }
  }
  return cache
}

function save() {
  if (cache) writeJson(settingsPath(), cache)
}

// ---- API keys: encrypted with the macOS Keychain via Electron safeStorage ----

function readSecrets(): Record<string, string> {
  return readJson<Record<string, string>>(secretsPath(), {})
}

export function getApiKey(providerId: string): string | undefined {
  const enc = readSecrets()[providerId]
  if (!enc) return undefined
  try {
    return safeStorage.decryptString(Buffer.from(enc, 'base64'))
  } catch (err) {
    console.error('Could not decrypt key for', providerId, err)
    return undefined
  }
}

export function setApiKey(providerId: string, key: string) {
  if (!safeStorage.isEncryptionAvailable()) {
    throw new Error('Secure storage (Keychain) is not available, so the key was not saved.')
  }
  const secrets = readSecrets()
  secrets[providerId] = safeStorage.encryptString(key.trim()).toString('base64')
  writeJson(secretsPath(), secrets)
  const prefs = (load().providers[providerId] ??= {})
  if (prefs.enabled === undefined) prefs.enabled = true
  save()
}

export function deleteApiKey(providerId: string) {
  const secrets = readSecrets()
  delete secrets[providerId]
  writeJson(secretsPath(), secrets)
}

function keyHint(key: string | undefined) {
  if (!key) return undefined
  return key.length <= 8 ? '••••' : `${key.slice(0, 4)}…${key.slice(-4)}`
}

// ---- Providers ----

function allProviderInfos(): ProviderInfo[] {
  return [...BUILTIN_PROVIDERS, ...load().customProviders]
}

export function resolveProvider(providerId: string): ResolvedProvider {
  const info = allProviderInfos().find((p) => p.id === providerId)
  if (!info) throw new Error(`Unknown AI provider "${providerId}".`)
  const prefs = load().providers[providerId] ?? {}
  return {
    ...info,
    apiKey: getApiKey(providerId),
    effectiveBaseURL: prefs.baseURLOverride || info.baseURL
  }
}

export function snapshot(): SettingsSnapshot {
  const s = load()
  const secrets = readSecrets()
  const providers: ProviderState[] = allProviderInfos().map((info) => {
    const prefs = s.providers[info.id] ?? {}
    const hasKey = Boolean(secrets[info.id])
    const models = prefs.models?.length ? prefs.models : (DEFAULT_MODELS[info.id] ?? [])
    return {
      ...info,
      enabled: prefs.enabled ?? false,
      hasKey,
      keyHint: hasKey ? keyHint(getApiKey(info.id)) : undefined,
      baseURLOverride: prefs.baseURLOverride,
      models,
      pinnedModels: prefs.pinnedModels ?? [],
      modelsFetchedAt: prefs.modelsFetchedAt
    }
  })
  return { settings: s.settings, providers, version: app.getVersion() }
}

export function updateSettings(patch: Partial<AppSettings>) {
  const s = load()
  s.settings = { ...s.settings, ...patch }
  save()
}

export function updateProvider(providerId: string, patch: ProviderPrefs) {
  const s = load()
  s.providers[providerId] = { ...(s.providers[providerId] ?? {}), ...patch }
  save()
}

export function addCustomProvider(name: string, baseURL: string): string {
  const s = load()
  const slug = name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'custom'
  let id = `custom-${slug}`
  let n = 2
  while (allProviderInfos().some((p) => p.id === id)) id = `custom-${slug}-${n++}`
  s.customProviders.push({ id, name, kind: 'openai-compatible', baseURL, custom: true, keyOptional: true })
  s.providers[id] = { enabled: true }
  save()
  return id
}

export function removeCustomProvider(providerId: string) {
  const s = load()
  s.customProviders = s.customProviders.filter((p) => p.id !== providerId)
  delete s.providers[providerId]
  deleteApiKey(providerId)
  save()
}
