import { contextBridge, ipcRenderer, webUtils, type IpcRendererEvent } from 'electron'
import type {
  AppSettings,
  ChatEvent,
  ChatRequest,
  ContextMenuItem,
  OpenedFile,
  SettingsSnapshot,
  ToolRequest
} from '../shared/types'

function on<T>(channel: string, cb: (payload: T) => void) {
  const listener = (_e: IpcRendererEvent, payload: T) => cb(payload)
  ipcRenderer.on(channel, listener)
  return () => {
    ipcRenderer.removeListener(channel, listener)
  }
}

const api = {
  file: {
    open: () => ipcRenderer.invoke('file:open'),
    newWindow: () => ipcRenderer.invoke('file:new'),
    /** Open a file dropped onto the window (opens in a new window if this one is in use). */
    openDropped: (file: File) => ipcRenderer.invoke('file:openPath', webUtils.getPathForFile(file)),
    saveDialog: (defaultName: string, extensions?: string[]): Promise<string | null> =>
      ipcRenderer.invoke('file:saveDialog', { defaultName, extensions }),
    write: (path: string, data: Uint8Array, opts?: { isDocument?: boolean }): Promise<void> =>
      ipcRenderer.invoke('file:write', path, data, opts),
    onLoad: (cb: (file: OpenedFile) => void) => on('file:load', cb)
  },
  win: {
    setState: (s: { title: string; edited: boolean }) => ipcRenderer.invoke('win:setState', s),
    close: () => ipcRenderer.invoke('win:close')
  },
  app: {
    onMenu: (cb: (command: string) => void) => on('menu:command', cb),
    contextMenu: (items: ContextMenuItem[]): Promise<string | null> => ipcRenderer.invoke('app:contextMenu', items),
    confirm: (message: string, buttons: string[], detail?: string): Promise<number> =>
      ipcRenderer.invoke('app:confirm', { message, buttons, detail }),
    openExternal: (url: string) => ipcRenderer.invoke('app:openExternal', url),
    printHtml: (html: string, mode: 'print' | 'pdf', defaultName: string): Promise<boolean> =>
      ipcRenderer.invoke('app:printHtml', html, mode, defaultName)
  },
  settings: {
    get: (): Promise<SettingsSnapshot> => ipcRenderer.invoke('settings:get'),
    update: (patch: Partial<AppSettings>) => ipcRenderer.invoke('settings:update', patch),
    setKey: (providerId: string, key: string) => ipcRenderer.invoke('settings:setKey', providerId, key),
    deleteKey: (providerId: string) => ipcRenderer.invoke('settings:deleteKey', providerId),
    updateProvider: (
      providerId: string,
      patch: { enabled?: boolean; baseURLOverride?: string; models?: string[]; pinnedModels?: string[] }
    ) => ipcRenderer.invoke('settings:updateProvider', providerId, patch),
    fetchModels: (providerId: string): Promise<string[]> => ipcRenderer.invoke('settings:fetchModels', providerId),
    addCustomProvider: (name: string, baseURL: string): Promise<string> =>
      ipcRenderer.invoke('settings:addCustomProvider', name, baseURL),
    removeCustomProvider: (providerId: string) => ipcRenderer.invoke('settings:removeCustomProvider', providerId),
    onChange: (cb: (s: SettingsSnapshot) => void) => on('settings:changed', cb)
  },
  ai: {
    chat: (req: ChatRequest) => ipcRenderer.invoke('ai:chat', req),
    stop: (chatId: string) => ipcRenderer.invoke('ai:stop', chatId),
    reset: (chatId: string) => ipcRenderer.invoke('ai:reset', chatId),
    onEvent: (cb: (ev: ChatEvent) => void) => on('ai:event', cb),
    onToolRequest: (cb: (req: ToolRequest) => void) => on('ai:tool-request', cb),
    respondTool: (reqId: string, ok: boolean, result?: unknown, error?: string) =>
      ipcRenderer.send('ai:tool-response', { reqId, ok, result, error })
  }
}

export type OpenSheetsApi = typeof api

contextBridge.exposeInMainWorld('api', api)
