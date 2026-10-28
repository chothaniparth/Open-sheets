import { app, BrowserWindow, dialog, ipcMain, Menu, nativeTheme, shell } from 'electron'
import { readFile, writeFile, unlink } from 'node:fs/promises'
import { basename, extname, join } from 'node:path'
import { tmpdir } from 'node:os'
import { randomUUID } from 'node:crypto'
import { FILE_EXTENSIONS, type AppSettings, type ContextMenuItem } from '../shared/types'
import { registerAiIpc } from './ai'
import { buildMenu } from './menu'
import { fetchModelList } from './providers'
import * as store from './store'

const dir = import.meta.dirname

interface DocState {
  path?: string
  edited: boolean
  forceClose: boolean
  ready: boolean
  queued?: string
}

const docs = new Map<BrowserWindow, DocState>()
const pendingOpen: string[] = []

function applyTheme() {
  nativeTheme.themeSource = store.snapshot().settings.theme
}

function createWindow(filePath?: string): BrowserWindow {
  const focused = BrowserWindow.getFocusedWindow()
  const [x, y] = focused ? focused.getPosition() : []
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 760,
    minHeight: 480,
    x: x !== undefined ? x + 24 : undefined,
    y: y !== undefined ? y + 24 : undefined,
    show: false,
    title: 'Untitled',
    titleBarStyle: 'hiddenInset',
    trafficLightPosition: { x: 14, y: 14 },
    backgroundColor: nativeTheme.shouldUseDarkColors ? '#1e1f22' : '#ffffff',
    webPreferences: {
      preload: join(dir, '../preload/index.cjs'),
      contextIsolation: true,
      sandbox: true,
      nodeIntegration: false,
      spellcheck: false
    }
  })
  const state: DocState = { edited: false, forceClose: false, ready: false, queued: filePath }
  docs.set(win, state)

  win.once('ready-to-show', () => win.show())
  win.webContents.on('did-finish-load', () => {
    state.ready = true
    if (state.queued) {
      const p = state.queued
      state.queued = undefined
      void sendFile(win, p)
    }
  })
  // Links in AI answers etc. open in the default browser, never inside the app.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:\/\//.test(url)) void shell.openExternal(url)
    return { action: 'deny' }
  })
  win.webContents.on('will-navigate', (e) => e.preventDefault())

  // Development aids: OPEN_SHEETS_DEBUG=1 forwards renderer console output to the terminal,
  // OPEN_SHEETS_CAPTURE=/path.png saves a screenshot of the first window once it has loaded.
  if (process.env.OPEN_SHEETS_DEBUG) {
    win.webContents.on('console-message', (e) => console.log(`[renderer:${e.level}] ${e.message}`))
  }
  const capture = process.env.OPEN_SHEETS_CAPTURE
  if (capture && docs.size === 1) {
    win.webContents.once('did-finish-load', () => {
      setTimeout(async () => {
        // optional script to put the UI in a given state before capturing
        if (process.env.OPEN_SHEETS_CAPTURE_JS) {
          await win.webContents.executeJavaScript(process.env.OPEN_SHEETS_CAPTURE_JS)
          await new Promise((r) => setTimeout(r, 1500))
        }
        const img = await win.webContents.capturePage()
        await writeFile(capture, img.toPNG())
        console.log(`captured ${capture}`)
      }, Number(process.env.OPEN_SHEETS_CAPTURE_DELAY ?? 2500))
    })
  }

  win.on('close', (e) => {
    if (state.forceClose || !state.edited) return
    e.preventDefault()
    const choice = dialog.showMessageBoxSync(win, {
      type: 'warning',
      message: `Do you want to save the changes you made to "${state.path ? basename(state.path) : 'Untitled'}"?`,
      detail: "Your changes will be lost if you don't save them.",
      buttons: ['Save', 'Cancel', "Don't Save"],
      defaultId: 0,
      cancelId: 1
    })
    if (choice === 0) win.webContents.send('menu:command', 'saveAndClose')
    if (choice === 2) {
      state.forceClose = true
      win.close()
    }
  })
  win.on('closed', () => docs.delete(win))

  if (process.env.ELECTRON_RENDERER_URL) {
    void win.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    void win.loadFile(join(dir, '../renderer/index.html'))
  }
  return win
}

async function sendFile(win: BrowserWindow, path: string) {
  try {
    const data = await readFile(path)
    const state = docs.get(win)
    if (state) state.path = path
    app.addRecentDocument(path)
    win.setRepresentedFilename(path)
    win.webContents.send('file:load', { path, name: basename(path), data: new Uint8Array(data) })
  } catch (err) {
    dialog.showErrorBox('Could not open file', `${path}\n\n${(err as Error).message}`)
  }
}

/** Open in the focused window if it is an untouched new workbook, otherwise in a new window. */
function openPath(path: string) {
  for (const [w, s] of docs) {
    if (s.path === path) {
      w.focus()
      return
    }
  }
  const focused = BrowserWindow.getFocusedWindow()
  const state = focused ? docs.get(focused) : undefined
  if (focused && state && !state.path && !state.edited && state.ready) {
    void sendFile(focused, path)
  } else {
    createWindow(path)
  }
}

async function showOpenDialog() {
  const res = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Spreadsheets', extensions: [...FILE_EXTENSIONS.open] },
      { name: 'All Files', extensions: ['*'] }
    ]
  })
  if (!res.canceled) res.filePaths.forEach(openPath)
}

function broadcastSettings() {
  const snap = store.snapshot()
  for (const w of BrowserWindow.getAllWindows()) w.webContents.send('settings:changed', snap)
}

function registerIpc() {
  ipcMain.handle('file:open', () => showOpenDialog())
  ipcMain.handle('file:new', () => void createWindow())
  ipcMain.handle('file:openPath', (_e, path: string) => {
    if (path && FILE_EXTENSIONS.open.includes(extname(path).slice(1).toLowerCase() as never)) openPath(path)
  })

  ipcMain.handle('file:saveDialog', async (e, opts: { defaultName: string; extensions?: string[] }) => {
    const win = BrowserWindow.fromWebContents(e.sender)!
    const ext = extname(opts.defaultName).slice(1)
    const res = await dialog.showSaveDialog(win, {
      defaultPath: opts.defaultName,
      filters: opts.extensions
        ? [{ name: ext.toUpperCase(), extensions: opts.extensions }]
        : [
            { name: 'Excel Workbook', extensions: ['xlsx'] },
            { name: 'Excel 97-2004 Workbook', extensions: ['xls'] },
            { name: 'Excel Binary Workbook', extensions: ['xlsb'] },
            { name: 'OpenDocument Spreadsheet', extensions: ['ods'] },
            { name: 'CSV (current sheet)', extensions: ['csv'] },
            { name: 'Apple Numbers', extensions: ['numbers'] }
          ]
    })
    return res.canceled ? null : res.filePath
  })

  ipcMain.handle('file:write', async (e, path: string, data: Uint8Array, opts?: { isDocument?: boolean }) => {
    await writeFile(path, data)
    const win = BrowserWindow.fromWebContents(e.sender)
    if (win && opts?.isDocument) {
      const state = docs.get(win)
      if (state) state.path = path
      win.setRepresentedFilename(path)
      app.addRecentDocument(path)
    }
  })

  ipcMain.handle('win:setState', (e, s: { title: string; edited: boolean }) => {
    const win = BrowserWindow.fromWebContents(e.sender)
    const state = win ? docs.get(win) : undefined
    if (!win || !state) return
    state.edited = s.edited
    win.setTitle(s.title)
    win.setDocumentEdited(s.edited)
  })

  ipcMain.handle('win:close', (e) => {
    const win = BrowserWindow.fromWebContents(e.sender)
    const state = win ? docs.get(win) : undefined
    if (win && state) {
      state.forceClose = true
      win.close()
    }
  })

  ipcMain.handle('app:confirm', (e, opts: { message: string; detail?: string; buttons: string[] }) => {
    const win = BrowserWindow.fromWebContents(e.sender)!
    return dialog.showMessageBoxSync(win, {
      type: 'question',
      message: opts.message,
      detail: opts.detail,
      buttons: opts.buttons,
      defaultId: 0,
      cancelId: opts.buttons.length - 1
    })
  })

  ipcMain.handle('app:contextMenu', (e, items: ContextMenuItem[]) => {
    const win = BrowserWindow.fromWebContents(e.sender)!
    return new Promise<string | null>((resolve) => {
      let chosen: string | null = null
      const menu = Menu.buildFromTemplate(
        items.map((it) =>
          it.type === 'separator'
            ? { type: 'separator' as const }
            : { label: it.label, enabled: it.enabled !== false, accelerator: it.accelerator, click: () => (chosen = it.id ?? null) }
        )
      )
      menu.popup({ window: win, callback: () => setTimeout(() => resolve(chosen), 0) })
    })
  })

  ipcMain.handle('app:openExternal', (_e, url: string) => {
    if (/^https:\/\//.test(url)) return shell.openExternal(url)
  })

  // Render sheet HTML in a hidden window to print it or save it as PDF.
  ipcMain.handle('app:printHtml', async (e, html: string, mode: 'print' | 'pdf', defaultName: string) => {
    const parent = BrowserWindow.fromWebContents(e.sender)!
    let target: string | undefined
    if (mode === 'pdf') {
      const res = await dialog.showSaveDialog(parent, { defaultPath: defaultName, filters: [{ name: 'PDF', extensions: ['pdf'] }] })
      if (res.canceled || !res.filePath) return false
      target = res.filePath
    }
    const tmp = join(tmpdir(), `open-sheets-print-${randomUUID()}.html`)
    await writeFile(tmp, html, 'utf8')
    const pw = new BrowserWindow({ show: false, webPreferences: { sandbox: true, javascript: false } })
    try {
      await pw.loadFile(tmp)
      if (mode === 'pdf') {
        const pdf = await pw.webContents.printToPDF({ printBackground: true, landscape: false, pageSize: 'A4' })
        await writeFile(target!, pdf)
        shell.showItemInFolder(target!)
      } else {
        await new Promise<void>((resolve) => pw.webContents.print({ printBackground: true }, () => resolve()))
      }
      return true
    } finally {
      pw.destroy()
      void unlink(tmp).catch(() => {})
    }
  })

  // ---- Settings & API keys ----
  ipcMain.handle('settings:get', () => store.snapshot())
  ipcMain.handle('settings:update', (_e, patch: Partial<AppSettings>) => {
    store.updateSettings(patch)
    if (patch.theme) applyTheme()
    broadcastSettings()
  })
  ipcMain.handle('settings:setKey', (_e, providerId: string, key: string) => {
    store.setApiKey(providerId, key)
    broadcastSettings()
  })
  ipcMain.handle('settings:deleteKey', (_e, providerId: string) => {
    store.deleteApiKey(providerId)
    broadcastSettings()
  })
  ipcMain.handle('settings:updateProvider', (_e, providerId: string, patch: Parameters<typeof store.updateProvider>[1]) => {
    store.updateProvider(providerId, patch)
    broadcastSettings()
  })
  ipcMain.handle('settings:fetchModels', async (_e, providerId: string) => {
    const models = await fetchModelList(store.resolveProvider(providerId))
    store.updateProvider(providerId, { models, modelsFetchedAt: Date.now() })
    broadcastSettings()
    return models
  })
  ipcMain.handle('settings:addCustomProvider', (_e, name: string, baseURL: string) => {
    const id = store.addCustomProvider(name, baseURL)
    broadcastSettings()
    return id
  })
  ipcMain.handle('settings:removeCustomProvider', (_e, providerId: string) => {
    store.removeCustomProvider(providerId)
    broadcastSettings()
  })
}

// macOS delivers Finder double-clicks / dock drops through open-file, possibly before ready.
app.on('open-file', (e, path) => {
  e.preventDefault()
  if (app.isReady()) openPath(path)
  else pendingOpen.push(path)
})

if (!app.requestSingleInstanceLock()) {
  app.quit()
} else {
  app.on('second-instance', (_e, argv) => {
    const files = argv.slice(1).filter((a) => FILE_EXTENSIONS.open.includes(extname(a).slice(1).toLowerCase() as never))
    if (files.length) files.forEach(openPath)
    else createWindow()
  })

  app.whenReady().then(() => {
    app.setName('Open Sheets')
    applyTheme()
    registerIpc()
    registerAiIpc()
    buildMenu({ newWindow: () => createWindow(), open: () => void showOpenDialog() })

    const cliFiles = process.argv
      .slice(app.isPackaged ? 1 : 2)
      .filter((a) => FILE_EXTENSIONS.open.includes(extname(a).slice(1).toLowerCase() as never))
    const initial = [...pendingOpen.splice(0), ...cliFiles]
    if (initial.length) initial.forEach((p) => createWindow(p))
    else createWindow()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) createWindow()
    })
  })

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
}
