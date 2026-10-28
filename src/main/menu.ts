import { app, BrowserWindow, Menu, shell, type MenuItemConstructorOptions } from 'electron'

type OpenHandler = () => void

/** Sends a command id to the focused window's renderer. */
function cmd(id: string, label: string, accelerator?: string): MenuItemConstructorOptions {
  return {
    label,
    accelerator,
    click: (_item, win) => {
      const target = (win as BrowserWindow | undefined) ?? BrowserWindow.getFocusedWindow()
      target?.webContents.send('menu:command', id)
    }
  }
}

const sep: MenuItemConstructorOptions = { type: 'separator' }

export function buildMenu(handlers: { newWindow: OpenHandler; open: OpenHandler }) {
  const isDev = !app.isPackaged
  const template: MenuItemConstructorOptions[] = [
    {
      label: app.name,
      submenu: [
        { role: 'about' },
        sep,
        cmd('settings', 'Settings…', 'Cmd+,'),
        sep,
        { role: 'services' },
        sep,
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        sep,
        { role: 'quit' }
      ]
    },
    {
      label: 'File',
      submenu: [
        { label: 'New Workbook', accelerator: 'Cmd+N', click: handlers.newWindow },
        { label: 'Open…', accelerator: 'Cmd+O', click: handlers.open },
        { role: 'recentDocuments', submenu: [{ role: 'clearRecentDocuments' }] },
        sep,
        cmd('save', 'Save', 'Cmd+S'),
        cmd('saveAs', 'Save As…', 'Shift+Cmd+S'),
        {
          label: 'Export',
          submenu: [
            cmd('export:xlsx', 'Excel Workbook (.xlsx)…'),
            cmd('export:xls', 'Excel 97-2004 (.xls)…'),
            cmd('export:ods', 'OpenDocument (.ods)…'),
            cmd('export:csv', 'CSV — current sheet (.csv)…'),
            cmd('export:tsv', 'Tab-separated (.tsv)…'),
            cmd('export:html', 'Web page (.html)…'),
            cmd('export:numbers', 'Apple Numbers (.numbers)…'),
            sep,
            cmd('export:pdf', 'PDF — current sheet…')
          ]
        },
        sep,
        cmd('print', 'Print…', 'Cmd+P'),
        sep,
        { role: 'close' }
      ]
    },
    {
      label: 'Edit',
      submenu: [
        cmd('undo', 'Undo', 'Cmd+Z'),
        cmd('redo', 'Redo', 'Shift+Cmd+Z'),
        sep,
        { role: 'cut' },
        { role: 'copy' },
        { role: 'paste' },
        cmd('pasteValues', 'Paste Values', 'Shift+Cmd+V'),
        sep,
        cmd('fillDown', 'Fill Down', 'Cmd+D'),
        cmd('fillRight', 'Fill Right', 'Cmd+R'),
        cmd('clear', 'Clear Contents'),
        cmd('clearFormats', 'Clear Formats'),
        sep,
        cmd('selectAll', 'Select All', 'Cmd+A'),
        cmd('find', 'Find & Replace…', 'Cmd+F')
      ]
    },
    {
      label: 'Insert',
      submenu: [
        cmd('insert:rowAbove', 'Row Above'),
        cmd('insert:rowBelow', 'Row Below'),
        cmd('insert:colLeft', 'Column Left'),
        cmd('insert:colRight', 'Column Right'),
        sep,
        cmd('insert:sheet', 'New Sheet', 'Shift+F11'),
        cmd('insert:dashboard', 'New Dashboard'),
        sep,
        cmd('insert:chart', 'Chart…', 'Alt+F1'),
        cmd('insert:kpi', 'KPI Card…'),
        cmd('insert:function', 'Function…', 'Shift+F3')
      ]
    },
    {
      label: 'Format',
      submenu: [
        cmd('format:bold', 'Bold', 'Cmd+B'),
        cmd('format:italic', 'Italic', 'Cmd+I'),
        cmd('format:underline', 'Underline', 'Cmd+U'),
        sep,
        cmd('format:merge', 'Merge & Center'),
        cmd('format:wrap', 'Wrap Text'),
        sep,
        cmd('format:conditional', 'Conditional Formatting…')
      ]
    },
    {
      label: 'Data',
      submenu: [
        cmd('data:sortAsc', 'Sort A → Z'),
        cmd('data:sortDesc', 'Sort Z → A'),
        cmd('data:filter', 'Filter', 'Shift+Cmd+L'),
        sep,
        cmd('data:removeDuplicates', 'Remove Duplicates')
      ]
    },
    {
      label: 'View',
      submenu: [
        cmd('view:ai', 'Show/Hide AI Assistant', 'Cmd+J'),
        cmd('view:formulas', 'Show Formulas', 'Ctrl+`'),
        cmd('view:gridlines', 'Show/Hide Gridlines'),
        sep,
        cmd('view:freezeTop', 'Freeze Top Row'),
        cmd('view:freezeFirst', 'Freeze First Column'),
        cmd('view:freezeSelection', 'Freeze Panes at Selection'),
        cmd('view:unfreeze', 'Unfreeze Panes'),
        sep,
        { role: 'resetZoom' },
        { role: 'zoomIn' },
        { role: 'zoomOut' },
        sep,
        { role: 'togglefullscreen' },
        ...(isDev ? [sep, { role: 'reload' } as const, { role: 'toggleDevTools' } as const] : [])
      ]
    },
    { role: 'windowMenu' },
    {
      role: 'help',
      submenu: [
        cmd('help:functions', 'Function Reference'),
        { label: 'Open Sheets on GitHub', click: () => shell.openExternal('https://github.com/chothaniparth/Open-sheets') },
        { label: 'Report an Issue', click: () => shell.openExternal('https://github.com/chothaniparth/Open-sheets/issues') }
      ]
    }
  ]
  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}
