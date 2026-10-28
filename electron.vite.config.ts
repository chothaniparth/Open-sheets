import { resolve } from 'node:path'
import { defineConfig } from 'electron-vite'
import react from '@vitejs/plugin-react'
import type { Plugin } from 'vite'

// Strict Content-Security-Policy for production builds (dev needs inline scripts for HMR).
const csp = (): Plugin => ({
  name: 'open-sheets-csp',
  apply: 'build',
  transformIndexHtml: (html) =>
    html.replace(
      '<meta charset="UTF-8" />',
      `<meta charset="UTF-8" />\n    <meta http-equiv="Content-Security-Policy" content="default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data:; connect-src 'self'" />`
    )
})

export default defineConfig({
  main: {
    build: { externalizeDeps: true }
  },
  preload: {
    // Sandboxed preload scripts must be CommonJS.
    build: {
      externalizeDeps: true,
      rollupOptions: {
        output: { format: 'cjs', entryFileNames: '[name].cjs' }
      }
    }
  },
  renderer: {
    resolve: {
      alias: { '@shared': resolve('src/shared') }
    },
    plugins: [react(), csp()],
    build: {
      chunkSizeWarningLimit: 4000
    }
  }
})
