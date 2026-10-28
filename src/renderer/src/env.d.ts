/// <reference types="vite/client" />
import type { OpenSheetsApi } from '../../preload/index'

declare global {
  interface Window {
    api: OpenSheetsApi
  }
}

declare module 'xlsx/dist/cpexcel.full.mjs' {
  export const utils: unknown
}
