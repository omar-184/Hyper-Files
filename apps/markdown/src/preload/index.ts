import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@genoffice/i18n'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { MARKDOWN_CHANNELS } from '../shared/ipc'
import type {
  AutoSaveDefault,
  DocTheme,
  ExportFormat,
  MarkdownApi,
  SaveMode,
  UiTheme,
} from '../shared/ipc'

const api: MarkdownApi = {
  consumePending: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.consumePending),
  consumeHeadlessExport: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.consumeHeadlessExport),
  headlessExportDone: (result: { ok: boolean; error?: string }) =>
    ipcRenderer.send(MARKDOWN_CHANNELS.headlessExportDone, result),
  readFile: (path) => ipcRenderer.invoke(MARKDOWN_CHANNELS.readFile, path),
  save: (request) => ipcRenderer.invoke(MARKDOWN_CHANNELS.save, request),
  setDirty: (dirty) => ipcRenderer.send(MARKDOWN_CHANNELS.dirtyChanged, dirty),
  onSaveRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, mode: SaveMode) => handler(mode)
    ipcRenderer.on(MARKDOWN_CHANNELS.saveRequest, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.saveRequest, listener)
  },
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(MARKDOWN_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(MARKDOWN_CHANNELS.closeSaveResult, ok),
  sendSaveRequestAck: (ok) => ipcRenderer.send(MARKDOWN_CHANNELS.saveRequestAck, ok),
  onFileRenamed: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on(MARKDOWN_CHANNELS.fileRenamed, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.fileRenamed, listener)
  },
  pickImage: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.pickImage),
  saveImage: (data) => ipcRenderer.invoke(MARKDOWN_CHANNELS.saveImage, data),
  readImage: (src) => ipcRenderer.invoke(MARKDOWN_CHANNELS.readImage, src),
  saveImageAs: (src) => ipcRenderer.invoke(MARKDOWN_CHANNELS.saveImageAs, src),
  onViewImage: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, src: string) => handler(src)
    ipcRenderer.on(MARKDOWN_CHANNELS.viewImage, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.viewImage, listener)
  },
  onExportRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, format: ExportFormat) => handler(format)
    ipcRenderer.on(MARKDOWN_CHANNELS.exportRequest, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.exportRequest, listener)
  },
  onPrintRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(MARKDOWN_CHANNELS.printRequest, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.printRequest, listener)
  },
  exportDocx: (request) => ipcRenderer.invoke(MARKDOWN_CHANNELS.exportDocx, request),
  exportPdf: (request) => ipcRenderer.invoke(MARKDOWN_CHANNELS.exportPdf, request),
  prepareImageExport: (request) =>
    ipcRenderer.invoke(MARKDOWN_CHANNELS.prepareImageExport, request),
  writeExportImage: (id, page, base64) =>
    ipcRenderer.invoke(MARKDOWN_CHANNELS.writeExportImage, id, page, base64),
  finishImageExport: (id, success) =>
    ipcRenderer.invoke(MARKDOWN_CHANNELS.finishImageExport, id, success),
  getLanguage: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(MARKDOWN_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.languageChanged, listener)
  },
  getTheme: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.getTheme),
  onThemeChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, theme: UiTheme) => handler(theme)
    ipcRenderer.on(MARKDOWN_CHANNELS.themeChanged, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.themeChanged, listener)
  },
  getDocumentTheme: async () => {
    const result: unknown = await ipcRenderer.invoke(MARKDOWN_CHANNELS.getDocumentTheme)
    return result === 'dark' || result === 'light' ? result : 'follow'
  },
  onDocumentThemeChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, theme: DocTheme) => handler(theme)
    ipcRenderer.on(MARKDOWN_CHANNELS.documentThemeChanged, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.documentThemeChanged, listener)
  },
  getAutoSaveDefault: () => ipcRenderer.invoke(MARKDOWN_CHANNELS.getAutoSaveDefault),
  onAutoSaveDefaultChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, value: AutoSaveDefault) => handler(value)
    ipcRenderer.on(MARKDOWN_CHANNELS.autoSaveDefaultChanged, listener)
    return () => ipcRenderer.removeListener(MARKDOWN_CHANNELS.autoSaveDefaultChanged, listener)
  },
  onChromePressed: (handler) => {
    const listener = () => handler()
    ipcRenderer.on('app:chrome-pressed', listener)
    return () => ipcRenderer.removeListener('app:chrome-pressed', listener)
  },
}

contextBridge.exposeInMainWorld('markdownApi', api)

// open documents dragged from the OS onto this tab as a new shell tab
installDropOpenBridge()
