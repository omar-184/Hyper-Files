import { contextBridge, ipcRenderer } from 'electron'
import type { Lang } from '@genoffice/i18n'
import { installDropOpenBridge } from '@genoffice/electron-utils/drop-open'
import { HTML_CHANNELS } from '../shared/ipc'
import type { AutoSaveDefault, ExportFormat, HtmlApi, SaveMode, UiTheme } from '../shared/ipc'

const api: HtmlApi = {
  consumePending: () => ipcRenderer.invoke(HTML_CHANNELS.consumePending),
  consumeHeadlessExport: () => ipcRenderer.invoke(HTML_CHANNELS.consumeHeadlessExport),
  headlessExportDone: (result: { ok: boolean; error?: string }) =>
    ipcRenderer.send(HTML_CHANNELS.headlessExportDone, result),
  readFile: (path) => ipcRenderer.invoke(HTML_CHANNELS.readFile, path),
  updatePreview: (text) => ipcRenderer.send(HTML_CHANNELS.previewUpdate, text),
  getPreviewInfo: () => ipcRenderer.invoke(HTML_CHANNELS.previewInfo),
  setPresentFullScreen: (on) => ipcRenderer.invoke(HTML_CHANNELS.presentFullScreen, on),
  presentInNewTab: (title) => ipcRenderer.invoke(HTML_CHANNELS.presentNewTab, title),
  save: (request) => ipcRenderer.invoke(HTML_CHANNELS.save, request),
  setDirty: (dirty) => ipcRenderer.send(HTML_CHANNELS.dirtyChanged, dirty),
  onSaveRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, mode: SaveMode) => handler(mode)
    ipcRenderer.on(HTML_CHANNELS.saveRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.saveRequest, listener)
  },
  onCloseSaveRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(HTML_CHANNELS.closeSaveRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.closeSaveRequest, listener)
  },
  sendCloseSaveResult: (ok) => ipcRenderer.send(HTML_CHANNELS.closeSaveResult, ok),
  sendSaveRequestAck: (ok) => ipcRenderer.send(HTML_CHANNELS.saveRequestAck, ok),
  onFileRenamed: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, newPath: string) => handler(newPath)
    ipcRenderer.on(HTML_CHANNELS.fileRenamed, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.fileRenamed, listener)
  },
  pickImage: () => ipcRenderer.invoke(HTML_CHANNELS.pickImage),
  saveImage: (data) => ipcRenderer.invoke(HTML_CHANNELS.saveImage, data),
  readImage: (src) => ipcRenderer.invoke(HTML_CHANNELS.readImage, src),
  onExportRequest: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, format: ExportFormat) => handler(format)
    ipcRenderer.on(HTML_CHANNELS.exportRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.exportRequest, listener)
  },
  onPrintRequest: (handler) => {
    const listener = () => handler()
    ipcRenderer.on(HTML_CHANNELS.printRequest, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.printRequest, listener)
  },
  exportDocx: (request) => ipcRenderer.invoke(HTML_CHANNELS.exportDocx, request),
  exportPdf: (request) => ipcRenderer.invoke(HTML_CHANNELS.exportPdf, request),
  exportHtml: (request) => ipcRenderer.invoke(HTML_CHANNELS.exportHtml, request),
  printHtml: (request) => ipcRenderer.invoke(HTML_CHANNELS.printHtml, request),
  getLanguage: () => ipcRenderer.invoke(HTML_CHANNELS.getLanguage),
  onLanguageChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, lang: Lang) => handler(lang)
    ipcRenderer.on(HTML_CHANNELS.languageChanged, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.languageChanged, listener)
  },
  getTheme: () => ipcRenderer.invoke(HTML_CHANNELS.getTheme),
  onThemeChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, theme: UiTheme) => handler(theme)
    ipcRenderer.on(HTML_CHANNELS.themeChanged, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.themeChanged, listener)
  },
  getAutoSaveDefault: () => ipcRenderer.invoke(HTML_CHANNELS.getAutoSaveDefault),
  onAutoSaveDefaultChanged: (handler) => {
    const listener = (_e: Electron.IpcRendererEvent, value: AutoSaveDefault) => handler(value)
    ipcRenderer.on(HTML_CHANNELS.autoSaveDefaultChanged, listener)
    return () => ipcRenderer.removeListener(HTML_CHANNELS.autoSaveDefaultChanged, listener)
  },
  onChromePressed: (handler) => {
    const listener = () => handler()
    ipcRenderer.on('app:chrome-pressed', listener)
    return () => ipcRenderer.removeListener('app:chrome-pressed', listener)
  },
  fetchImage: (url) => ipcRenderer.invoke(HTML_CHANNELS.fetchImage, url),
}

contextBridge.exposeInMainWorld('htmlApi', api)

// open documents dragged from the OS onto this tab as a new shell tab
installDropOpenBridge()
