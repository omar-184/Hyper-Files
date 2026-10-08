import { createHash, randomUUID } from 'node:crypto'
import {
  createReadStream,
  existsSync,
  mkdirSync,
  readFileSync,
  renameSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from 'node:fs'
import { copyFile, mkdir, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { createServer } from 'node:http'
import { basename, dirname, isAbsolute, join } from 'node:path'

import {
  app,
  BrowserWindow,
  desktopCapturer,
  dialog,
  ipcMain,
  Menu,
  screen,
  shell,
  systemPreferences,
  WebContentsView,
} from 'electron'
import type {
  IpcMainInvokeEvent,
  MenuItemConstructorOptions,
  OpenDialogOptions,
  SaveDialogOptions,
  WebContents,
} from 'electron'
import { z } from 'zod'
import {
  appMenuLabels,
  configuredDefaultSaveDir,
  contextMenuLabels,
  installContextMenu,
  installNavigationGuard,
  isHeadlessMode,
  safeExternalUrl,
  showOpenDialogWithMemory,
  showSaveDialogWithMemory,
  helpMenuTemplate,
  viewMenuTemplate,
  windowMenuTemplate,
  installRendererProtocol,
  registerRendererScheme,
  rendererUrl,
} from '@genoffice/electron-utils'
import { createI18n, getUiLang, type Lang, normalizeLang, setUiLang } from '@genoffice/i18n'
import { csvToXlsxBufferForOpen, decodeCsvBuffer } from '@genoffice/xlsx-gateway/gateway/csv-import'
import type { CellEdit, SheetStructuralOps } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import {
  readArchiveEntryText,
  saveWorkbookViaSidecar,
} from '@genoffice/xlsx-gateway/gateway/xlsx-package-io'
import { parsePivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import type { SheetEditPlan } from '@genoffice/xlsx-gateway/gateway/xlsx-sheets'
import type { WorkbookFile } from '../shared/desktop-api'
import {
  workbookFileSchema,
  workbookFormulaCellsRequestSchema,
  workbookFormulaCellsResultSchema,
  workbookRecalcRequestSchema,
  workbookRecalcResultSchema,
  workbookMediaRequestSchema,
  workbookMediaResultSchema,
  workbookPivotRequestSchema,
  localImageRequestSchema,
  localImageResultSchema,
  screenCaptureRequestSchema,
  screenCaptureResultSchema,
  screenSourcesResultSchema,
  workbookPivotDefinitionSchema,
  workbookExportCsvRequestSchema,
  workbookExportPdfRequestSchema,
  workbookRangeRequestSchema,
  workbookRangeResultSchema,
  workbookSaveEditsAbortSchema,
  workbookSaveEditsBeginSchema,
  saveEditsChunkArraySchema,
  workbookSaveEditsChunkSchema,
  workbookSaveRequestSchema,
  type WorkbookSaveRequest,
} from '../shared/desktop-api'
import { IPC_CHANNELS } from '../shared/ipc-channels'
import { atomicWriteFile } from './atomic-write'
import { closeGuardDecision, ShutdownLatch } from './close-guard'
import { SaveEditsTransferStore } from './save-edits-transfer'
import { exportPdf, printWorkbook } from './pdf-export'
import { allowsAutomaticWorkbookRecovery } from './recovery-policy'
import {
  setSystemShortDate,
  shortDatePatternForSystemLocale,
} from '@genoffice/xlsx-gateway/shared/short-date'
import { cleanupImportTempDirectory, cleanupSessionResources } from './temp-files'
import { XlsxSidecarClient } from './xlsx-sidecar-client'
import { sessionAfterRename } from './session-rename'

/**
 * Sheets main-process logic as an embeddable module: no top-level lifecycle.
 * Standalone mode (apps/sheets entry) calls startSheetsStandalone(); the
 * unified shell calls configureSheetsRuntime() + createSheetsWindow() and
 * owns the app lifecycle.
 */

const tMain = createI18n({
  zh: {
    filterSpreadsheets: '电子表格',
    filterXlsx: 'Excel 工作簿',
    filterXlsm: 'Excel 启用宏的工作簿',
    errFileTooLarge: '文件超过大小上限',
    errImgAbsPath: '图片路径必须是绝对路径。',
    errImgNotFound: '找不到图片文件: {path}',
    errImgTooLarge20: '图片超过 20MB,不支持插入。',
    errImgBadType: '该文件不是 PNG/JPEG/GIF 图片。',
    errDiskChanged: '工作簿在打开后被磁盘上的改动覆盖——请改用另存为。',
    autosaveFoundTitle: '发现自动恢复版本',
    autosaveFoundBody:
      '上次会话有未保存的更改。要恢复自动保存的版本吗?恢复后,保存将直接覆盖原文件。',
    autosaveRestore: '恢复',
    autosaveDiscard: '放弃',
    menuFile: '文件',
    menuOpenWorkbook: '打开工作簿…',
    menuSave: '保存',
    menuSaveAs: '另存为…',
    menuExportPdf: '导出 PDF…',
    menuPrint: '打印…',
    menuClose: '关闭',
    menuQuit: '退出',
    menuEdit: '编辑',
    menuUndo: '撤销',
    menuRedo: '重做',
    closeUnsavedMsg: '有 {count} 处未保存的修改',
    closeUnsavedDetail: '不保存直接关闭,这些修改将丢失。',
    btnDontSave: '不保存',
    btnCancel: '取消',
    csvSaveAsNotice: 'CSV 格式不保留样式等格式修改——另存为 .xlsx 可保留全部内容。',
    menuExportCsv: '导出 CSV…',
    filterCsv: 'CSV (逗号分隔)',
    csvFormulaLossMsg: '当前工作表包含公式,CSV 格式无法保留。',
    csvFormulaLossDetail: 'CSV 只保留纯文本值——公式会被替换为当前计算结果,格式也会丢失。',
    csvKeepXlsxBtn: '另存为 .xlsx',
    csvContinueBtn: '继续保存为 CSV',
    csvActiveSheetOnlyNotice: 'CSV 文件只包含一张工作表——只会导出当前工作表“{name}”。',
    csvKeepFormatMsg: '继续以 CSV 格式保存吗?',
    csvKeepFormatDetail:
      'CSV 只保留单张工作表的纯文本值——公式、格式和其他工作表不会存入 .csv 文件。',
  },
  en: {
    filterSpreadsheets: 'Spreadsheets',
    filterXlsx: 'Excel Workbooks',
    filterXlsm: 'Excel Macro-Enabled Workbooks',
    errFileTooLarge: 'File exceeds the size limit',
    errImgAbsPath: 'Image path must be absolute.',
    errImgNotFound: 'Image file not found: {path}',
    errImgTooLarge20: 'Image exceeds 20MB and cannot be inserted.',
    errImgBadType: 'The file is not a PNG/JPEG/GIF image.',
    errDiskChanged: 'The workbook changed on disk after it was opened — use Save As instead.',
    autosaveFoundTitle: 'Recovered version found',
    autosaveFoundBody:
      'There are unsaved changes from your last session. Restore the autosaved version? Saving after a restore overwrites the original file.',
    autosaveRestore: 'Restore',
    autosaveDiscard: 'Discard',
    menuFile: 'File',
    menuOpenWorkbook: 'Open Workbook…',
    menuSave: 'Save',
    menuSaveAs: 'Save As…',
    menuExportPdf: 'Export PDF…',
    menuPrint: 'Print…',
    menuClose: 'Close',
    menuQuit: 'Quit',
    menuEdit: 'Edit',
    menuUndo: 'Undo',
    menuRedo: 'Redo',
    closeUnsavedMsg: '{count} unsaved change(s)',
    closeUnsavedDetail: 'Your changes will be lost if you close without saving.',
    btnDontSave: "Don't Save",
    btnCancel: 'Cancel',
    csvSaveAsNotice: "CSV files can't keep formatting — saving as .xlsx keeps all your changes.",
    menuExportCsv: 'Export CSV…',
    filterCsv: 'CSV (Comma delimited)',
    csvFormulaLossMsg: 'This sheet contains formulas that CSV cannot keep.',
    csvFormulaLossDetail:
      'CSV keeps plain values only — formulas are flattened to their current results, and formatting is lost.',
    csvKeepXlsxBtn: 'Save as .xlsx',
    csvContinueBtn: 'Continue as CSV',
    csvActiveSheetOnlyNotice:
      'CSV files hold a single sheet — only the active sheet "{name}" will be exported.',
    csvKeepFormatMsg: 'Keep saving in CSV format?',
    csvKeepFormatDetail:
      'CSV keeps plain values of a single sheet only — formulas, formatting, and any additional sheets are not saved to the .csv file.',
  },
  vi: {
    filterSpreadsheets: 'Bảng tính',
    filterXlsx: 'Sổ làm việc Excel',
    filterXlsm: 'Sổ làm việc Excel hỗ trợ Macro',
    errFileTooLarge: 'Tệp vượt quá giới hạn kích thước',
    errImgAbsPath: 'Đường dẫn hình ảnh phải là đường dẫn tuyệt đối.',
    errImgNotFound: 'Không tìm thấy tệp hình ảnh: {path}',
    errImgTooLarge20: 'Hình ảnh vượt quá 20MB và không thể chèn.',
    errImgBadType: 'Tệp không phải là hình ảnh PNG/JPEG/GIF.',
    errDiskChanged: 'Sổ làm việc đã thay đổi trên đĩa sau khi mở — hãy sử dụng Lưu dưới dạng.',
    autosaveFoundTitle: 'Tìm thấy phiên bản khôi phục',
    autosaveFoundBody:
      'Có những thay đổi chưa được lưu từ phiên làm việc trước của bạn. Khôi phục phiên bản đã lưu tự động? Việc lưu sau khi khôi phục sẽ ghi đè lên tệp gốc.',
    autosaveRestore: 'Khôi phục',
    autosaveDiscard: 'Bỏ qua',
    menuFile: 'Tệp',
    menuOpenWorkbook: 'Mở sổ làm việc…',
    menuSave: 'Lưu',
    menuSaveAs: 'Lưu dưới dạng…',
    menuExportPdf: 'Xuất PDF…',
    menuPrint: 'In…',
    menuClose: 'Đóng',
    menuQuit: 'Thoát',
    menuEdit: 'Chỉnh sửa',
    menuUndo: 'Hoàn tác',
    menuRedo: 'Làm lại',
    closeUnsavedMsg: '{count} thay đổi chưa lưu',
    closeUnsavedDetail: 'Các thay đổi của bạn sẽ bị mất nếu bạn đóng mà không lưu.',
    btnDontSave: 'Không lưu',
    btnCancel: 'Hủy',
    csvSaveAsNotice:
      'Tệp CSV không thể giữ lại định dạng — lưu dưới dạng .xlsx sẽ giữ lại tất cả các thay đổi của bạn.',
    menuExportCsv: 'Xuất CSV…',
    filterCsv: 'CSV (Phân tách bằng dấu phẩy)',
    csvFormulaLossMsg: 'Trang tính này chứa các công thức mà định dạng CSV không thể lưu giữ.',
    csvFormulaLossDetail:
      'CSV chỉ giữ các giá trị thuần — các công thức được làm phẳng thành kết quả hiện tại của chúng và định dạng sẽ bị mất.',
    csvKeepXlsxBtn: 'Lưu dưới dạng .xlsx',
    csvContinueBtn: 'Tiếp tục dưới dạng CSV',
    csvActiveSheetOnlyNotice:
      'Tệp CSV chỉ chứa một trang tính duy nhất — chỉ trang tính đang hoạt động "{name}" sẽ được xuất.',
    csvKeepFormatMsg: 'Tiếp tục lưu ở định dạng CSV?',
    csvKeepFormatDetail:
      'CSV chỉ giữ các giá trị thuần của một trang tính duy nhất — các công thức, định dạng và mọi trang tính bổ sung sẽ không được lưu vào tệp .csv.',
  },
  ja: {
    filterSpreadsheets: 'スプレッドシート',
    filterXlsx: 'Excel ブック',
    filterXlsm: 'Excel マクロ有効ブック',
    errFileTooLarge: 'ファイルがサイズ上限を超えています',
    errImgAbsPath: '画像パスは絶対パスで指定してください。',
    errImgNotFound: '画像ファイルが見つかりません: {path}',
    errImgTooLarge20: '画像が 20MB を超えているため挿入できません。',
    errImgBadType: 'このファイルは PNG/JPEG/GIF 画像ではありません。',
    errDiskChanged:
      'ブックを開いた後にディスク上で変更されています — 名前を付けて保存を使用してください。',
    autosaveFoundTitle: '自動回復バージョンがあります',
    autosaveFoundBody:
      '前回のセッションに未保存の変更があります。自動保存版を復元しますか?復元後に保存すると、元のファイルは上書きされます。',
    autosaveRestore: '復元',
    autosaveDiscard: '破棄',
    menuFile: 'ファイル',
    menuOpenWorkbook: 'ブックを開く…',
    menuSave: '保存',
    menuSaveAs: '名前を付けて保存…',
    menuExportPdf: 'PDF をエクスポート…',
    menuPrint: '印刷…',
    menuClose: '閉じる',
    menuQuit: '終了',
    menuEdit: '編集',
    menuUndo: '元に戻す',
    menuRedo: 'やり直し',
    closeUnsavedMsg: '未保存の変更が {count} 件あります',
    closeUnsavedDetail: '保存せずに閉じると、これらの変更は失われます。',
    btnDontSave: '保存しない',
    btnCancel: 'キャンセル',
    csvSaveAsNotice:
      'CSV 形式は書式を保持できません。.xlsx として保存すると変更をすべて保持できます。',
    menuExportCsv: 'CSV をエクスポート…',
    filterCsv: 'CSV (コンマ区切り)',
    csvFormulaLossMsg: 'このシートには CSV 形式では保持できない数式が含まれています。',
    csvFormulaLossDetail:
      'CSV は値のみを保持します。数式は現在の計算結果に置き換えられ、書式も失われます。',
    csvKeepXlsxBtn: '.xlsx として保存',
    csvContinueBtn: 'CSV のまま保存',
    csvActiveSheetOnlyNotice:
      'CSV ファイルには 1 枚のシートしか含められません — アクティブなシート「{name}」のみがエクスポートされます。',
    csvKeepFormatMsg: 'CSV 形式のまま保存しますか?',
    csvKeepFormatDetail:
      'CSV は 1 枚のシートの値のみを保持します。数式、書式、追加のシートは .csv ファイルには保存されません。',
  },
  ko: {
    filterSpreadsheets: '스프레드시트',
    filterXlsx: 'Excel 통합 문서',
    filterXlsm: 'Excel 매크로 사용 통합 문서',
    errFileTooLarge: '파일이 크기 제한을 초과했습니다',
    errImgAbsPath: '이미지 경로는 절대 경로여야 합니다.',
    errImgNotFound: '이미지 파일을 찾을 수 없습니다: {path}',
    errImgTooLarge20: '이미지가 20MB를 초과하여 삽입할 수 없습니다.',
    errImgBadType: '이 파일은 PNG/JPEG/GIF 이미지가 아닙니다.',
    errDiskChanged:
      '통합 문서가 열린 후 디스크에서 변경되었습니다. 다른 이름으로 저장을 사용하세요.',
    autosaveFoundTitle: '자동 복구 버전 발견',
    autosaveFoundBody:
      '마지막 세션에 저장되지 않은 변경 내용이 있습니다. 자동 저장 버전을 복원할까요? 복원 후 저장하면 원본 파일을 덮어씁니다.',
    autosaveRestore: '복원',
    autosaveDiscard: '취소',
    menuFile: '파일',
    menuOpenWorkbook: '통합 문서 열기…',
    menuSave: '저장',
    menuSaveAs: '다른 이름으로 저장…',
    menuExportPdf: 'PDF 내보내기…',
    menuPrint: '인쇄…',
    menuClose: '닫기',
    menuQuit: '끝내기',
    menuEdit: '편집',
    menuUndo: '실행 취소',
    menuRedo: '다시 실행',
    closeUnsavedMsg: '저장하지 않은 변경이 {count}건 있습니다',
    closeUnsavedDetail: '저장하지 않고 닫으면 변경 내용이 손실됩니다.',
    btnDontSave: '저장 안 함',
    btnCancel: '취소',
    csvSaveAsNotice:
      'CSV 형식은 서식을 저장할 수 없습니다. .xlsx로 저장하면 모든 변경 내용이 유지됩니다.',
    menuExportCsv: 'CSV 내보내기…',
    filterCsv: 'CSV (쉼표로 분리)',
    csvFormulaLossMsg: '현재 시트에 CSV 형식이 유지할 수 없는 수식이 포함되어 있습니다.',
    csvFormulaLossDetail:
      'CSV는 값만 유지합니다 — 수식은 현재 계산 결과로 바뀌고 서식은 손실됩니다.',
    csvKeepXlsxBtn: '.xlsx로 저장',
    csvContinueBtn: 'CSV로 계속 저장',
    csvActiveSheetOnlyNotice:
      'CSV 파일에는 시트 하나만 포함됩니다 — 활성 시트 "{name}"만 내보냅니다.',
    csvKeepFormatMsg: 'CSV 형식으로 계속 저장하시겠습니까?',
    csvKeepFormatDetail:
      'CSV는 시트 하나의 값만 유지합니다 — 수식, 서식, 추가 시트는 .csv 파일에 저장되지 않습니다.',
  },
  fr: {
    filterSpreadsheets: 'Feuilles de calcul',
    filterXlsx: 'Classeurs Excel',
    filterXlsm: 'Classeurs Excel prenant en charge les macros',
    errFileTooLarge: 'Le fichier dépasse la taille limite',
    errImgAbsPath: "Le chemin de l'image doit être absolu.",
    errImgNotFound: 'Fichier image introuvable : {path}',
    errImgTooLarge20: "L'image dépasse 20 Mo et ne peut pas être insérée.",
    errImgBadType: "Ce fichier n'est pas une image PNG/JPEG/GIF.",
    errDiskChanged:
      'Le classeur a été modifié sur le disque après son ouverture — utilisez Enregistrer sous.',
    autosaveFoundTitle: 'Version récupérée trouvée',
    autosaveFoundBody:
      "Des modifications non enregistrées existent. Restaurer la version auto-enregistrée ? Après restauration, l'enregistrement remplacera le fichier d'origine.",
    autosaveRestore: 'Restaurer',
    autosaveDiscard: 'Ignorer',
    menuFile: 'Fichier',
    menuOpenWorkbook: 'Ouvrir un classeur…',
    menuSave: 'Enregistrer',
    menuSaveAs: 'Enregistrer sous…',
    menuExportPdf: 'Exporter en PDF…',
    menuPrint: 'Imprimer…',
    menuClose: 'Fermer',
    menuQuit: 'Quitter',
    menuEdit: 'Édition',
    menuUndo: 'Annuler',
    menuRedo: 'Rétablir',
    closeUnsavedMsg: '{count} modification(s) non enregistrée(s)',
    closeUnsavedDetail: 'Vos modifications seront perdues si vous fermez sans enregistrer.',
    btnDontSave: 'Ne pas enregistrer',
    btnCancel: 'Annuler',
    csvSaveAsNotice:
      'Le format CSV ne conserve pas la mise en forme — enregistrez en .xlsx pour conserver toutes vos modifications.',
    menuExportCsv: 'Exporter en CSV…',
    filterCsv: 'CSV (délimité par des virgules)',
    csvFormulaLossMsg:
      'Cette feuille contient des formules que le format CSV ne peut pas conserver.',
    csvFormulaLossDetail:
      'Le CSV ne conserve que les valeurs — les formules sont remplacées par leur résultat actuel et la mise en forme est perdue.',
    csvKeepXlsxBtn: 'Enregistrer en .xlsx',
    csvContinueBtn: 'Continuer en CSV',
    csvActiveSheetOnlyNotice:
      "Les fichiers CSV ne contiennent qu'une seule feuille — seule la feuille active « {name} » sera exportée.",
    csvKeepFormatMsg: 'Continuer à enregistrer au format CSV ?',
    csvKeepFormatDetail:
      "Le CSV ne conserve que les valeurs d'une seule feuille — les formules, la mise en forme et les feuilles supplémentaires ne sont pas enregistrées dans le fichier .csv.",
  },
  de: {
    filterSpreadsheets: 'Tabellenkalkulationen',
    filterXlsx: 'Excel-Arbeitsmappen',
    filterXlsm: 'Excel-Arbeitsmappen mit Makros',
    errFileTooLarge: 'Datei überschreitet die Größenbeschränkung',
    errImgAbsPath: 'Der Bildpfad muss absolut sein.',
    errImgNotFound: 'Bilddatei nicht gefunden: {path}',
    errImgTooLarge20: 'Das Bild überschreitet 20 MB und kann nicht eingefügt werden.',
    errImgBadType: 'Die Datei ist kein PNG/JPEG/GIF-Bild.',
    errDiskChanged:
      'Die Arbeitsmappe wurde nach dem Öffnen auf dem Datenträger geändert — verwenden Sie stattdessen „Speichern unter“.',
    autosaveFoundTitle: 'Wiederhergestellte Version gefunden',
    autosaveFoundBody:
      'Es gibt ungespeicherte Änderungen. Automatisch gespeicherte Version wiederherstellen? Nach der Wiederherstellung überschreibt Speichern die Originaldatei.',
    autosaveRestore: 'Wiederherstellen',
    autosaveDiscard: 'Verwerfen',
    menuFile: 'Datei',
    menuOpenWorkbook: 'Arbeitsmappe öffnen…',
    menuSave: 'Speichern',
    menuSaveAs: 'Speichern unter…',
    menuExportPdf: 'PDF exportieren…',
    menuPrint: 'Drucken…',
    menuClose: 'Schließen',
    menuQuit: 'Beenden',
    menuEdit: 'Bearbeiten',
    menuUndo: 'Rückgängig',
    menuRedo: 'Wiederholen',
    closeUnsavedMsg: '{count} nicht gespeicherte Änderung(en)',
    closeUnsavedDetail: 'Ihre Änderungen gehen verloren, wenn Sie ohne Speichern schließen.',
    btnDontSave: 'Nicht speichern',
    btnCancel: 'Abbrechen',
    csvSaveAsNotice:
      'CSV-Dateien können keine Formatierung speichern – als .xlsx speichern, um alle Änderungen zu behalten.',
    menuExportCsv: 'CSV exportieren…',
    filterCsv: 'CSV (Trennzeichen-getrennt)',
    csvFormulaLossMsg: 'Dieses Blatt enthält Formeln, die das CSV-Format nicht speichern kann.',
    csvFormulaLossDetail:
      'CSV speichert nur reine Werte – Formeln werden durch ihre aktuellen Ergebnisse ersetzt, und die Formatierung geht verloren.',
    csvKeepXlsxBtn: 'Als .xlsx speichern',
    csvContinueBtn: 'Als CSV fortfahren',
    csvActiveSheetOnlyNotice:
      'CSV-Dateien enthalten nur ein Blatt – nur das aktive Blatt „{name}“ wird exportiert.',
    csvKeepFormatMsg: 'Weiter im CSV-Format speichern?',
    csvKeepFormatDetail:
      'CSV speichert nur die Werte eines einzelnen Blatts – Formeln, Formatierungen und weitere Blätter werden nicht in der .csv-Datei gespeichert.',
  },
  es: {
    filterSpreadsheets: 'Hojas de cálculo',
    filterXlsx: 'Libros de Excel',
    filterXlsm: 'Libros de Excel habilitados para macros',
    errFileTooLarge: 'El archivo supera el límite de tamaño',
    errImgAbsPath: 'La ruta de la imagen debe ser absoluta.',
    errImgNotFound: 'No se encontró el archivo de imagen: {path}',
    errImgTooLarge20: 'La imagen supera los 20 MB y no se puede insertar.',
    errImgBadType: 'El archivo no es una imagen PNG/JPEG/GIF.',
    errDiskChanged: 'El libro cambió en el disco después de abrirse; usa Guardar como en su lugar.',
    autosaveFoundTitle: 'Se encontró una versión recuperada',
    autosaveFoundBody:
      'Hay cambios sin guardar de la última sesión. ¿Restaurar la versión autoguardada? Tras restaurar, guardar sobrescribirá el archivo original.',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    menuFile: 'Archivo',
    menuOpenWorkbook: 'Abrir libro…',
    menuSave: 'Guardar',
    menuSaveAs: 'Guardar como…',
    menuExportPdf: 'Exportar a PDF…',
    menuPrint: 'Imprimir…',
    menuClose: 'Cerrar',
    menuQuit: 'Salir',
    menuEdit: 'Edición',
    menuUndo: 'Deshacer',
    menuRedo: 'Rehacer',
    closeUnsavedMsg: '{count} cambio(s) sin guardar',
    closeUnsavedDetail: 'Los cambios se perderán si cierras sin guardar.',
    btnDontSave: 'No guardar',
    btnCancel: 'Cancelar',
    csvSaveAsNotice:
      'El formato CSV no conserva el formato: guarda como .xlsx para conservar todos tus cambios.',
    menuExportCsv: 'Exportar a CSV…',
    filterCsv: 'CSV (delimitado por comas)',
    csvFormulaLossMsg: 'Esta hoja contiene fórmulas que el formato CSV no puede conservar.',
    csvFormulaLossDetail:
      'El CSV solo conserva valores: las fórmulas se sustituyen por sus resultados actuales y el formato se pierde.',
    csvKeepXlsxBtn: 'Guardar como .xlsx',
    csvContinueBtn: 'Continuar como CSV',
    csvActiveSheetOnlyNotice:
      'Los archivos CSV solo contienen una hoja: solo se exportará la hoja activa «{name}».',
    csvKeepFormatMsg: '¿Seguir guardando en formato CSV?',
    csvKeepFormatDetail:
      'CSV solo conserva los valores de una única hoja: las fórmulas, el formato y las hojas adicionales no se guardan en el archivo .csv.',
  },
  th: {
    filterSpreadsheets: 'สเปรดชีต',
    filterXlsx: 'เวิร์กบุ๊ก Excel',
    filterXlsm: 'เวิร์กบุ๊ก Excel ที่เปิดใช้งานแมโคร',
    errFileTooLarge: 'ไฟล์มีขนาดเกินขีดจำกัด',
    errImgAbsPath: 'เส้นทางรูปภาพต้องเป็นเส้นทางแบบสัมบูรณ์',
    errImgNotFound: 'ไม่พบไฟล์รูปภาพ: {path}',
    errImgTooLarge20: 'รูปภาพเกิน 20MB ไม่สามารถแทรกได้',
    errImgBadType: 'ไฟล์นี้ไม่ใช่รูปภาพ PNG/JPEG/GIF',
    errDiskChanged: 'เวิร์กบุ๊กถูกเปลี่ยนแปลงบนดิสก์หลังจากเปิด — โปรดใช้บันทึกเป็นแทน',
    autosaveFoundTitle: 'พบเวอร์ชันกู้คืนอัตโนมัติ',
    autosaveFoundBody:
      'มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึกจากครั้งก่อน ต้องการกู้คืนหรือไม่? หลังกู้คืน การบันทึกจะเขียนทับไฟล์ต้นฉบับ',
    autosaveRestore: 'กู้คืน',
    autosaveDiscard: 'ละทิ้ง',
    menuFile: 'ไฟล์',
    menuOpenWorkbook: 'เปิดเวิร์กบุ๊ก…',
    menuSave: 'บันทึก',
    menuSaveAs: 'บันทึกเป็น…',
    menuExportPdf: 'ส่งออก PDF…',
    menuPrint: 'พิมพ์…',
    menuClose: 'ปิด',
    menuQuit: 'ออก',
    menuEdit: 'แก้ไข',
    menuUndo: 'เลิกทำ',
    menuRedo: 'ทำซ้ำ',
    closeUnsavedMsg: 'มีการเปลี่ยนแปลงที่ยังไม่ได้บันทึก {count} รายการ',
    closeUnsavedDetail: 'หากปิดโดยไม่บันทึก การเปลี่ยนแปลงเหล่านี้จะหายไป',
    btnDontSave: 'ไม่บันทึก',
    btnCancel: 'ยกเลิก',
    csvSaveAsNotice:
      'ไฟล์ CSV ไม่สามารถเก็บการจัดรูปแบบได้ — บันทึกเป็น .xlsx เพื่อเก็บการเปลี่ยนแปลงทั้งหมดของคุณ',
    menuExportCsv: 'ส่งออก CSV…',
    filterCsv: 'CSV (คั่นด้วยเครื่องหมายจุลภาค)',
    csvFormulaLossMsg: 'ชีตนี้มีสูตรที่รูปแบบ CSV เก็บไว้ไม่ได้',
    csvFormulaLossDetail:
      'CSV เก็บเฉพาะค่าเท่านั้น — สูตรจะถูกแทนที่ด้วยผลลัพธ์ปัจจุบัน และการจัดรูปแบบจะหายไป',
    csvKeepXlsxBtn: 'บันทึกเป็น .xlsx',
    csvContinueBtn: 'บันทึกเป็น CSV ต่อไป',
    csvActiveSheetOnlyNotice:
      'ไฟล์ CSV มีได้เพียงชีตเดียว — จะส่งออกเฉพาะชีตที่ใช้งานอยู่ “{name}” เท่านั้น',
    csvKeepFormatMsg: 'บันทึกเป็นรูปแบบ CSV ต่อไปหรือไม่',
    csvKeepFormatDetail:
      'CSV เก็บเฉพาะค่าของชีตเดียวเท่านั้น — สูตร การจัดรูปแบบ และชีตอื่น ๆ จะไม่ถูกบันทึกลงในไฟล์ .csv',
  },
  id: {
    filterSpreadsheets: 'Lembar bentang',
    filterXlsx: 'Buku kerja Excel',
    filterXlsm: 'Buku kerja Excel dengan makro aktif',
    errFileTooLarge: 'File melebihi batas ukuran',
    errImgAbsPath: 'Jalur gambar harus berupa jalur absolut.',
    errImgNotFound: 'File gambar tidak ditemukan: {path}',
    errImgTooLarge20: 'Gambar melebihi 20MB dan tidak dapat disisipkan.',
    errImgBadType: 'File ini bukan gambar PNG/JPEG/GIF.',
    errDiskChanged: 'Buku kerja berubah di disk setelah dibuka — gunakan Simpan Sebagai.',
    autosaveFoundTitle: 'Versi pemulihan ditemukan',
    autosaveFoundBody:
      'Ada perubahan yang belum disimpan dari sesi terakhir. Pulihkan versi tersimpan otomatis? Setelah dipulihkan, menyimpan akan menimpa file asli.',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    menuFile: 'File',
    menuOpenWorkbook: 'Buka Buku Kerja…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuExportPdf: 'Ekspor PDF…',
    menuPrint: 'Cetak…',
    menuClose: 'Tutup',
    menuQuit: 'Keluar',
    menuEdit: 'Edit',
    menuUndo: 'Urungkan',
    menuRedo: 'Ulangi',
    closeUnsavedMsg: '{count} perubahan belum disimpan',
    closeUnsavedDetail: 'Perubahan Anda akan hilang jika menutup tanpa menyimpan.',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    csvSaveAsNotice:
      'File CSV tidak dapat menyimpan pemformatan — simpan sebagai .xlsx untuk mempertahankan semua perubahan Anda.',
    menuExportCsv: 'Ekspor CSV…',
    filterCsv: 'CSV (dipisahkan koma)',
    csvFormulaLossMsg: 'Lembar ini berisi rumus yang tidak dapat disimpan dalam format CSV.',
    csvFormulaLossDetail:
      'CSV hanya menyimpan nilai — rumus diganti dengan hasil saat ini, dan pemformatan akan hilang.',
    csvKeepXlsxBtn: 'Simpan sebagai .xlsx',
    csvContinueBtn: 'Lanjutkan sebagai CSV',
    csvActiveSheetOnlyNotice:
      'File CSV hanya memuat satu lembar — hanya lembar aktif “{name}” yang akan diekspor.',
    csvKeepFormatMsg: 'Terus menyimpan dalam format CSV?',
    csvKeepFormatDetail:
      'CSV hanya menyimpan nilai dari satu lembar — rumus, pemformatan, dan lembar tambahan tidak disimpan ke file .csv.',
  },
  ru: {
    filterSpreadsheets: 'Электронные таблицы',
    filterXlsx: 'Книги Excel',
    filterXlsm: 'Книги Excel с поддержкой макросов',
    errFileTooLarge: 'Файл превышает предельный размер',
    errImgAbsPath: 'Путь к изображению должен быть абсолютным.',
    errImgNotFound: 'Файл изображения не найден: {path}',
    errImgTooLarge20: 'Изображение превышает 20 МБ и не может быть вставлено.',
    errImgBadType: 'Этот файл не является изображением PNG/JPEG/GIF.',
    errDiskChanged: 'Книга была изменена на диске после открытия — используйте «Сохранить как».',
    autosaveFoundTitle: 'Найдена восстановленная версия',
    autosaveFoundBody:
      'Есть несохранённые изменения из прошлого сеанса. Восстановить автосохранённую версию? После восстановления сохранение перезапишет исходный файл.',
    autosaveRestore: 'Восстановить',
    autosaveDiscard: 'Отклонить',
    menuFile: 'Файл',
    menuOpenWorkbook: 'Открыть книгу…',
    menuSave: 'Сохранить',
    menuSaveAs: 'Сохранить как…',
    menuExportPdf: 'Экспорт в PDF…',
    menuPrint: 'Печать…',
    menuClose: 'Закрыть',
    menuQuit: 'Выход',
    menuEdit: 'Правка',
    menuUndo: 'Отменить',
    menuRedo: 'Повторить',
    closeUnsavedMsg: 'Несохранённых изменений: {count}',
    closeUnsavedDetail: 'Если закрыть без сохранения, эти изменения будут потеряны.',
    btnDontSave: 'Не сохранять',
    btnCancel: 'Отмена',
    csvSaveAsNotice:
      'Формат CSV не сохраняет форматирование — сохраните в .xlsx, чтобы не потерять изменения.',
    menuExportCsv: 'Экспорт в CSV…',
    filterCsv: 'CSV (разделители — запятые)',
    csvFormulaLossMsg: 'Этот лист содержит формулы, которые формат CSV не сохраняет.',
    csvFormulaLossDetail:
      'CSV сохраняет только значения — формулы заменяются текущими результатами, а форматирование теряется.',
    csvKeepXlsxBtn: 'Сохранить как .xlsx',
    csvContinueBtn: 'Продолжить в CSV',
    csvActiveSheetOnlyNotice:
      'Файлы CSV содержат только один лист — будет экспортирован только активный лист «{name}».',
    csvKeepFormatMsg: 'Продолжить сохранение в формате CSV?',
    csvKeepFormatDetail:
      'CSV сохраняет только значения одного листа — формулы, форматирование и дополнительные листы не сохраняются в файле .csv.',
  },
  ar: {
    filterSpreadsheets: 'جداول البيانات',
    filterXlsx: 'مصنفات Excel',
    filterXlsm: 'مصنفات Excel ممكّنة بوحدات الماكرو',
    errFileTooLarge: 'الملف يتجاوز حد الحجم',
    errImgAbsPath: 'يجب أن يكون مسار الصورة مسارًا مطلقًا.',
    errImgNotFound: 'لم يتم العثور على ملف الصورة: {path}',
    errImgTooLarge20: 'الصورة تتجاوز 20 ميغابايت ولا يمكن إدراجها.',
    errImgBadType: 'هذا الملف ليس صورة PNG/JPEG/GIF.',
    errDiskChanged: 'تم تغيير المصنف على القرص بعد فتحه — استخدم «حفظ باسم» بدلاً من ذلك.',
    autosaveFoundTitle: 'تم العثور على نسخة مستردة',
    autosaveFoundBody:
      'توجد تغييرات غير محفوظة من الجلسة الأخيرة. هل تريد استعادة النسخة المحفوظة تلقائيًا؟ بعد الاستعادة، سيؤدي الحفظ إلى استبدال الملف الأصلي.',
    autosaveRestore: 'استعادة',
    autosaveDiscard: 'تجاهل',
    menuFile: 'ملف',
    menuOpenWorkbook: 'فتح مصنف…',
    menuSave: 'حفظ',
    menuSaveAs: 'حفظ باسم…',
    menuExportPdf: 'تصدير PDF…',
    menuPrint: 'طباعة…',
    menuClose: 'إغلاق',
    menuQuit: 'إنهاء',
    menuEdit: 'تحرير',
    menuUndo: 'تراجع',
    menuRedo: 'إعادة',
    closeUnsavedMsg: 'يوجد {count} من التغييرات غير المحفوظة',
    closeUnsavedDetail: 'ستفقد هذه التغييرات إذا أغلقت دون حفظ.',
    btnDontSave: 'عدم الحفظ',
    btnCancel: 'إلغاء',
    csvSaveAsNotice: 'ملفات CSV لا تحتفظ بالتنسيق — احفظ بصيغة ‎.xlsx للاحتفاظ بجميع تغييراتك.',
    menuExportCsv: 'تصدير CSV…',
    filterCsv: 'CSV (محدد بفواصل)',
    csvFormulaLossMsg: 'تحتوي هذه الورقة على صيغ لا يمكن لتنسيق CSV الاحتفاظ بها.',
    csvFormulaLossDetail: 'يحتفظ CSV بالقيم فقط — تُستبدل الصيغ بنتائجها الحالية ويُفقد التنسيق.',
    csvKeepXlsxBtn: 'حفظ بصيغة .xlsx',
    csvContinueBtn: 'المتابعة بتنسيق CSV',
    csvActiveSheetOnlyNotice:
      'ملفات CSV تحتوي على ورقة واحدة فقط — سيتم تصدير الورقة النشطة «{name}» فقط.',
    csvKeepFormatMsg: 'هل تريد متابعة الحفظ بتنسيق CSV؟',
    csvKeepFormatDetail:
      'يحتفظ CSV بقيم ورقة واحدة فقط — لا تُحفظ الصيغ والتنسيق والأوراق الإضافية في ملف .csv.',
  },
  pt: {
    filterSpreadsheets: 'Planilhas',
    filterXlsx: 'Pastas de Trabalho do Excel',
    filterXlsm: 'Pastas de Trabalho Habilitadas para Macro do Excel',
    errFileTooLarge: 'O arquivo excede o limite de tamanho',
    errImgAbsPath: 'O caminho da imagem deve ser absoluto.',
    errImgNotFound: 'Arquivo de imagem não encontrado: {path}',
    errImgTooLarge20: 'A imagem excede 20MB e não pode ser inserida.',
    errImgBadType: 'O arquivo não é uma imagem PNG/JPEG/GIF.',
    errDiskChanged: 'A pasta de trabalho foi alterada no disco após ser aberta — use Salvar Como.',
    autosaveFoundTitle: 'Versão recuperada encontrada',
    autosaveFoundBody:
      'Há alterações não salvas da sua última sessão. Restaurar a versão salva automaticamente? Após restaurar, salvar sobrescreverá o arquivo original.',
    autosaveRestore: 'Restaurar',
    autosaveDiscard: 'Descartar',
    menuFile: 'Arquivo',
    menuOpenWorkbook: 'Abrir Pasta de Trabalho…',
    menuSave: 'Salvar',
    menuSaveAs: 'Salvar Como…',
    menuExportPdf: 'Exportar PDF…',
    menuPrint: 'Imprimir…',
    menuClose: 'Fechar',
    menuQuit: 'Sair',
    menuEdit: 'Editar',
    menuUndo: 'Desfazer',
    menuRedo: 'Refazer',
    closeUnsavedMsg: '{count} alteração(ões) não salva(s)',
    closeUnsavedDetail: 'Suas alterações serão perdidas se você fechar sem salvar.',
    btnDontSave: 'Não Salvar',
    btnCancel: 'Cancelar',
    csvSaveAsNotice:
      'Arquivos CSV não mantêm a formatação — salve como .xlsx para manter todas as suas alterações.',
    menuExportCsv: 'Exportar CSV…',
    filterCsv: 'CSV (separado por vírgulas)',
    csvFormulaLossMsg: 'Esta planilha contém fórmulas que o formato CSV não pode manter.',
    csvFormulaLossDetail:
      'O CSV mantém apenas valores — as fórmulas são substituídas pelos resultados atuais e a formatação é perdida.',
    csvKeepXlsxBtn: 'Salvar como .xlsx',
    csvContinueBtn: 'Continuar como CSV',
    csvActiveSheetOnlyNotice:
      'Arquivos CSV contêm apenas uma planilha — apenas a planilha ativa “{name}” será exportada.',
    csvKeepFormatMsg: 'Continuar salvando no formato CSV?',
    csvKeepFormatDetail:
      'O CSV mantém apenas os valores de uma única planilha — fórmulas, formatação e planilhas adicionais não são salvas no arquivo .csv.',
  },
  it: {
    filterSpreadsheets: 'Fogli di calcolo',
    filterXlsx: 'Cartelle di lavoro di Excel',
    filterXlsm: 'Cartelle di lavoro di Excel con attivazione macro',
    errFileTooLarge: 'Il file supera il limite di dimensione',
    errImgAbsPath: "Il percorso dell'immagine deve essere assoluto.",
    errImgNotFound: 'File immagine non trovato: {path}',
    errImgTooLarge20: "L'immagine supera i 20 MB e non può essere inserita.",
    errImgBadType: "Il file non è un'immagine PNG/JPEG/GIF.",
    errDiskChanged:
      "La cartella di lavoro è stata modificata sul disco dopo l'apertura — usa Salva con nome.",
    autosaveFoundTitle: 'Trovata versione recuperata',
    autosaveFoundBody:
      "Ci sono modifiche non salvate dall'ultima sessione. Ripristinare la versione salvata automaticamente? Dopo il ripristino, il salvataggio sovrascriverà il file originale.",
    autosaveRestore: 'Ripristina',
    autosaveDiscard: 'Ignora',
    menuFile: 'File',
    menuOpenWorkbook: 'Apri cartella di lavoro…',
    menuSave: 'Salva',
    menuSaveAs: 'Salva con nome…',
    menuExportPdf: 'Esporta PDF…',
    menuPrint: 'Stampa…',
    menuClose: 'Chiudi',
    menuQuit: 'Esci',
    menuEdit: 'Modifica',
    menuUndo: 'Annulla',
    menuRedo: 'Ripeti',
    closeUnsavedMsg: '{count} modifica/e non salvata/e',
    closeUnsavedDetail: 'Le modifiche andranno perse se chiudi senza salvare.',
    btnDontSave: 'Non salvare',
    btnCancel: 'Annulla',
    csvSaveAsNotice:
      'I file CSV non conservano la formattazione: salva come .xlsx per mantenere tutte le modifiche.',
    menuExportCsv: 'Esporta CSV…',
    filterCsv: 'CSV (delimitato da virgole)',
    csvFormulaLossMsg: 'Questo foglio contiene formule che il formato CSV non può conservare.',
    csvFormulaLossDetail:
      'Il CSV conserva solo i valori: le formule vengono sostituite dai risultati attuali e la formattazione viene persa.',
    csvKeepXlsxBtn: 'Salva come .xlsx',
    csvContinueBtn: 'Continua come CSV',
    csvActiveSheetOnlyNotice:
      'I file CSV contengono un solo foglio: verrà esportato solo il foglio attivo “{name}”.',
    csvKeepFormatMsg: 'Continuare a salvare in formato CSV?',
    csvKeepFormatDetail:
      'Il CSV conserva solo i valori di un singolo foglio: formule, formattazione e fogli aggiuntivi non vengono salvati nel file .csv.',
  },
  pl: {
    filterSpreadsheets: 'Arkusze kalkulacyjne',
    filterXlsx: 'Skoroszyty programu Excel',
    filterXlsm: 'Skoroszyty programu Excel z obsługą makr',
    errFileTooLarge: 'Plik przekracza limit rozmiaru',
    errImgAbsPath: 'Ścieżka obrazu musi być bezwzględna.',
    errImgNotFound: 'Nie znaleziono pliku obrazu: {path}',
    errImgTooLarge20: 'Obraz przekracza 20 MB i nie może zostać wstawiony.',
    errImgBadType: 'Plik nie jest obrazem PNG/JPEG/GIF.',
    errDiskChanged: 'Skoroszyt został zmieniony na dysku po otwarciu — użyj polecenia Zapisz jako.',
    autosaveFoundTitle: 'Znaleziono odzyskaną wersję',
    autosaveFoundBody:
      'Istnieją niezapisane zmiany z ostatniej sesji. Przywrócić wersję zapisaną automatycznie? Po przywróceniu zapisanie nadpisze oryginalny plik.',
    autosaveRestore: 'Przywróć',
    autosaveDiscard: 'Odrzuć',
    menuFile: 'Plik',
    menuOpenWorkbook: 'Otwórz skoroszyt…',
    menuSave: 'Zapisz',
    menuSaveAs: 'Zapisz jako…',
    menuExportPdf: 'Eksportuj PDF…',
    menuPrint: 'Drukuj…',
    menuClose: 'Zamknij',
    menuQuit: 'Zakończ',
    menuEdit: 'Edycja',
    menuUndo: 'Cofnij',
    menuRedo: 'Ponów',
    closeUnsavedMsg: 'Niezapisane zmiany: {count}',
    closeUnsavedDetail: 'Zmiany zostaną utracone, jeśli zamkniesz bez zapisywania.',
    btnDontSave: 'Nie zapisuj',
    btnCancel: 'Anuluj',
    csvSaveAsNotice:
      'Pliki CSV nie zachowują formatowania — zapisz jako .xlsx, aby zachować wszystkie zmiany.',
    menuExportCsv: 'Eksportuj CSV…',
    filterCsv: 'CSV (rozdzielany przecinkami)',
    csvFormulaLossMsg: 'Ten arkusz zawiera formuły, których format CSV nie zachowuje.',
    csvFormulaLossDetail:
      'CSV zachowuje tylko wartości — formuły są zastępowane bieżącymi wynikami, a formatowanie jest tracone.',
    csvKeepXlsxBtn: 'Zapisz jako .xlsx',
    csvContinueBtn: 'Kontynuuj jako CSV',
    csvActiveSheetOnlyNotice:
      'Pliki CSV zawierają tylko jeden arkusz — wyeksportowany zostanie tylko aktywny arkusz „{name}”.',
    csvKeepFormatMsg: 'Kontynuować zapisywanie w formacie CSV?',
    csvKeepFormatDetail:
      'CSV zachowuje tylko wartości jednego arkusza — formuły, formatowanie i dodatkowe arkusze nie są zapisywane w pliku .csv.',
  },
  cs: {
    filterSpreadsheets: 'Tabulky',
    filterXlsx: 'Sešity Excelu',
    filterXlsm: 'Sešity Excelu s podporou maker',
    errFileTooLarge: 'Soubor překračuje limit velikosti',
    errImgAbsPath: 'Cesta k obrázku musí být absolutní.',
    errImgNotFound: 'Soubor obrázku nebyl nalezen: {path}',
    errImgTooLarge20: 'Obrázek překračuje 20 MB a nelze ho vložit.',
    errImgBadType: 'Soubor není obrázek PNG/JPEG/GIF.',
    errDiskChanged: 'Sešit byl po otevření změněn na disku — použijte místo toho Uložit jako.',
    autosaveFoundTitle: 'Nalezena obnovená verze',
    autosaveFoundBody:
      'Z poslední relace existují neuložené změny. Obnovit automaticky uloženou verzi? Uložení po obnovení přepíše původní soubor.',
    autosaveRestore: 'Obnovit',
    autosaveDiscard: 'Zahodit',
    menuFile: 'Soubor',
    menuOpenWorkbook: 'Otevřít sešit…',
    menuSave: 'Uložit',
    menuSaveAs: 'Uložit jako…',
    menuExportPdf: 'Exportovat PDF…',
    menuPrint: 'Tisk…',
    menuClose: 'Zavřít',
    menuQuit: 'Ukončit',
    menuEdit: 'Úpravy',
    menuUndo: 'Zpět',
    menuRedo: 'Znovu',
    closeUnsavedMsg: 'Neuložené změny: {count}',
    closeUnsavedDetail: 'Pokud zavřete bez uložení, změny budou ztraceny.',
    btnDontSave: 'Neukládat',
    btnCancel: 'Zrušit',
    csvSaveAsNotice:
      'Soubory CSV nezachovávají formátování — uložením jako .xlsx zachováte všechny změny.',
    menuExportCsv: 'Exportovat CSV…',
    filterCsv: 'CSV (oddělený čárkami)',
    csvFormulaLossMsg: 'Tento list obsahuje vzorce, které formát CSV nezachová.',
    csvFormulaLossDetail:
      'CSV zachovává pouze hodnoty — vzorce se nahradí aktuálními výsledky a formátování se ztratí.',
    csvKeepXlsxBtn: 'Uložit jako .xlsx',
    csvContinueBtn: 'Pokračovat jako CSV',
    csvActiveSheetOnlyNotice:
      'Soubory CSV obsahují jen jeden list — exportován bude pouze aktivní list „{name}“.',
    csvKeepFormatMsg: 'Pokračovat v ukládání ve formátu CSV?',
    csvKeepFormatDetail:
      'CSV zachovává pouze hodnoty jednoho listu — vzorce, formátování a další listy se do souboru .csv neuloží.',
  },
  nl: {
    filterSpreadsheets: 'Spreadsheets',
    filterXlsx: 'Excel-werkmappen',
    filterXlsm: "Excel-werkmappen met macro's",
    errFileTooLarge: 'Bestand overschrijdt de maximale grootte',
    errImgAbsPath: 'Het afbeeldingspad moet absoluut zijn.',
    errImgNotFound: 'Afbeeldingsbestand niet gevonden: {path}',
    errImgTooLarge20: 'De afbeelding is groter dan 20 MB en kan niet worden ingevoegd.',
    errImgBadType: 'Het bestand is geen PNG/JPEG/GIF-afbeelding.',
    errDiskChanged:
      'De werkmap is op de schijf gewijzigd nadat deze was geopend — gebruik Opslaan als.',
    autosaveFoundTitle: 'Herstelde versie gevonden',
    autosaveFoundBody:
      'Er zijn niet-opgeslagen wijzigingen van uw laatste sessie. De automatisch opgeslagen versie herstellen? Na herstel overschrijft opslaan het originele bestand.',
    autosaveRestore: 'Herstellen',
    autosaveDiscard: 'Negeren',
    menuFile: 'Bestand',
    menuOpenWorkbook: 'Werkmap openen…',
    menuSave: 'Opslaan',
    menuSaveAs: 'Opslaan als…',
    menuExportPdf: 'PDF exporteren…',
    menuPrint: 'Afdrukken…',
    menuClose: 'Sluiten',
    menuQuit: 'Stoppen',
    menuEdit: 'Bewerken',
    menuUndo: 'Ongedaan maken',
    menuRedo: 'Opnieuw',
    closeUnsavedMsg: '{count} niet-opgeslagen wijziging(en)',
    closeUnsavedDetail: 'Uw wijzigingen gaan verloren als u sluit zonder op te slaan.',
    btnDontSave: 'Niet opslaan',
    btnCancel: 'Annuleren',
    csvSaveAsNotice:
      'CSV-bestanden bewaren geen opmaak — sla op als .xlsx om al uw wijzigingen te behouden.',
    menuExportCsv: 'CSV exporteren…',
    filterCsv: 'CSV (kommagescheiden)',
    csvFormulaLossMsg: 'Dit blad bevat formules die het CSV-formaat niet kan bewaren.',
    csvFormulaLossDetail:
      'CSV bewaart alleen waarden — formules worden vervangen door hun huidige resultaten en opmaak gaat verloren.',
    csvKeepXlsxBtn: 'Opslaan als .xlsx',
    csvContinueBtn: 'Doorgaan als CSV',
    csvActiveSheetOnlyNotice:
      'CSV-bestanden bevatten slechts één blad — alleen het actieve blad “{name}” wordt geëxporteerd.',
    csvKeepFormatMsg: 'Doorgaan met opslaan in CSV-indeling?',
    csvKeepFormatDetail:
      'CSV bewaart alleen de waarden van één blad — formules, opmaak en extra bladen worden niet in het .csv-bestand opgeslagen.',
  },
  ms: {
    filterSpreadsheets: 'Hamparan',
    filterXlsx: 'Buku Kerja Excel',
    filterXlsm: 'Buku Kerja Excel Didayakan Makro',
    errFileTooLarge: 'Fail melebihi had saiz',
    errImgAbsPath: 'Laluan imej mestilah laluan mutlak.',
    errImgNotFound: 'Fail imej tidak ditemui: {path}',
    errImgTooLarge20: 'Imej melebihi 20MB dan tidak boleh disisipkan.',
    errImgBadType: 'Fail ini bukan imej PNG/JPEG/GIF.',
    errDiskChanged: 'Buku kerja telah diubah pada cakera selepas dibuka — gunakan Simpan Sebagai.',
    autosaveFoundTitle: 'Versi pulihan ditemui',
    autosaveFoundBody:
      'Terdapat perubahan yang belum disimpan daripada sesi terakhir anda. Pulihkan versi yang disimpan secara automatik? Selepas pemulihan, menyimpan akan menulis ganti fail asal.',
    autosaveRestore: 'Pulihkan',
    autosaveDiscard: 'Buang',
    menuFile: 'Fail',
    menuOpenWorkbook: 'Buka Buku Kerja…',
    menuSave: 'Simpan',
    menuSaveAs: 'Simpan Sebagai…',
    menuExportPdf: 'Eksport PDF…',
    menuPrint: 'Cetak…',
    menuClose: 'Tutup',
    menuQuit: 'Keluar',
    menuEdit: 'Edit',
    menuUndo: 'Buat Asal',
    menuRedo: 'Buat Semula',
    closeUnsavedMsg: '{count} perubahan belum disimpan',
    closeUnsavedDetail: 'Perubahan anda akan hilang jika anda menutup tanpa menyimpan.',
    btnDontSave: 'Jangan Simpan',
    btnCancel: 'Batal',
    csvSaveAsNotice:
      'Fail CSV tidak dapat menyimpan pemformatan — simpan sebagai .xlsx untuk mengekalkan semua perubahan anda.',
    menuExportCsv: 'Eksport CSV…',
    filterCsv: 'CSV (dipisahkan koma)',
    csvFormulaLossMsg:
      'Helaian ini mengandungi formula yang tidak dapat disimpan dalam format CSV.',
    csvFormulaLossDetail:
      'CSV hanya menyimpan nilai — formula digantikan dengan hasil semasa, dan pemformatan akan hilang.',
    csvKeepXlsxBtn: 'Simpan sebagai .xlsx',
    csvContinueBtn: 'Teruskan sebagai CSV',
    csvActiveSheetOnlyNotice:
      'Fail CSV hanya mengandungi satu helaian — hanya helaian aktif “{name}” akan dieksport.',
    csvKeepFormatMsg: 'Terus simpan dalam format CSV?',
    csvKeepFormatDetail:
      'CSV hanya menyimpan nilai satu helaian — formula, pemformatan dan helaian tambahan tidak disimpan ke fail .csv.',
  },
  he: {
    filterSpreadsheets: 'גיליונות אלקטרוניים',
    filterXlsx: 'חוברות עבודה של Excel',
    filterXlsm: 'חוברות עבודה של Excel מותאמות מאקרו',
    errFileTooLarge: 'הקובץ חורג ממגבלת הגודל',
    errImgAbsPath: 'נתיב התמונה חייב להיות מוחלט.',
    errImgNotFound: 'קובץ התמונה לא נמצא: {path}',
    errImgTooLarge20: 'התמונה חורגת מ-20MB ולא ניתן להוסיף אותה.',
    errImgBadType: 'הקובץ אינו תמונת PNG/JPEG/GIF.',
    errDiskChanged: 'חוברת העבודה השתנתה בדיסק לאחר פתיחתה — השתמש בשמירה בשם.',
    autosaveFoundTitle: 'נמצאה גרסה משוחזרת',
    autosaveFoundBody:
      'קיימים שינויים שלא נשמרו מהפעלה הקודמת. לשחזר את הגרסה שנשמרה אוטומטית? לאחר השחזור, שמירה תדרוס את הקובץ המקורי.',
    autosaveRestore: 'שחזר',
    autosaveDiscard: 'התעלם',
    menuFile: 'קובץ',
    menuOpenWorkbook: 'פתיחת חוברת עבודה…',
    menuSave: 'שמירה',
    menuSaveAs: 'שמירה בשם…',
    menuExportPdf: 'ייצוא PDF…',
    menuPrint: 'הדפסה…',
    menuClose: 'סגירה',
    menuQuit: 'יציאה',
    menuEdit: 'עריכה',
    menuUndo: 'בטל',
    menuRedo: 'בצע שוב',
    closeUnsavedMsg: '{count} שינויים שלא נשמרו',
    closeUnsavedDetail: 'השינויים שלך יאבדו אם תסגור בלי לשמור.',
    btnDontSave: 'אל תשמור',
    btnCancel: 'ביטול',
    csvSaveAsNotice: 'קובצי CSV אינם שומרים עיצוב — שמרו כ‑.xlsx כדי לשמור על כל השינויים.',
    menuExportCsv: 'ייצוא CSV…',
    filterCsv: 'CSV (מופרד באמצעות פסיקים)',
    csvFormulaLossMsg: 'גיליון זה מכיל נוסחאות שתבנית CSV אינה יכולה לשמור.',
    csvFormulaLossDetail:
      'CSV שומר ערכים בלבד — נוסחאות מוחלפות בתוצאות הנוכחיות שלהן, והעיצוב אובד.',
    csvKeepXlsxBtn: 'שמירה כ-.xlsx',
    csvContinueBtn: 'המשך שמירה כ-CSV',
    csvActiveSheetOnlyNotice: 'קובצי CSV מכילים גיליון אחד בלבד — רק הגיליון הפעיל "{name}" ייוצא.',
    csvKeepFormatMsg: 'להמשיך לשמור בתבנית CSV?',
    csvKeepFormatDetail:
      'CSV שומר רק את הערכים של גיליון אחד — נוסחאות, עיצוב וגיליונות נוספים אינם נשמרים בקובץ ה-.csv.',
  },
  hi: {
    filterSpreadsheets: 'स्प्रेडशीट',
    filterXlsx: 'Excel कार्यपुस्तिकाएँ',
    filterXlsm: 'Excel मैक्रो-सक्षम कार्यपुस्तिकाएँ',
    errFileTooLarge: 'फ़ाइल आकार सीमा से अधिक है',
    errImgAbsPath: 'छवि पथ निरपेक्ष होना चाहिए।',
    errImgNotFound: 'छवि फ़ाइल नहीं मिली: {path}',
    errImgTooLarge20: 'छवि 20MB से अधिक है और सम्मिलित नहीं की जा सकती।',
    errImgBadType: 'यह फ़ाइल PNG/JPEG/GIF छवि नहीं है।',
    errDiskChanged:
      'खोले जाने के बाद कार्यपुस्तिका डिस्क पर बदल गई — इसके बजाय इस रूप में सहेजें का उपयोग करें।',
    autosaveFoundTitle: 'पुनर्प्राप्त संस्करण मिला',
    autosaveFoundBody:
      'आपके पिछले सत्र से सहेजे नहीं गए परिवर्तन हैं। स्वतः सहेजा गया संस्करण पुनर्स्थापित करें? पुनर्स्थापना के बाद, सहेजने पर मूल फ़ाइल अधिलेखित हो जाएगी।',
    autosaveRestore: 'पुनर्स्थापित करें',
    autosaveDiscard: 'छोड़ें',
    menuFile: 'फ़ाइल',
    menuOpenWorkbook: 'कार्यपुस्तिका खोलें…',
    menuSave: 'सहेजें',
    menuSaveAs: 'इस रूप में सहेजें…',
    menuExportPdf: 'PDF निर्यात करें…',
    menuPrint: 'प्रिंट करें…',
    menuClose: 'बंद करें',
    menuQuit: 'बाहर निकलें',
    menuEdit: 'संपादन',
    menuUndo: 'पूर्ववत करें',
    menuRedo: 'फिर से करें',
    closeUnsavedMsg: '{count} सहेजे नहीं गए परिवर्तन',
    closeUnsavedDetail: 'यदि आप बिना सहेजे बंद करते हैं तो आपके परिवर्तन खो जाएँगे।',
    btnDontSave: 'न सहेजें',
    btnCancel: 'रद्द करें',
    csvSaveAsNotice:
      'CSV फ़ाइलें फ़ॉर्मेटिंग सहेज नहीं सकतीं — सभी बदलाव बनाए रखने के लिए .xlsx के रूप में सहेजें।',
    menuExportCsv: 'CSV निर्यात करें…',
    filterCsv: 'CSV (अल्पविराम द्वारा सीमांकित)',
    csvFormulaLossMsg: 'इस शीट में ऐसे सूत्र हैं जिन्हें CSV प्रारूप सहेज नहीं सकता।',
    csvFormulaLossDetail:
      'CSV केवल मान रखता है — सूत्र उनके वर्तमान परिणामों से बदल दिए जाते हैं और फ़ॉर्मेटिंग खो जाती है।',
    csvKeepXlsxBtn: '.xlsx के रूप में सहेजें',
    csvContinueBtn: 'CSV के रूप में जारी रखें',
    csvActiveSheetOnlyNotice:
      'CSV फ़ाइलों में केवल एक शीट होती है — केवल सक्रिय शीट “{name}” निर्यात की जाएगी।',
    csvKeepFormatMsg: 'CSV प्रारूप में सहेजना जारी रखें?',
    csvKeepFormatDetail:
      'CSV केवल एक शीट के मान रखता है — सूत्र, स्वरूपण और अतिरिक्त शीट .csv फ़ाइल में सहेजे नहीं जाते।',
  },
  'zh-TW': {
    filterSpreadsheets: '電子試算表',
    filterXlsx: 'Excel 活頁簿',
    filterXlsm: 'Excel 啟用巨集的活頁簿',
    errFileTooLarge: '檔案超過大小上限',
    errImgAbsPath: '圖片路徑必須是絕對路徑。',
    errImgNotFound: '找不到圖片檔案: {path}',
    errImgTooLarge20: '圖片超過 20MB,不支援插入。',
    errImgBadType: '該檔案不是 PNG/JPEG/GIF 圖片。',
    errDiskChanged: '活頁簿在開啟後被磁碟上的變更覆蓋——請改用另存新檔。',
    autosaveFoundTitle: '發現自動復原版本',
    autosaveFoundBody:
      '上次工作階段有未儲存的變更。要復原自動儲存的版本嗎?復原後,儲存將直接覆寫原檔案。',
    autosaveRestore: '復原',
    autosaveDiscard: '放棄',
    menuFile: '檔案',
    menuOpenWorkbook: '開啟活頁簿…',
    menuSave: '儲存',
    menuSaveAs: '另存新檔…',
    menuExportPdf: '匯出 PDF…',
    menuPrint: '列印…',
    menuClose: '關閉',
    menuQuit: '結束',
    menuEdit: '編輯',
    menuUndo: '復原',
    menuRedo: '重做',
    closeUnsavedMsg: '有 {count} 處未儲存的修改',
    closeUnsavedDetail: '不儲存直接關閉,這些修改將遺失。',
    btnDontSave: '不儲存',
    btnCancel: '取消',
    csvSaveAsNotice: 'CSV 格式不保留樣式等格式修改——另存為 .xlsx 可保留全部內容。',
    menuExportCsv: '匯出 CSV…',
    filterCsv: 'CSV (逗號分隔)',
    csvFormulaLossMsg: '目前工作表包含公式,CSV 格式無法保留。',
    csvFormulaLossDetail: 'CSV 只保留純文字值——公式會被取代為目前計算結果,格式也會遺失。',
    csvKeepXlsxBtn: '另存為 .xlsx',
    csvContinueBtn: '繼續儲存為 CSV',
    csvActiveSheetOnlyNotice: 'CSV 檔案只包含一張工作表——只會匯出目前工作表「{name}」。',
    csvKeepFormatMsg: '要繼續以 CSV 格式儲存嗎?',
    csvKeepFormatDetail: 'CSV 只保留單張工作表的純值——公式、格式和其他工作表不會存入 .csv 檔案。',
  },
})
const tm = (key: Parameters<typeof tMain>[1], params?: Parameters<typeof tMain>[2]) =>
  tMain(getUiLang(), key, params)

interface SessionInfo {
  readonly path: string
  /// Byte-for-byte copy of the file as it was opened (in the OS temp dir).
  /// Saves patch this snapshot rather than the live path, so an external
  /// overwrite of the file can never corrupt the save base — and Save As
  /// stays usable after one. Removed when the session closes.
  readonly snapshotPath: string
  /// Digest of the snapshot (== the file at open time).
  readonly sha256: string
  readonly sheetNames: ReadonlyMap<string, string>
  readonly automaticRecoveryDisabled: boolean
  /// Set when the session opened a converted copy (.xls import): the
  /// first save routes through Save As, defaulting to this .xlsx path.
  readonly suggestSaveAs?: string
  /// The shell's "New spreadsheet" backing file in a temp directory: like an
  /// import the first Save is a Save As, but AutoSave and the recovery copy
  /// keep running (keyed on suggestSaveAs), and a quiet save writes in place.
  readonly unsavedNew?: boolean
  /// The converted copy came from a CSV: the Save As dialog explains that
  /// formatting requires .xlsx (CSV keeps values only).
  readonly csvImport?: boolean
  /// CSV session: the original .csv on disk. Save keeps the CSV identity —
  /// the xlsx save lands on the temp copy and the serialized csvContent is
  /// written back here.
  readonly csvSourcePath?: string
  /// Digest of the original .csv at open/save time — guards the write-back
  /// against external modification, like restoreTargetSha.
  readonly csvSourceSha?: string
  /// App-owned directory containing the converted CSV/XLS copy. Removed only
  /// after the sidecar session and its independent snapshot are closed.
  readonly importTempDir?: string
  /// Set when the session opened a restored crash-recovery copy: the restore
  /// prompt was the user's confirmation, so a plain Save silently writes back
  /// to this original path (no Save As detour).
  readonly restoreTarget?: string
  /// Digest of the original file at restore time — guards the silent
  /// write-back against external modification, mirroring the sha256 check on
  /// the session's own path.
  readonly restoreTargetSha?: string
  /// The opened file was 0 bytes: the session runs on a blank archive in the
  /// temp dir and restoreTarget is the empty original. Unlike a restored
  /// recovery copy, this session still needs its own crash-recovery copy,
  /// keyed by the original.
  readonly emptySource?: boolean
}

// ---- runtime configuration (paths differ when bundled into the shell) ----

interface SheetsRuntimeConfig {
  /** absolute path to the sheets preload bundle */
  preloadPath: string
  /** dev-server URL for the sheets renderer (wins over rendererFile) */
  rendererUrl?: string | undefined
  /** absolute path to the built sheets renderer index.html */
  rendererFile: string
  /** absolute path to the Rust xlsx-sidecar binary */
  sidecarPath?: string | undefined
  /** Shell router used to open exported files in a new GenOffice tab. */
  openGeneratedPath?: (path: string) => boolean
}

let runtime: SheetsRuntimeConfig = {
  preloadPath: join(__dirname, '../preload/index.js'),
  rendererUrl: process.env.ELECTRON_RENDERER_URL,
  rendererFile: join(__dirname, '../renderer/index.html'),
}

export function configureSheetsRuntime(config: SheetsRuntimeConfig): void {
  runtime = config
}

/** After writing an exported file: open it in the right tab
 * (shell) or reveal it in the folder (standalone). Tab-opening failure must
 * not report the write itself as failed — the file is already persisted. */
function openGeneratedFile(path: string): void {
  // Headless export must stay silent: no tab, no file-manager window
  // (same guard as markdown-main's openExportedPdf, #1815).
  if (isHeadlessMode()) return
  try {
    if (runtime.openGeneratedPath?.(path)) return
  } catch (err) {
    console.warn('[sheets] Failed to open generated file:', err)
  }
  shell.showItemInFolder(path)
}

let mainWindow: BrowserWindow | null = null
let sidecar: XlsxSidecarClient | null = null

/** the single real BrowserWindow hosting the tab strip, used as dialog parent in tab mode */
let sheetsShellWindow: BrowserWindow | null = null
export function setSheetsShellWindow(win: BrowserWindow | null): void {
  sheetsShellWindow = win
}

/** the window hosting a tab's WebContentsView when BrowserWindow.fromWebContents
 *  cannot tell (detached "Open in New Window" editors) */
let hostWindowHook: ((wc: WebContents) => BrowserWindow | undefined) | null = null
export function setSheetsHostWindowHook(
  fn: ((wc: WebContents) => BrowserWindow | undefined) | null,
): void {
  hostWindowHook = fn
}

function hostWindowFor(wc: WebContents): BrowserWindow | undefined {
  const own = hostWindowHook?.(wc) ?? BrowserWindow.fromWebContents(wc)
  if (own && !own.isDestroyed()) return own
  return sheetsShellWindow && !sheetsShellWindow.isDestroyed() ? sheetsShellWindow : undefined
}

interface SheetsTabSession {
  readonly webContents: WebContents
  readonly client: XlsxSidecarClient
  readonly sessions: Map<string, SessionInfo>
  /// Chunked uploads of large saves' cell edits, pending their save request.
  readonly saveTransfers: SaveEditsTransferStore
}

/** per-tab session state, keyed by webContents.id — replaces the old single-window closures
 * that `registerIpcHandlers`/`validateSender` used to capture, which broke as soon as a second
 * tab (or a closed-then-reopened tab) registered and overwrote the previous closure. */
/** Max .csv/.tsv bytes converted on open: stops a 500MB text file OOMing main before sidecar limits. */
const MAX_DELIMITED_IMPORT_BYTES = 32 * 1024 * 1024

const sheetsTabs = new Map<number, SheetsTabSession>()
let activeSheetsWebContents: WebContents | null = null
function sessionFor(event: IpcMainInvokeEvent): SheetsTabSession {
  const entry = sheetsTabs.get(event.sender.id)
  if (!entry) throw new Error('Untrusted IPC sender.')
  return entry
}

/// A save request referencing a chunked edit transfer gets the accumulated
/// edits spliced back in; the transfer is consumed either way.
function resolveTransferredEdits(
  entry: SheetsTabSession,
  request: WorkbookSaveRequest,
): WorkbookSaveRequest {
  if (request.editsTransferId === undefined) return request
  if (request.edits.length > 0) throw new Error('Save request mixes inline and transferred edits.')
  const edits = entry.saveTransfers.take(request.editsTransferId, request.sessionId)
  return { ...request, edits }
}

function dialogParent(event: IpcMainInvokeEvent): BrowserWindow | undefined {
  return hostWindowFor(event.sender)
}

async function openFileDialog(event: IpcMainInvokeEvent, options: OpenDialogOptions) {
  return showOpenDialogWithMemory(dialog, dialogParent(event), options)
}

async function saveFileDialog(event: IpcMainInvokeEvent, options: SaveDialogOptions) {
  // before any pick is remembered, bare-name suggestions anchor in the
  // configurable default save folder instead of Electron's Downloads pin
  return showSaveDialogWithMemory(
    dialog,
    dialogParent(event),
    options,
    configuredDefaultSaveDir(app),
  )
}

/** Workbook file extensions the open pipeline accepts. Single source for the
 *  picker filters, the merge-source check, and the crash re-open check. */
const WORKBOOK_EXTS = new Set(['xlsx', 'xlsm', 'xls', 'csv', 'tsv'])

/** register a tab's webContents/client pair and wire up cleanup on teardown */
/** Clients whose process-death notification is already wired. A sidecar client
 *  is shared by every tab (the `sidecar ?? new XlsxSidecarClient(...)` reuse),
 *  so the wiring is per client and the crash is broadcast to the tabs using
 *  it — not attached once per tab, which would notify N times over. */
const crashNotifiedClients = new WeakSet<XlsxSidecarClient>()

/** Tell every live tab that the sidecar died, so it can re-open its workbook
 *  and adopt a live session. Positive signal: a workbook the app closed on
 *  purpose (closeWorkbook) or swapped during Save never reaches the client,
 *  because those only send a `close` command down the live pipe. */
function broadcastSidecarCrash(client: XlsxSidecarClient): void {
  for (const entry of sheetsTabs.values()) {
    if (entry.client !== client) continue
    if (entry.webContents.isDestroyed()) continue
    entry.webContents.send(IPC_CHANNELS.sidecarCrashed)
  }
}

function registerSheetsSession(webContents: WebContents, client: XlsxSidecarClient): void {
  if (!crashNotifiedClients.has(client)) {
    crashNotifiedClients.add(client)
    client.onProcessExit(() => broadcastSidecarCrash(client))
  }
  sheetsTabs.set(webContents.id, {
    webContents,
    client,
    sessions: new Map(),
    saveTransfers: new SaveEditsTransferStore(),
  })
  activeSheetsWebContents = webContents
  webContents.once('destroyed', () => {
    const entry = sheetsTabs.get(webContents.id)
    sheetsTabs.delete(webContents.id)
    if (entry) {
      // Free pending chunked-save uploads with the tab (the sweep timer's
      // closure would otherwise keep them reachable until the idle expiry).
      entry.saveTransfers.dispose()
      void closeAllSessions(entry)
    }
    if (activeSheetsWebContents === webContents) activeSheetsWebContents = null
  })
}

export function getSheetsWindow(): BrowserWindow | null {
  return mainWindow
}

/** the webContents of whichever sheets tab most recently registered or activated */
export function getActiveSheetsWebContents(): WebContents | null {
  return activeSheetsWebContents
}

/** shell tab switching keeps menu actions routed at the visible sheets tab */
export function setActiveSheetsWebContents(wc: WebContents | null): void {
  activeSheetsWebContents = wc
}

/** Shell notification: an open view's file was renamed on disk (renamed in the
 *  Home list) — sync the matching session's path in that tab (later saves write
 *  the new file) and push the renderer to update the title-bar file name. */
export function sheetsFileRenamed(wc: WebContents, oldPath: string, newPath: string): void {
  const entry = sheetsTabs.get(wc.id)
  if (!entry) return
  let matched = false
  for (const [id, session] of entry.sessions) {
    // Converted copies (.csv / .xls / .tsv) and restored recovery copies keep
    // the user's file in a side field, not in `path`; move those too, or the
    // next Save recreates the file under the old name.
    let renamed = sessionAfterRename(session, oldPath, newPath)
    if (renamed === null) continue
    // An unsaved new workbook's tab renames its hidden backing file; the name
    // the user typed is what the first Save As should offer, and the crash copy
    // is keyed on that name.
    if (session.unsavedNew && session.path === oldPath && session.suggestSaveAs !== undefined) {
      const suggested = join(dirname(session.suggestSaveAs), basename(newPath))
      if (
        suggested !== session.suggestSaveAs &&
        !existsSync(suggested) &&
        !sheetsSuggestedPathTaken(suggested) &&
        retargetUnsavedNewRecovery(session.suggestSaveAs, suggested)
      )
        renamed = { ...renamed, suggestSaveAs: suggested }
    }
    entry.sessions.set(id, renamed)
    matched = true
  }
  if (matched) wc.send(IPC_CHANNELS.workbookRenamed, basename(newPath))
}

/**
 * Backing workbooks the shell created in a temp directory for "New
 * spreadsheet". The user has no file yet — it only comes into existence at the
 * first Save, which this routes through Save As — and a tab closed without
 * saving discards the temp directory with it. Same shape as an .xls/.tsv
 * import: the file the session edits is not the user's file.
 */
const unsavedNewWorkbooks = new Map<string, { suggestSaveAs: string; tempDir: string }>()

/**
 * shell: a would-be path already promised to another unsaved new workbook (open
 * tab, tab about to open, or a crash copy waiting to be re-offered). The file
 * itself does not exist yet, so uniqueness on disk is not enough.
 */
export function sheetsSuggestedPathTaken(path: string): boolean {
  for (const pending of unsavedNewWorkbooks.values())
    if (pending.suggestSaveAs === path) return true
  for (const tab of sheetsTabs.values())
    for (const session of tab.sessions.values()) if (session.suggestSaveAs === path) return true
  return readUnsavedNewIndex().includes(path)
}

/** shell: mark a backing workbook it just created as not-yet-saved (see above) */
export function markSheetsUnsavedNew(
  openPath: string,
  suggestSaveAs: string,
  tempDir: string,
): void {
  unsavedNewWorkbooks.set(openPath, { suggestSaveAs, tempDir })
}

/** shell: the backing workbook could not be written, release its name */
export function unmarkSheetsUnsavedNew(openPath: string): void {
  unsavedNewWorkbooks.delete(openPath)
}

/** shell hook: a tab opened a workbook (dialog or queued path) — used for tab titles/dedupe */
let workbookOpenedHook: ((wc: WebContents, path: string) => void) | null = null
export function setSheetsWorkbookOpenedHook(
  fn: ((wc: WebContents, path: string) => void) | null,
): void {
  workbookOpenedHook = fn
}

/** forward an application-menu File command into the sheets renderer */
export function sendSheetsMenuAction(
  action: 'open' | 'save' | 'save-as' | 'print' | 'export-pdf' | 'export-csv' | 'undo' | 'redo',
): void {
  activeSheetsWebContents?.send(IPC_CHANNELS.menuAction, action)
}

/** An already-mounted renderer polled for its queued workbook before one existed. */
export function nudgeQueuedWorkbook(contents: WebContents): void {
  contents.send(IPC_CHANNELS.menuAction, 'open')
}

function userDataPath(...parts: string[]): string {
  return join(app.getPath('userData'), ...parts)
}

// ── Crash recovery ──────────────────────────────────────────
// A dirty renderer asks for a recovery copy every 30s; it is written through the
// normal save pipeline (writeWorkbookTo) to a userData path, so it is a real .xlsx.
// A successful save removes it; opening a file whose copy is newer offers Restore.
const recoveryDir = () => userDataPath('sheets-autosave')
const recoveryPathFor = (filePath: string) =>
  join(recoveryDir(), `${createHash('sha1').update(filePath).digest('hex').slice(0, 16)}.xlsx`)

function clearWorkbookRecovery(filePath: string): void {
  try {
    unlinkSync(recoveryPathFor(filePath))
  } catch {
    /* nothing to clean */
  }
}

// An unsaved new workbook has no file for pendingRecoveryFor to compare
// against, so the would-be paths of its recovery copies are listed here and
// the shell re-offers them at launch (and "New spreadsheet" reuses the name).
const unsavedNewIndexPath = () => join(recoveryDir(), 'unsaved-new.json')

function readUnsavedNewIndex(): string[] {
  const list = readJson<unknown>(unsavedNewIndexPath(), [])
  return Array.isArray(list) ? list.filter((p): p is string => typeof p === 'string') : []
}

function writeUnsavedNewIndex(list: string[]): void {
  try {
    mkdirSync(recoveryDir(), { recursive: true })
    writeFileSync(unsavedNewIndexPath(), JSON.stringify(list))
  } catch (error) {
    console.warn('[sheets] unsaved-new index write failed:', error)
  }
}

function rememberUnsavedNewRecovery(suggestSaveAs: string): void {
  const list = readUnsavedNewIndex()
  if (!list.includes(suggestSaveAs)) writeUnsavedNewIndex([...list, suggestSaveAs])
}

function forgetUnsavedNewRecovery(suggestSaveAs: string): void {
  clearWorkbookRecovery(suggestSaveAs)
  const list = readUnsavedNewIndex()
  if (list.includes(suggestSaveAs)) writeUnsavedNewIndex(list.filter((p) => p !== suggestSaveAs))
}

/** shell, at launch: would-be paths of new workbooks whose recovery copy survived a crash */
export function pendingUnsavedNewRecoveries(): string[] {
  const live = readUnsavedNewIndex().filter((p) => existsSync(recoveryPathFor(p)))
  writeUnsavedNewIndex(live)
  return live
}

/** the would-be name changed: the crash copy keyed on it follows (false = copy stays under the old key) */
function retargetUnsavedNewRecovery(from: string, to: string): boolean {
  if (!pendingUnsavedNewRecoveryFor(from)) return true
  try {
    renameSync(recoveryPathFor(from), recoveryPathFor(to))
    rememberUnsavedNewRecovery(to)
    forgetUnsavedNewRecovery(from)
    return true
  } catch (err) {
    console.warn('[sheets] recovery copy rename failed:', err)
    return false
  }
}

function pendingUnsavedNewRecoveryFor(suggestSaveAs: string): string | null {
  const copy = recoveryPathFor(suggestSaveAs)
  return readUnsavedNewIndex().includes(suggestSaveAs) && existsSync(copy) ? copy : null
}

/// Restore/Discard choice for a pending recovery copy. Rendered as a styled
/// in-app dialog by the renderer (the native message box looks dated,
/// especially on Windows); strings ship pre-localized in the payload.
/// 'dismissed' (renderer gone before answering) opens the original file and
/// keeps the copy, so the offer repeats on the next open.
type RecoveryChoice = 'restore' | 'discard' | 'dismissed'

const recoveryPromptWaiters = new Map<number, (choice: RecoveryChoice) => void>()

function promptRecoveryRestore(
  contents: WebContents,
  filePath: string,
  recoveryPath: string,
): Promise<RecoveryChoice> {
  return new Promise((resolve) => {
    let savedAtMs = Date.now()
    try {
      savedAtMs = statSync(recoveryPath).mtimeMs
    } catch {
      /* copy vanished: the prompt still works, just without a precise time */
    }
    const settle = (choice: RecoveryChoice): void => {
      recoveryPromptWaiters.delete(contents.id)
      contents.removeListener('destroyed', onDestroyed)
      resolve(choice)
    }
    const onDestroyed = (): void => settle('dismissed')
    recoveryPromptWaiters.set(contents.id, settle)
    contents.once('destroyed', onDestroyed)
    contents.send(IPC_CHANNELS.recoveryPrompt, {
      title: tm('autosaveFoundTitle'),
      body: tm('autosaveFoundBody'),
      restoreLabel: tm('autosaveRestore'),
      discardLabel: tm('autosaveDiscard'),
      fileName: basename(filePath),
      savedAtMs,
    })
  })
}

/** Native message-box fallback for the rare open with no live renderer to draw the prompt. */
async function promptRecoveryRestoreNative(
  parent?: BrowserWindow | undefined,
): Promise<RecoveryChoice> {
  const options = {
    type: 'question' as const,
    buttons: [tm('autosaveRestore'), tm('autosaveDiscard')],
    defaultId: 0,
    cancelId: 1,
    message: tm('autosaveFoundTitle'),
    detail: tm('autosaveFoundBody'),
  }
  const answer = parent
    ? await dialog.showMessageBox(parent, options)
    : await dialog.showMessageBox(options)
  return answer.response === 0 ? 'restore' : 'discard'
}

/** Recovery copy newer than the file itself, i.e. unsaved work from a lost session. */
function pendingRecoveryFor(filePath: string): string | null {
  const copy = recoveryPathFor(filePath)
  try {
    if (!existsSync(copy)) return null
    if (statSync(copy).mtimeMs <= statSync(filePath).mtimeMs) {
      unlinkSync(copy)
      return null
    }
    return copy
  } catch {
    return null
  }
}

function readJson<T>(path: string, fallback: T): T {
  try {
    if (existsSync(path)) return JSON.parse(readFileSync(path, 'utf-8')) as T
  } catch {
    /* corrupted state file: fall back to defaults */
  }
  return fallback
}

// Dev-only automation hooks: a fixed CDP port for driving the app from test
// scripts, and a workbook path that bypasses the native file dialog.
const debugPort = app.isPackaged ? undefined : process.env.XLSX_DEBUG_PORT
if (debugPort) app.commandLine.appendSwitch('remote-debugging-port', debugPort)
let forcedWorkbookPath = app.isPackaged ? undefined : process.env.XLSX_OPEN_PATH
/** shell-queued workbook paths keyed by tab webContents id: a multi-select Open
 * creates several sheets tabs at once, so the path must be bound to its own tab
 * (a single global would be overwritten by the next iteration). One-shot, unlike
 * the sticky dev env/capture-server path above. */
const queuedWorkbookPaths = new Map<number, string>()

/** queue a workbook this tab's first selectWorkbook call opens without a dialog (shell routing) */
export function queueWorkbookForView(contents: WebContents, path: string): void {
  queuedWorkbookPaths.set(contents.id, path)
  contents.once('destroyed', () => {
    queuedWorkbookPaths.delete(contents.id)
  })
}

/** is the active tab still waiting for the renderer to consume a shell-queued workbook? */
export function hasActiveQueuedWorkbook(): boolean {
  return activeSheetsWebContents !== null && queuedWorkbookPaths.has(activeSheetsWebContents.id)
}

/** set by shell for home:new-sheet: renderer opens blank workbook instead of demo */
let pendingNewBlank = false

/** signal the next sheets renderer to open a new blank workbook (shell mode only) */
export function setSheetsNewBlank(): void {
  pendingNewBlank = true
}

// capturePage forces a renderer frame even when the window is occluded or on
// another Space, unlike CDP Page.captureScreenshot / macOS screencapture.
function startCaptureServer(): void {
  if (!debugPort) return
  const server = createServer((request, response) => {
    const url = new URL(request.url ?? '/', 'http://127.0.0.1')
    if (url.pathname === '/open') {
      forcedWorkbookPath = url.searchParams.get('path') ?? undefined
      response.writeHead(200)
      response.end('ok')
      return
    }
    // Drives the File menu from test scripts: CDP input can't reach native
    // menu accelerators, and osascript focus-stealing is flaky.
    if (url.pathname === '/menu') {
      const action = url.searchParams.get('action')
      if (
        action === 'open' ||
        action === 'save' ||
        action === 'save-as' ||
        action === 'print' ||
        action === 'export-pdf' ||
        action === 'export-csv' ||
        action === 'undo' ||
        action === 'redo'
      ) {
        sendSheetsMenuAction(action)
        response.writeHead(200)
        response.end('ok')
      } else {
        response.writeHead(400)
        response.end('unknown action')
      }
      return
    }
    const webContents = getActiveSheetsWebContents()
    if (url.pathname !== '/capture' || !webContents) {
      response.writeHead(404)
      response.end()
      return
    }
    webContents
      .capturePage()
      .then((image) => {
        response.writeHead(200, { 'Content-Type': 'image/png' })
        response.end(image.toPNG())
      })
      .catch((error: unknown) => {
        response.writeHead(500)
        response.end(String(error))
      })
  })
  server.listen(Number(debugPort) + 1, '127.0.0.1')
}

const sidecarOpenResultSchema = workbookFileSchema.omit({
  sha256: true,
  readOnly: true,
})

export async function createSheetsWindow(): Promise<BrowserWindow> {
  const client = sidecar ?? new XlsxSidecarClient(resolveSidecarPath())
  sidecar = client
  client.start()
  const window = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 720,
    minHeight: 550,
    show: false,
    title: 'GenOffice Sheets',
    // Traffic lights sit inside the toolbar row.
    ...(process.platform === 'darwin' ? { titleBarStyle: 'hiddenInset' as const } : {}),
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })
  mainWindow = window
  registerSheetsIpc()
  registerSheetsSession(window.webContents, client)
  window.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  window.webContents.on('will-navigate', (event) => event.preventDefault())
  if (!app.isPackaged) {
    window.webContents.on('console-message', (details) => {
      process.stderr.write(`[renderer:${details.level}] ${details.message}\n`)
    })
  }
  window.once('ready-to-show', () => window.show())
  window.on('close', (event) => {
    if (sheetsPendingEditCount(window.webContents.id) === 0) return
    event.preventDefault()
    void requestSheetsClose(window.webContents, window).then((proceed) => {
      // destroy() skips this handler on the way out (close() would re-enter
      // with the count possibly still non-zero after a discard).
      if (proceed && !window.isDestroyed()) window.destroy()
      // not proceeding vetoes any quit that was in flight: go back to
      // prompting on later closes (see resetSheetsShuttingDown)
      else resetSheetsShuttingDown()
    })
  })
  window.on('closed', () => {
    mainWindow = null
  })

  await window.loadURL(rendererUrl(runtime.rendererUrl, 'sheets'))
  return window
}

/** hidden export windows: webContents id -> the PDF path the renderer must write */
const headlessExportTargets = new Map<number, string>()
/** settled by 'sheets:headless-export-done' (or by the renderer dying) */
const headlessExportWaiters = new Map<number, (result: HeadlessSheetsReport) => void>()

interface HeadlessSheetsReport {
  ok: boolean
  error?: string
}

/** End a headless run early (a failure the renderer can never observe). */
function failHeadlessExport(wcId: number, message: string): void {
  const settle = headlessExportWaiters.get(wcId)
  if (!settle) return
  headlessExportWaiters.delete(wcId)
  settle({ ok: false, error: message })
}

/**
 * Render `input` to `outPath` with no visible window: a hidden sheets
 * renderer opens the workbook through the normal queued-open path, lays the
 * active sheet out with its Page Layout settings and prints via the existing
 * hidden print window (main/pdf-export.ts).
 */
export async function exportSheetsPdfHeadless(
  input: string,
  outPath: string,
  timeoutMs = 600_000,
): Promise<void> {
  const client = sidecar ?? new XlsxSidecarClient(resolveSidecarPath())
  sidecar = client
  client.start()
  const win = new BrowserWindow({
    show: false,
    width: 1440,
    height: 900,
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })
  registerSheetsIpc()
  registerSheetsSession(win.webContents, client)
  win.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  win.webContents.on('will-navigate', (event) => event.preventDefault())
  const wcId = win.webContents.id
  queueWorkbookForView(win.webContents, input)
  headlessExportTargets.set(wcId, outPath)
  let timer: ReturnType<typeof setTimeout> | undefined
  try {
    const report = await new Promise<HeadlessSheetsReport>((resolve) => {
      headlessExportWaiters.set(wcId, resolve)
      win.webContents.on('render-process-gone', (_event, details) =>
        resolve({ ok: false, error: `sheets renderer stopped (${details.reason})` }),
      )
      timer = setTimeout(
        () => resolve({ ok: false, error: `sheets export timed out after ${timeoutMs}ms` }),
        timeoutMs,
      )
      void win.webContents.loadURL(rendererUrl(runtime.rendererUrl, 'sheets'))
    })
    if (!report.ok) throw new Error(report.error ?? 'sheets export failed')
  } finally {
    if (timer) clearTimeout(timer)
    headlessExportWaiters.delete(wcId)
    headlessExportTargets.delete(wcId)
    if (!win.isDestroyed()) win.destroy()
  }
}

/** tab-mode equivalent of createSheetsWindow: same runtime/IPC wiring, no BrowserWindow of its own. */
export function createSheetsView(options: { openingWorkbook?: boolean } = {}): WebContentsView {
  const client = sidecar ?? new XlsxSidecarClient(resolveSidecarPath())
  sidecar = client
  client.start()
  const view = new WebContentsView({
    webPreferences: {
      preload: runtime.preloadPath,
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      backgroundThrottling: false,
    },
  })
  registerSheetsIpc()
  registerSheetsSession(view.webContents, client)
  view.webContents.setWindowOpenHandler(() => ({ action: 'deny' }))
  view.webContents.on('will-navigate', (event) => event.preventDefault())
  if (!app.isPackaged) {
    view.webContents.on('console-message', (details) => {
      process.stderr.write(`[renderer:${details.level}] ${details.message}\n`)
    })
  }
  // mode=tab: the shell's tab strip owns the traffic lights / caption buttons,
  // so the ribbon must not reserve space for them
  void view.webContents.loadURL(
    rendererUrl(runtime.rendererUrl, 'sheets', {
      mode: 'tab',
      ...(options.openingWorkbook ? { openingWorkbook: '1' } : {}),
    }),
  )
  return view
}

// Close guard: the renderer mirrors its pending-save count here, used to show a
// save confirmation before closing the window/tab.
const pendingEditCounts = new Map<number, number>()
const closeSaveWaiters = new Map<number, (ok: boolean) => void>()
const trackedEditSenders = new Set<number>()

export function sheetsPendingEditCount(webContentsId: number): number {
  return pendingEditCounts.get(webContentsId) ?? 0
}

/**
 * Close guard for a sheets renderer: true means proceed with the close.
 * Clean → true; dirty → Save/Don't Save/Cancel dialog. Save asks the renderer to run
 * its journal save and waits for the outcome — a failed or canceled save
 * keeps the window open (the renderer already surfaced the error).
 */
/**
 * The app is shutting down (quit menu, SIGTERM from a restart/installer/killall,
 * SIGINT from a terminal). The close guard must not save then: nobody answered the
 * prompt, and a dialog raised during shutdown resolves to its default button, which
 * silently overwrote the user's original file. Unsaved work is covered
 * by the 30s recovery copy instead — the next launch offers to restore it.
 */
const shutdownLatch = new ShutdownLatch()

export function markSheetsShuttingDown(): void {
  shutdownLatch.mark()
}

/**
 * The quit was vetoed (a prevented window close that ended up not proceeding).
 * The latch must go back down: from now on closes are interactive again and
 * the save prompt must run — leaving it set would silently discard unsaved
 * edits on every later close.
 */
export function resetSheetsShuttingDown(): void {
  shutdownLatch.reset()
}

app.on('before-quit', markSheetsShuttingDown)
for (const signal of ['SIGTERM', 'SIGINT'] as const) {
  process.on(signal, () => {
    shutdownLatch.mark()
    app.quit()
  })
}

export async function requestSheetsClose(
  contents: WebContents,
  parent?: BrowserWindow | null,
): Promise<boolean> {
  // A quiet AutoSave leaves an unsaved new workbook's journal empty, but its
  // work still sits on no user-visible path: closing it is discarding it.
  const autosavedNew = [...(sheetsTabs.get(contents.id)?.sessions.values() ?? [])].filter(
    (s) =>
      s.unsavedNew &&
      s.suggestSaveAs !== undefined &&
      pendingUnsavedNewRecoveryFor(s.suggestSaveAs),
  ).length
  const count = (pendingEditCounts.get(contents.id) ?? 0) + autosavedNew
  const decision = closeGuardDecision({
    pendingEdits: count,
    destroyed: contents.isDestroyed(),
    shuttingDown: shutdownLatch.active,
  })
  if (decision === 'proceed') return true
  const options = {
    // On macOS 'warning' shows the system warning triangle + app-icon badge
    type: 'warning' as const,
    message: tm('closeUnsavedMsg', { count }),
    detail: tm('closeUnsavedDetail'),
    buttons: [tm('menuSave'), tm('btnDontSave'), tm('btnCancel')],
    defaultId: 0,
    cancelId: 2,
    noLink: true,
  }
  const { response } =
    parent && !parent.isDestroyed()
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
  if (response === 2) return false
  if (response === 1) return true
  // The window went away (or a quit started) while the prompt was up: don't save
  if (shutdownLatch.active || contents.isDestroyed()) return true
  return await new Promise<boolean>((resolve) => {
    const timer = setTimeout(() => {
      closeSaveWaiters.delete(contents.id)
      resolve(false)
    }, 120_000)
    closeSaveWaiters.set(contents.id, (ok) => {
      clearTimeout(timer)
      resolve(ok)
    })
    contents.send(IPC_CHANNELS.closeSaveRequest)
  })
}

let coreIpcRegistered = false

export function registerSheetsIpc(): void {
  if (coreIpcRegistered) return
  coreIpcRegistered = true

  ipcMain.on(IPC_CHANNELS.recoveryPromptReply, (event, restore: unknown) => {
    recoveryPromptWaiters.get(event.sender.id)?.(restore === true ? 'restore' : 'discard')
  })

  ipcMain.on(IPC_CHANNELS.pendingEditsChanged, (event, count: unknown) => {
    if (typeof count !== 'number' || !Number.isFinite(count) || count < 0) return
    const senderId = event.sender.id
    pendingEditCounts.set(senderId, Math.floor(count))
    if (!trackedEditSenders.has(senderId)) {
      trackedEditSenders.add(senderId)
      event.sender.once('destroyed', () => {
        trackedEditSenders.delete(senderId)
        pendingEditCounts.delete(senderId)
        closeSaveWaiters.get(senderId)?.(false)
        closeSaveWaiters.delete(senderId)
      })
    }
  })

  ipcMain.on(IPC_CHANNELS.closeSaveResult, (event, ok: unknown) => {
    const waiter = closeSaveWaiters.get(event.sender.id)
    if (!waiter) return
    closeSaveWaiters.delete(event.sender.id)
    waiter(ok === true)
  })

  // shared with the other editor modules — last (identical) registration wins
  ipcMain.removeHandler('app:get-language')
  ipcMain.handle('app:get-language', () => getUiLang())

  /** returns true once when shell opened this tab for a new blank workbook */
  ipcMain.handle('sheets:consume-new-blank', () => {
    if (pendingNewBlank) {
      pendingNewBlank = false
      return true
    }
    return false
  })

  /**
   * Is a shell-queued workbook still waiting to be opened? The shell's 'open'
   * nudge loop gives up after 30s; on slow dev cold starts (vite compiles the
   * renderer on demand) Univer mounts later than that and the queued path
   * would strand the tab as a blank in-memory workbook. The renderer polls
   * this once it is ready and triggers the open itself.
   */
  ipcMain.handle('sheets:has-queued-workbook', (event) => queuedWorkbookPaths.has(event.sender.id))

  // ---- headless export mode (--headless-export) ----

  ipcMain.handle('sheets:consume-headless-export', (event): string | null => {
    const target = headlessExportTargets.get(event.sender.id) ?? null
    headlessExportTargets.delete(event.sender.id)
    return target
  })

  ipcMain.on('sheets:headless-export-done', (event, result: unknown) => {
    const settle = headlessExportWaiters.get(event.sender.id)
    if (!settle) return
    headlessExportWaiters.delete(event.sender.id)
    const state = result as { ok?: unknown; error?: unknown } | null
    settle({
      ok: state?.ok === true,
      ...(typeof state?.error === 'string' ? { error: state.error } : {}),
    })
  })

  /// `explicitPath` re-opens a known file without the picker or the shell
  /// queue — the crash-recovery entry point. It runs the SAME pipeline
  /// (prepareWorkbookForOpen + openWorkbookSession + the opened hook) as a
  /// user-picked or shell-queued open, so a recovered workbook is a normal
  /// session rather than a merge source.
  const openSelectedWorkbook = async (event: IpcMainInvokeEvent, explicitPath?: string) => {
    const entry = sessionFor(event)
    let path = explicitPath ?? queuedWorkbookPaths.get(event.sender.id) ?? forcedWorkbookPath
    // consume immediately (before the slow session open) so the shell's
    // retry loop stops re-sending 'open' for the same file
    queuedWorkbookPaths.delete(event.sender.id)
    if (!path) {
      const selection = await openFileDialog(event, {
        properties: ['openFile'],
        filters: [
          { name: tm('filterSpreadsheets'), extensions: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv'] },
        ],
      })
      if (selection.canceled || !selection.filePaths[0]) return null
      path = selection.filePaths[0]
    }
    const prepared = await prepareWorkbookForOpen(
      entry.client,
      path,
      event.sender,
      dialogParent(event),
    )
    // The recovery prompt (or the file dialog / import conversion) can outlive
    // the tab: once the renderer is destroyed, its 'destroyed' handler has
    // already run closeAllSessions and dropped the tab entry, so a session
    // opened now would never be closed and its snapshot would leak.
    if (event.sender.isDestroyed()) {
      await discardPreparedOpen(prepared)
      return null
    }
    const result = await openWorkbookSession(entry.client, prepared.openPath, entry.sessions, {
      suggestSaveAs: prepared.suggestSaveAs,
      csvImport: prepared.csvImport,
      csvSourcePath: prepared.csvSourcePath,
      emptyCsv: prepared.emptyCsv,
      importTempDir: prepared.importTempDir,
      restoreTarget: prepared.restoreTarget,
      emptySource: prepared.emptySource,
      unsavedNew: prepared.unsavedNew,
    })
    // The sidecar open itself can also outlive the tab after the pre-open
    // check. Close the newly registered session instead of stranding it in
    // the detached entry map.
    if (event.sender.isDestroyed()) {
      const session = entry.sessions.get(result.sessionId)
      entry.sessions.delete(result.sessionId)
      if (session !== undefined) {
        await cleanupSessionResources({
          tempRoot: app.getPath('temp'),
          snapshotPath: session.snapshotPath,
          importTempDir: session.importTempDir,
          closeSidecar: () => entry.client.close(result.sessionId),
        })
      }
      return null
    }
    // an unsaved new workbook's tab already shows the untitled name; its temp
    // backing path must not become a recent file
    if (!prepared.unsavedNew) workbookOpenedHook?.(event.sender, path)
    return result
  }

  ipcMain.handle(IPC_CHANNELS.selectWorkbook, async (event) => {
    try {
      return await openSelectedWorkbook(event)
    } catch (err) {
      // A headless export has no dialog to report a failed open through, and
      // its renderer would poll to the deadline waiting for a workbook that
      // will never arrive — settle the run with the real reason instead.
      failHeadlessExport(event.sender.id, `the input workbook did not open (${String(err)})`)
      throw err
    }
  })

  ipcMain.handle(IPC_CHANNELS.reopenWorkbook, async (event, path: unknown) => {
    const validated = z.string().min(1).parse(path)
    // A crash re-open targets the file the user already has open, so a
    // missing/deleted file is a real failure rather than a reason to pop the
    // picker at someone who did not ask to open anything.
    if (!existsSync(validated)) throw new Error('Workbook file not found.')
    const ext = validated.slice(validated.lastIndexOf('.') + 1).toLowerCase()
    if (!WORKBOOK_EXTS.has(ext)) throw new Error(`Unsupported workbook: ${ext}`)
    return openSelectedWorkbook(event, validated)
  })

  // Merge sources: same open pipeline as selectWorkbook, but multi-select,
  // never consuming the shell's queued open path and never retitling the tab
  // (workbookOpenedHook) — these sessions exist only to be read from and
  // closed by the renderer's merge routine.
  /** Open the given spreadsheet paths as merge-source sessions; cleans up
   *  everything already opened when a later file fails or the tab dies. */
  const openMergeSources = async (
    event: Electron.IpcMainInvokeEvent,
    paths: readonly string[],
  ): Promise<unknown[] | null> => {
    const entry = sessionFor(event)
    const opened: { sessionId: string }[] = []
    const closeOpened = async () => {
      for (const { sessionId } of opened) {
        const session = entry.sessions.get(sessionId)
        entry.sessions.delete(sessionId)
        if (session !== undefined) {
          await cleanupSessionResources({
            tempRoot: app.getPath('temp'),
            snapshotPath: session.snapshotPath,
            importTempDir: session.importTempDir,
            closeSidecar: () => entry.client.close(sessionId),
          })
        }
      }
    }
    try {
      for (const path of paths) {
        const prepared = await prepareWorkbookForOpen(
          entry.client,
          path,
          event.sender,
          dialogParent(event),
          { skipRecoveryPrompt: true },
        )
        if (event.sender.isDestroyed()) {
          await discardPreparedOpen(prepared)
          break
        }
        const result = await openWorkbookSession(entry.client, prepared.openPath, entry.sessions, {
          suggestSaveAs: prepared.suggestSaveAs,
          csvImport: prepared.csvImport,
          csvSourcePath: prepared.csvSourcePath,
          importTempDir: prepared.importTempDir,
          restoreTarget: prepared.restoreTarget,
          emptySource: prepared.emptySource,
          unsavedNew: prepared.unsavedNew,
        })
        opened.push(result as { sessionId: string })
        if (event.sender.isDestroyed()) break
      }
    } catch (error) {
      // a later file failing must not strand the sessions already opened
      await closeOpened()
      throw error
    }
    if (event.sender.isDestroyed()) {
      await closeOpened()
      return null
    }
    return opened.length > 0 ? opened : null
  }

  ipcMain.handle(IPC_CHANNELS.selectWorkbooksForMerge, async (event) => {
    const selection = await openFileDialog(event, {
      properties: ['openFile', 'multiSelections'],
      filters: [
        { name: tm('filterSpreadsheets'), extensions: ['xlsx', 'xlsm', 'xls', 'csv', 'tsv'] },
      ],
    })
    if (selection.canceled || selection.filePaths.length === 0) return null
    return openMergeSources(event, selection.filePaths)
  })

  ipcMain.handle(IPC_CHANNELS.readWorkbookRange, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookRangeRequestSchema.parse(input)
    if (!entry.sessions.has(request.sessionId)) throw new Error('Unknown workbook session.')
    const result = await entry.client.readRange(request)
    return workbookRangeResultSchema.parse(result)
  })

  ipcMain.handle(IPC_CHANNELS.readWorkbookFormulas, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookFormulaCellsRequestSchema.parse(input)
    if (!entry.sessions.has(request.sessionId)) throw new Error('Unknown workbook session.')
    const result = await entry.client.readFormulaCells(request)
    return workbookFormulaCellsResultSchema.parse(result)
  })

  // IronCalc recalculation: sheet ids resolve through the session's file
  // sheet names, so the renderer never sees paths and sheets added this
  // session (no file part) fail closed before reaching the engine.
  const sidecarRecalcResultSchema = z
    .object({
      cells: z.array(
        z
          .object({
            sheet: z.string(),
            row: z.number().int().nonnegative(),
            column: z.number().int().nonnegative(),
            formatted: z.string(),
            number: z.number().optional(),
            isError: z.boolean().optional(),
            isFormula: z.boolean(),
          })
          .strict(),
      ),
      cached: z.boolean().optional(),
    })
    .strict()
  ipcMain.handle(IPC_CHANNELS.recalcWorkbook, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookRecalcRequestSchema.parse(input)
    const session = entry.sessions.get(request.sessionId)
    if (!session) throw new Error('Unknown workbook session.')
    const fileSheetName = (sheetId: string): string => {
      const name = session.sheetNames.get(sheetId)
      if (name === undefined) throw new Error(`Unknown sheet for recalculation: ${sheetId}`)
      return name
    }
    const result = sidecarRecalcResultSchema.parse(
      await entry.client.recalcCells({
        // The snapshot, not the live path: recalculated values are painted on
        // the session's grid (and saved into its formula cells), so they must
        // come from the session's own bytes even if the file changed on disk.
        path: session.snapshotPath,
        edits: request.edits.map((edit) => ({
          sheet: fileSheetName(edit.sheetId),
          row: edit.row,
          column: edit.column,
          input: edit.input,
        })),
        reads: request.reads.map((read) => ({
          sheet: fileSheetName(read.sheetId),
          range: read.range,
        })),
      }),
    )
    const idsByName = new Map([...session.sheetNames].map(([id, name]) => [name, id]))
    return workbookRecalcResultSchema.parse({
      cells: result.cells.flatMap((cell) => {
        const sheetId = idsByName.get(cell.sheet)
        if (sheetId === undefined) return []
        return [
          {
            sheetId,
            row: cell.row,
            column: cell.column,
            formatted: cell.formatted,
            ...(cell.number === undefined ? {} : { number: cell.number }),
            ...(cell.isError ? { isError: true } : {}),
            isFormula: cell.isFormula,
          },
        ]
      }),
    })
  })

  ipcMain.handle(IPC_CHANNELS.readWorkbookMedia, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookMediaRequestSchema.parse(input)
    if (!entry.sessions.has(request.sessionId)) throw new Error('Unknown workbook session.')
    const result = await entry.client.readMedia(request)
    return workbookMediaResultSchema.parse(result)
  })

  function sniffImageType(bytes: Buffer): 'image/png' | 'image/jpeg' | 'image/gif' | null {
    if (
      bytes.length >= 8 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47
    )
      return 'image/png'
    if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
      return 'image/jpeg'
    }
    if (bytes.length >= 6 && bytes.subarray(0, 4).toString('latin1') === 'GIF8') return 'image/gif'
    return null
  }

  ipcMain.handle(IPC_CHANNELS.readLocalImage, async (event, input: unknown) => {
    sessionFor(event)
    const request = localImageRequestSchema.parse(input)
    const resolved = request.path.startsWith('~/')
      ? join(app.getPath('home'), request.path.slice(2))
      : request.path
    if (!isAbsolute(resolved)) throw new Error(tm('errImgAbsPath'))
    const info = await stat(resolved).catch(() => null)
    if (!info?.isFile()) throw new Error(tm('errImgNotFound', { path: request.path }))
    if (info.size > 20 * 1024 * 1024) throw new Error(tm('errImgTooLarge20'))
    const bytes = await readFile(resolved)
    const mediaType = sniffImageType(bytes)
    if (mediaType === null) {
      throw new Error(tm('errImgBadType'))
    }
    return localImageResultSchema.parse({ mediaType, base64: bytes.toString('base64') })
  })

  ipcMain.handle(IPC_CHANNELS.captureScreenSources, async (event) => {
    sessionFor(event)
    // macOS gates desktopCapturer behind the Screen Recording permission and
    // returns black frames instead of failing; surface a real denied state.
    if (process.platform === 'darwin') {
      const status = systemPreferences.getMediaAccessStatus('screen')
      if (status !== 'granted' && status !== 'not-determined') {
        return screenSourcesResultSchema.parse({ status: 'denied', sources: [] })
      }
    }
    const sources = await desktopCapturer.getSources({
      types: ['screen', 'window'],
      thumbnailSize: { width: 320, height: 200 },
      fetchWindowIcons: false,
    })
    if (
      process.platform === 'darwin' &&
      systemPreferences.getMediaAccessStatus('screen') !== 'granted'
    ) {
      return screenSourcesResultSchema.parse({ status: 'denied', sources: [] })
    }
    const selfWindow = hostWindowFor(event.sender)
    const selfId = selfWindow?.getMediaSourceId()
    return screenSourcesResultSchema.parse({
      status: 'ok',
      sources: sources
        .filter((source) => source.id !== selfId)
        .map((source) => ({
          id: source.id,
          name: source.name,
          kind: source.id.startsWith('screen') ? 'screen' : 'window',
          thumbnail: source.thumbnail.isEmpty() ? '' : source.thumbnail.toDataURL(),
        })),
    })
  })

  ipcMain.handle(IPC_CHANNELS.captureScreenSource, async (event, input: unknown) => {
    sessionFor(event)
    const request = screenCaptureRequestSchema.parse(input)
    // desktopCapturer only ever returns thumbnails, so a full-res capture is
    // a re-listing with the thumbnail sized to the largest physical display.
    const displays = screen.getAllDisplays()
    const captureSize = {
      width: Math.min(
        4096,
        Math.max(1920, ...displays.map((d) => Math.ceil(d.size.width * d.scaleFactor))),
      ),
      height: Math.min(
        4096,
        Math.max(1080, ...displays.map((d) => Math.ceil(d.size.height * d.scaleFactor))),
      ),
    }
    const sources = await desktopCapturer.getSources({
      types: ['screen', 'window'],
      thumbnailSize: captureSize,
      fetchWindowIcons: false,
    })
    const source = sources.find((candidate) => candidate.id === request.id)
    if (!source || source.thumbnail.isEmpty()) return null
    let image = source.thumbnail
    let png = image.toPNG()
    if (png.length > 20 * 1024 * 1024) {
      image = image.resize({ width: Math.round(image.getSize().width / 2) })
      png = image.toPNG()
    }
    const { width, height } = image.getSize()
    return screenCaptureResultSchema.parse({
      mediaType: 'image/png',
      base64: png.toString('base64'),
      width,
      height,
    })
  })

  ipcMain.handle(IPC_CHANNELS.readPivotDefinition, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookPivotRequestSchema.parse(input)
    const session = entry.sessions.get(request.sessionId)
    if (!session) throw new Error('Unknown workbook session.')
    // Read from the session snapshot so the definition matches what the
    // renderer shows even if the file on disk changed since open.
    const [pivotXml, cacheXml] = await Promise.all([
      readArchiveEntryText(entry.client, session.snapshotPath, request.path),
      readArchiveEntryText(entry.client, session.snapshotPath, request.cachePath),
    ])
    return workbookPivotDefinitionSchema.parse(parsePivotDefinition(pivotXml, cacheXml))
  })

  ipcMain.handle(IPC_CHANNELS.exportPdf, async (event, input: unknown) => {
    sessionFor(event)
    const request = workbookExportPdfRequestSchema.parse(input)
    const result = await exportPdf(event, request)
    if (!result.canceled && result.path) openGeneratedFile(result.path)
    return result
  })

  ipcMain.handle(IPC_CHANNELS.printWorkbook, async (event, input: unknown) => {
    sessionFor(event)
    return printWorkbook(event, workbookExportPdfRequestSchema.parse(input))
  })

  ipcMain.handle(IPC_CHANNELS.exportCsv, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookExportCsvRequestSchema.parse(input)
    const parent = dialogParent(event)
    if (request.hasFormulas) {
      // Excel's CSV warning flow: offer keeping the formulas via .xlsx first.
      const options = {
        type: 'warning' as const,
        message: tm('csvFormulaLossMsg'),
        detail: tm('csvFormulaLossDetail'),
        buttons: [tm('csvKeepXlsxBtn'), tm('csvContinueBtn'), tm('btnCancel')],
        defaultId: 0,
        cancelId: 2,
        noLink: true,
      }
      const { response } = parent
        ? await dialog.showMessageBox(parent, options)
        : await dialog.showMessageBox(options)
      if (response === 0) return { canceled: true, saveAsXlsxInstead: true }
      if (response === 2) return { canceled: true }
    }
    let pickedPath = request.targetPath
    if (pickedPath === undefined) {
      const selection = await saveFileDialog(event, {
        defaultPath: request.fileName,
        filters: [{ name: tm('filterCsv'), extensions: ['csv'] }],
        ...(request.activeSheetName
          ? {
              title: tm('csvActiveSheetOnlyNotice', { name: request.activeSheetName }),
              message: tm('csvActiveSheetOnlyNotice', { name: request.activeSheetName }),
            }
          : {}),
      })
      if (selection.canceled || !selection.filePath) return { canceled: true }
      pickedPath = selection.filePath
    }
    const targetPath = pickedPath.toLowerCase().endsWith('.csv') ? pickedPath : `${pickedPath}.csv`
    // UTF-8 BOM so Excel decodes the reopened file correctly. Written beside
    // the destination and renamed into place: a plain writeFile creates the
    // (empty) file before the data lands, and anything watching for the
    // export — the e2e retry loop included — can read zero bytes in that
    // window.
    const csvBytes = Buffer.concat([
      Buffer.from([0xef, 0xbb, 0xbf]),
      Buffer.from(request.content, 'utf8'),
    ])
    await atomicWriteFile(targetPath, csvBytes)
    // An export can land on a CSV session's own source file — refresh that
    // session's guard digest so its next Save doesn't mistake this write for
    // an external change.
    const writtenSha = await sha256File(targetPath).catch(() => undefined)
    if (writtenSha !== undefined) {
      for (const [sessionId, session] of entry.sessions) {
        if (session.csvSourcePath === targetPath) {
          entry.sessions.set(sessionId, { ...session, csvSourceSha: writtenSha })
        }
      }
    }
    return { canceled: false, path: targetPath }
  })

  // First Save of a CSV session: Excel's "keep this format?" question. The
  // renderer remembers the answer for the file, so it is asked once.
  ipcMain.handle(IPC_CHANNELS.csvSaveConfirm, async (event) => {
    sessionFor(event)
    const options = {
      type: 'warning' as const,
      message: tm('csvKeepFormatMsg'),
      detail: tm('csvKeepFormatDetail'),
      buttons: [tm('csvContinueBtn'), tm('csvKeepXlsxBtn'), tm('btnCancel')],
      defaultId: 0,
      cancelId: 2,
      noLink: true,
    }
    const parent = dialogParent(event)
    const { response } = parent
      ? await dialog.showMessageBox(parent, options)
      : await dialog.showMessageBox(options)
    return response === 0 ? 'csv' : response === 1 ? 'xlsx' : 'cancel'
  })

  ipcMain.handle(IPC_CHANNELS.saveWorkbook, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const client = entry.client
    const request = resolveTransferredEdits(entry, workbookSaveRequestSchema.parse(input))
    const session = entry.sessions.get(request.sessionId)
    if (!session) throw new Error('Unknown workbook session.')

    // A CSV session's plain Save keeps the CSV identity: the xlsx save lands
    // on the temp copy and the serialized csvContent is written back to the
    // original .csv afterwards.
    const csvInPlace = request.mode === 'save' && session.csvSourcePath !== undefined
    let targetPath = session.path
    if (request.mode === 'save' && request.quiet && session.unsavedNew) {
      // AutoSave of a workbook the user has not named yet:
      // no dialog mid-flow, the edits land in the backing temp file and the
      // session stays an unsaved new one (Ctrl+S still asks where to save)
      targetPath = session.path
    } else if (request.mode === 'save-as' || session.suggestSaveAs !== undefined) {
      // .xlsm keeps its extension: untouched archive entries (vbaProject.bin,
      // the macro-enabled content type) round-trip verbatim through the save.
      const macroEnabled = /\.xlsm$/i.test(
        session.suggestSaveAs ?? session.restoreTarget ?? session.path,
      )
      const ext = macroEnabled ? 'xlsm' : 'xlsx'
      const selection = await saveFileDialog(event, {
        defaultPath:
          session.suggestSaveAs ??
          session.csvSourcePath?.replace(/\.[^.]+$/, '.xlsx') ??
          session.restoreTarget ??
          session.path,
        filters: macroEnabled
          ? [{ name: tm('filterXlsm'), extensions: ['xlsm'] }]
          : [
              { name: tm('filterXlsx'), extensions: ['xlsx'] },
              { name: tm('filterCsv'), extensions: ['csv'] },
            ],
        // CSV import: explain why the save goes through .xlsx (CSV keeps values only)
        ...(session.csvImport
          ? { title: tm('csvSaveAsNotice'), message: tm('csvSaveAsNotice') }
          : {}),
      })
      if (selection.canceled || !selection.filePath) return { canceled: true }
      // A CSV pick can't ride the xlsx pipeline: hand the path back so the
      // renderer serializes the active sheet through the CSV export channel.
      if (!macroEnabled && selection.filePath.toLowerCase().endsWith('.csv')) {
        return { canceled: true, csvSaveAsPath: selection.filePath }
      }
      targetPath = selection.filePath.toLowerCase().endsWith(`.${ext}`)
        ? selection.filePath
        : `${selection.filePath}.${ext}`
    } else if (session.restoreTarget !== undefined) {
      // Restored crash-recovery copy: the restore prompt was the confirmation,
      // so Save writes straight back to the original — unless someone else
      // changed it since the restore.
      const currentSha = await sha256File(session.restoreTarget).catch(() => undefined)
      if (currentSha !== undefined && currentSha !== session.restoreTargetSha) {
        throw new Error(tm('errDiskChanged'))
      }
      targetPath = session.restoreTarget
    } else {
      // Plain in-place save: refuse to silently overwrite a file some other
      // program changed after this session opened it. Save As (above) skips
      // this guard on purpose — it patches the session snapshot, not the live
      // file, and writes to a path the user just confirmed, so it stays
      // usable as the escape hatch this error message points to. A file that
      // was deleted on disk is fine: saving recreates it.
      const currentSha = await sha256File(session.path).catch(() => undefined)
      if (currentSha !== undefined && currentSha !== session.sha256) {
        throw new Error(tm('errDiskChanged'))
      }
    }
    // The CSV write-back gets the same external-change guard as restoreTarget;
    // a deleted .csv is fine — the write recreates it.
    if (csvInPlace && session.csvSourcePath !== undefined) {
      const csvSha = await sha256File(session.csvSourcePath).catch(() => undefined)
      if (csvSha !== undefined && csvSha !== session.csvSourceSha) {
        throw new Error(tm('errDiskChanged'))
      }
    }

    const mutation = await writeWorkbookTo(client, session, request, targetPath)

    if (csvInPlace && session.csvSourcePath !== undefined && request.csvContent !== undefined) {
      // The temp copy already holds the saved bytes: refresh the session's
      // guard digest first, so a failed CSV write-back below leaves a
      // retryable session instead of stranding the next Save on
      // errDiskChanged against its own write.
      const savedSha = await sha256File(session.path).catch(() => undefined)
      if (savedSha !== undefined && entry.sessions.has(request.sessionId)) {
        entry.sessions.set(request.sessionId, { ...session, sha256: savedSha })
      }
      // UTF-8 BOM so Excel decodes the reopened file correctly.
      await writeFile(
        session.csvSourcePath,
        Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(request.csvContent, 'utf8')]),
      )
    }

    // The sidecar session still streams the pre-save bytes; swap it for a
    // fresh session over the saved file so future reads match the disk state.
    const quietUnsavedNew = targetPath === session.path && session.unsavedNew === true
    entry.sessions.delete(request.sessionId)
    await cleanupSessionResources({
      tempRoot: app.getPath('temp'),
      snapshotPath: session.snapshotPath,
      // A CSV in-place save keeps saving into the temp copy, and a quiet save of
      // an unsaved new workbook just wrote its backing file — the directory
      // must survive the session swap.
      importTempDir: csvInPlace || quietUnsavedNew ? undefined : session.importTempDir,
      closeSidecar: () => client.close(request.sessionId),
    })
    const file = await openWorkbookSession(
      client,
      targetPath,
      entry.sessions,
      csvInPlace
        ? {
            csvImport: true,
            csvSourcePath: session.csvSourcePath,
            importTempDir: session.importTempDir,
          }
        : quietUnsavedNew
          ? {
              suggestSaveAs: session.suggestSaveAs,
              importTempDir: session.importTempDir,
              unsavedNew: true,
            }
          : undefined,
    )
    if (quietUnsavedNew && session.suggestSaveAs !== undefined) {
      // the temp path must not reach the title or recents; the recovery copy
      // is what a crash restores from, so it tracks the saved bytes
      try {
        await mkdir(recoveryDir(), { recursive: true })
        await copyFile(targetPath, recoveryPathFor(session.suggestSaveAs))
        rememberUnsavedNewRecovery(session.suggestSaveAs)
      } catch (error) {
        console.warn('[sheets] recovery copy of an unsaved new workbook failed:', error)
      }
      return { canceled: false, file, touchedEntries: mutation.touchedEntries }
    }
    // Notify shell (if running) so it can update the tab title and record the
    // saved path in recent files (mirrors the open hook; covers Save As + first
    // save after converting an .xls/.csv import). A CSV session's user-visible
    // file is the original .csv, not the temp copy the xlsx save landed on.
    workbookOpenedHook?.(
      event.sender,
      (csvInPlace ? session.csvSourcePath : undefined) ?? targetPath,
    )
    // The file on disk now carries these edits
    clearWorkbookRecovery(targetPath)
    if (session.suggestSaveAs !== undefined) {
      if (session.unsavedNew) forgetUnsavedNewRecovery(session.suggestSaveAs)
      else clearWorkbookRecovery(session.suggestSaveAs)
    }
    // Restored session saved (possibly Save As elsewhere): the unsaved work is
    // persisted, so the original's recovery copy must not re-offer it.
    if (session.restoreTarget !== undefined) clearWorkbookRecovery(session.restoreTarget)
    return { canceled: false, file, touchedEntries: mutation.touchedEntries }
  })

  // Chunked upload for edit sets too large to inline in one save request:
  // the renderer opens a transfer, streams ordered slices, then references
  // the transfer id from the save (or recovery) request that follows.
  ipcMain.handle(IPC_CHANNELS.saveEditsBegin, (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookSaveEditsBeginSchema.parse(input)
    if (!entry.sessions.has(request.sessionId)) throw new Error('Unknown workbook session.')
    entry.saveTransfers.begin(request)
  })

  ipcMain.handle(IPC_CHANNELS.saveEditsChunk, (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookSaveEditsChunkSchema.parse(input)
    // The chunk crosses the bridge and the IPC hop as a flat JSON string;
    // the edits stay untrusted input until they pass the cell-edit schema.
    entry.saveTransfers.addChunk({
      sessionId: request.sessionId,
      transferId: request.transferId,
      seq: request.seq,
      edits: saveEditsChunkArraySchema.parse(JSON.parse(request.editsJson)),
    })
  })

  // Best-effort cleanup from renderer failure paths; a no-op if the transfer
  // was already consumed or expired.
  ipcMain.handle(IPC_CHANNELS.saveEditsAbort, (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = workbookSaveEditsAbortSchema.parse(input)
    entry.saveTransfers.discard(request.transferId, request.sessionId)
  })

  // Crash-recovery copy of a dirty workbook: the same save pipeline with a
  // userData target, no session swap and no dialogs — best-effort, silent on failure.
  ipcMain.handle(IPC_CHANNELS.writeWorkbookRecovery, async (event, input: unknown) => {
    const entry = sessionFor(event)
    const request = resolveTransferredEdits(entry, workbookSaveRequestSchema.parse(input))
    const session = entry.sessions.get(request.sessionId)
    // A converted import has no original file to recover into (and a CSV
    // session's original can't hold the workbook bytes); a restored recovery
    // session is backed by the recovery copy itself — writing over the file
    // the sidecar streams from would corrupt the open session.
    if (
      !session ||
      (session.suggestSaveAs !== undefined && !session.unsavedNew) ||
      session.csvSourcePath !== undefined ||
      (session.restoreTarget !== undefined && !session.emptySource)
    )
      return { ok: false }
    if (session.automaticRecoveryDisabled) return { ok: false }
    // an unsaved new workbook recovers under the name its first save would get;
    // an empty original's session runs on a temp blank and recovers into the file the user opened
    const recoveryKey = session.unsavedNew
      ? session.suggestSaveAs!
      : session.emptySource
        ? (session.restoreTarget ?? session.path)
        : session.path
    try {
      await mkdir(recoveryDir(), { recursive: true })
      await writeWorkbookTo(entry.client, session, request, recoveryPathFor(recoveryKey))
      if (session.unsavedNew) rememberUnsavedNewRecovery(recoveryKey)
      return { ok: true }
    } catch (error) {
      console.warn('[sheets] recovery copy failed:', error)
      return { ok: false }
    }
  })

  ipcMain.handle(IPC_CHANNELS.closeWorkbook, async (event, sessionId: unknown) => {
    const entry = sessionFor(event)
    const validatedSessionId = z.string().uuid().parse(sessionId)
    entry.saveTransfers.discardSession(validatedSessionId)
    const session = entry.sessions.get(validatedSessionId)
    if (!entry.sessions.delete(validatedSessionId)) return
    if (session === undefined) return
    if (session.unsavedNew && session.suggestSaveAs !== undefined)
      forgetUnsavedNewRecovery(session.suggestSaveAs)
    await cleanupSessionResources({
      tempRoot: app.getPath('temp'),
      snapshotPath: session.snapshotPath,
      importTempDir: session.importTempDir,
      closeSidecar: () => entry.client.close(validatedSessionId),
    })
  })

  ipcMain.handle(IPC_CHANNELS.openExternal, async (event, url: unknown) => {
    sessionFor(event)
    const validatedUrl = safeExternalUrl(url)
    if (!validatedUrl) {
      throw new Error('Only http(s) links can be opened.')
    }
    await shell.openExternal(validatedUrl)
  })
}

/**
 * Resolve a save request's sheet ops / name mappings and write the workbook through
 * the sidecar. Split out of the save handler so a crash-recovery copy can reuse the
 * exact same pipeline with a different targetPath.
 */
async function writeWorkbookTo(
  client: XlsxSidecarClient,
  session: SessionInfo,
  request: WorkbookSaveRequest,
  targetPath: string,
): Promise<Awaited<ReturnType<typeof saveWorkbookViaSidecar>>> {
  // Sheet ops resolve first: added sheets have Univer ids the session map
  // doesn't know, so cell edits into them resolve through the op's name.
  const addedSheetNames = new Map<string, string>()
  // Added sheet id → file name of the sheet whose part seeds the new part.
  const duplicateSources = new Map<string, string>()
  const renames: { sheetName: string; newName: string }[] = []
  const removals: string[] = []
  const hiddenChanges: { sheetName: string; hidden: boolean }[] = []
  let orderChanged = false
  for (const op of request.sheetOps) {
    if (op.kind === 'add-sheet') {
      addedSheetNames.set(op.sheetId, op.name)
      continue
    }
    if (op.kind === 'duplicate-sheet') {
      // The renderer resolves duplicate chains to a sheet the file knows,
      // so the source must be in the session map.
      const sourceName = session.sheetNames.get(op.sourceSheetId)
      if (!sourceName) throw new Error(`Unknown duplicate source ${op.sourceSheetId}.`)
      addedSheetNames.set(op.sheetId, op.name)
      duplicateSources.set(op.sheetId, sourceName)
      continue
    }
    if (op.kind === 'reorder-sheets') {
      orderChanged = true
      continue
    }
    const sheetName = addedSheetNames.get(op.sheetId) ?? session.sheetNames.get(op.sheetId)
    if (!sheetName) throw new Error(`Unknown worksheet ${op.sheetId}.`)
    if (op.kind === 'rename-sheet') renames.push({ sheetName, newName: op.newName })
    else if (op.kind === 'set-sheet-hidden') {
      hiddenChanges.push({ sheetName, hidden: op.hidden })
    } else removals.push(sheetName)
  }
  const renameByOriginal = new Map(renames.map((rename) => [rename.sheetName, rename.newName]))
  const resolveSheetName = (sheetId: string): string => {
    const sheetName = addedSheetNames.get(sheetId) ?? session.sheetNames.get(sheetId)
    if (!sheetName) throw new Error(`Unknown worksheet ${sheetId}.`)
    return sheetName
  }
  let sheetPlan: SheetEditPlan | undefined
  if (request.sheetOps.length > 0) {
    sheetPlan = {
      renames,
      additions: [...addedSheetNames].map(([sheetId, name]) => ({
        name,
        sourceSheetName: duplicateSources.get(sheetId),
      })),
      removals,
      hiddenChanges,
      orderChanged,
      order: request.sheetOrder.map((sheetId) => {
        const original = resolveSheetName(sheetId)
        return addedSheetNames.has(sheetId)
          ? original
          : (renameByOriginal.get(original) ?? original)
      }),
    }
  }

  const edits: CellEdit[] = request.edits.map((edit) => ({
    sheetName: resolveSheetName(edit.sheetId),
    row: edit.row,
    column: edit.column,
    writeValue: edit.writeValue,
    cell: { value: edit.value, formula: edit.formula },
    style: edit.style,
    rich: edit.rich,
    styleReset: edit.styleReset,
  }))
  const bulkConstantFills = (request.bulkConstantFills ?? []).map(({ sheetId, ...fill }) => ({
    sheetName: resolveSheetName(sheetId),
    ...fill,
  }))
  const opsBySheet = new Map<string, SheetStructuralOps['ops'][number][]>()
  for (const op of request.structuralOps) {
    const sheetName = resolveSheetName(op.sheetId)
    const sheetOps = opsBySheet.get(sheetName) ?? []
    if ('range' in op) {
      sheetOps.push({ kind: op.kind, range: op.range })
    } else if ('size' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, size: op.size })
    } else if ('level' in op) {
      sheetOps.push({
        kind: op.kind,
        start: op.start,
        end: op.end,
        level: op.level,
        ...(op.collapsed === undefined ? {} : { collapsed: op.collapsed }),
      })
    } else if ('hidden' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, hidden: op.hidden })
    } else if ('style' in op) {
      sheetOps.push({ kind: op.kind, start: op.start, end: op.end, style: op.style })
    } else if ('before' in op) {
      sheetOps.push({ kind: op.kind, index: op.index, count: op.count, before: op.before })
    } else {
      sheetOps.push({ kind: op.kind, index: op.index, count: op.count })
    }
    opsBySheet.set(sheetName, sheetOps)
  }
  const structuralOps: SheetStructuralOps[] = [...opsBySheet].map(([sheetName, ops]) => ({
    sheetName,
    ops,
  }))
  const filterStates = request.filterStates.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    filter: state.filter,
    hiddenRows: state.hiddenRows,
    visibilityRange: state.visibilityRange,
  }))
  const linksBySheet = new Map<string, { row: number; column: number; target: string | null }[]>()
  for (const link of request.hyperlinkEdits) {
    const sheetName = resolveSheetName(link.sheetId)
    const sheetLinks = linksBySheet.get(sheetName) ?? []
    sheetLinks.push({ row: link.row, column: link.column, target: link.target })
    linksBySheet.set(sheetName, sheetLinks)
  }
  const hyperlinkEdits = [...linksBySheet].map(([sheetName, links]) => ({
    sheetName,
    edits: links,
  }))
  const cfStates = request.cfStates.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    rules: state.rules,
  }))
  const dvStates = request.dvStates.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    rules: state.rules,
  }))
  const sheetProtections = request.sheetProtections.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    protected: state.protected,
  }))
  const protectedRangeStates = request.protectedRangeStates.map((state) => ({
    sheetName: resolveSheetName(state.sheetId),
    ranges: state.ranges,
  }))
  const pageSetupStates = request.pageSetupStates.map(({ sheetId, ...state }) => ({
    sheetName: resolveSheetName(sheetId),
    ...state,
  }))
  const noteStates = request.noteStates.map(({ sheetId, notes }) => ({
    sheetName: resolveSheetName(sheetId),
    notes,
  }))
  const visualAdditions = request.visualAdditions.map((addition) => ({
    sheetName: resolveSheetName(addition.sheetId),
    anchor: addition.anchor,
    chart: addition.chart,
    shape: addition.shape,
    image: addition.image,
  }))
  const tableAdditions = request.tableAdditions.map((table) => ({
    sheetName: resolveSheetName(table.sheetId),
    area: table.area,
    name: table.name,
    columnNames: table.columnNames,
    style: table.style,
    bandedRows: table.bandedRows,
  }))
  const pivotAdditions = request.pivotAdditions.map((pivot) => ({
    sheetName: resolveSheetName(pivot.sheetId),
    sourceSheetName: resolveSheetName(pivot.sourceSheetId),
    sourceArea: pivot.sourceArea,
    location: pivot.location,
    name: pivot.name,
    fieldNames: pivot.fieldNames,
    rowFieldIndices: pivot.rowFieldIndices,
    columnFieldIndex: pivot.columnFieldIndex,
    pageFieldIndices: pivot.pageFieldIndices,
    rowItems: pivot.rowItems,
    rowLevelItems: pivot.rowLevelItems,
    rowLines: pivot.rowLines,
    columnItems: pivot.columnItems,
    columnFieldIndices: pivot.columnFieldIndices,
    colLevelItems: pivot.colLevelItems,
    colLines: pivot.colLines,
    groupings: pivot.groupings,
    filters: pivot.filters,
    rowHiddenItems: pivot.rowHiddenItems,
    colHiddenItems: pivot.colHiddenItems,
    values: pivot.values,
  }))
  const sparklineAdditions = request.sparklineAdditions.map(({ sheetId, ...group }) => ({
    sheetName: resolveSheetName(sheetId),
    ...group,
  }))
  // Recalculated formula values: sheetId → file sheet name, the same
  // resolution the cell edits use.
  const formulaValuesBySheet = new Map<
    string,
    { row: number; column: number; value: string | number | boolean | null | { error: string } }[]
  >()
  for (const cell of request.formulaValues) {
    const sheetName = resolveSheetName(cell.sheetId)
    const list = formulaValuesBySheet.get(sheetName) ?? []
    list.push({ row: cell.row, column: cell.column, value: cell.value })
    formulaValuesBySheet.set(sheetName, list)
  }
  const formulaValues = [...formulaValuesBySheet].map(([sheetName, cells]) => ({
    sheetName,
    cells,
  }))
  const mutation = await saveWorkbookViaSidecar({
    client,
    // The snapshot, not the live path: the save base must be the bytes this
    // session's pending edits were made against, regardless of what other
    // programs did to the file since.
    sourcePath: session.snapshotPath,
    targetPath,
    edits,
    bulkConstantFills,
    structuralOps,
    chartEdits: request.chartEdits,
    // Located by package-absolute drawingPath, so no sheet-name mapping.
    visualEdits: request.visualEdits,
    sheetPlan,
    filterStates,
    hyperlinkEdits,
    cfStates,
    dvStates,
    sheetProtections,
    definedNamesState: request.definedNamesState,
    themeState: request.themeState,
    workbookProtectionState: request.workbookProtectionState,
    protectedRangeStates,
    visualAdditions,
    pageSetupStates,
    noteStates,
    tableAdditions,
    pivotAdditions,
    sparklineAdditions,
    formulaValues,
    pivotCacheRefreshPaths: request.pivotCacheRefreshPaths,
    // Output-area expansion from layout growth: sheetId → sheet name; the part
    // path is resolved by the gateway.
    pivotRefreshUpdates: request.pivotRefreshUpdates.map((update) => ({
      cachePath: update.cachePath,
      sheetName: resolveSheetName(update.sheetId),
      newOutputRef: update.newOutputRef,
      ...(update.relayout === undefined
        ? {}
        : {
            relayout: (({ sheetId: _sheetId, sourceSheetId, ...rest }) => ({
              ...rest,
              sourceSheetName: resolveSheetName(sourceSheetId),
            }))(update.relayout),
          }),
    })),
  })
  return mutation
}

/** Copies the workbook into the temp snapshot dir; the copy is the session's
 * save base (see SessionInfo.snapshotPath). */
async function snapshotWorkbook(path: string): Promise<string> {
  const dir = join(app.getPath('temp'), 'genoffice-sheets-sessions')
  await mkdir(dir, { recursive: true })
  const snapshotPath = join(dir, `${randomUUID()}.xlsx`)
  await copyFile(path, snapshotPath)
  return snapshotPath
}

let cachedShortDate: string | undefined

/// Derived from the OS region (not the UI language) and shared with the
/// gateway's save-side numFmtId mapping via setSystemShortDate.
function systemShortDate(): string {
  if (cachedShortDate === undefined) {
    cachedShortDate = shortDatePatternForSystemLocale(app.getSystemLocale())
    setSystemShortDate(cachedShortDate)
  }
  return cachedShortDate
}

/** the tab died between prepare and open: release the temp copy and any name reserved for it */
async function discardPreparedOpen(prepared: {
  openPath: string
  importTempDir?: string | undefined
}): Promise<void> {
  unsavedNewWorkbooks.delete(prepared.openPath)
  if (prepared.importTempDir !== undefined)
    await cleanupImportTempDirectory(app.getPath('temp'), prepared.importTempDir)
}

async function openWorkbookSession(
  client: XlsxSidecarClient,
  path: string,
  sessions: Map<string, SessionInfo>,
  options?: {
    suggestSaveAs?: string | undefined
    csvImport?: boolean | undefined
    csvSourcePath?: string | undefined
    emptyCsv?: boolean | undefined
    importTempDir?: string | undefined
    restoreTarget?: string | undefined
    emptySource?: boolean | undefined
    unsavedNew?: boolean | undefined
  },
): Promise<WorkbookFile> {
  const {
    suggestSaveAs,
    csvImport,
    csvSourcePath,
    emptyCsv,
    importTempDir,
    restoreTarget,
    emptySource,
    unsavedNew,
  } = options ?? {}
  // Snapshot first, then the sidecar opens the snapshot (not the live path):
  // everything the session serves — cell reads, media, recalc, saves — comes
  // from the same bytes, even if the file on disk changes right after the
  // copy. The digest also describes exactly those bytes.
  const snapshotPath = await snapshotWorkbook(path)
  try {
    const [opened, digest, snapshotStat, restoreTargetSha, csvSourceSha] = await Promise.all([
      client
        .open(snapshotPath, getUiLang(), systemShortDate())
        .then((result) => sidecarOpenResultSchema.parse(result)),
      sha256File(snapshotPath),
      stat(snapshotPath),
      // Missing original (deleted since the crash) is fine: the write-back recreates it.
      restoreTarget === undefined
        ? Promise.resolve(undefined)
        : sha256File(restoreTarget).catch(() => undefined),
      csvSourcePath === undefined
        ? Promise.resolve(undefined)
        : sha256File(csvSourcePath).catch(() => undefined),
    ])
    sessions.set(opened.sessionId, {
      path,
      snapshotPath,
      sha256: digest,
      sheetNames: new Map(opened.sheets.map((sheet) => [sheet.id, sheet.name])),
      automaticRecoveryDisabled: !allowsAutomaticWorkbookRecovery(opened.sheets),
      ...(suggestSaveAs === undefined ? {} : { suggestSaveAs }),
      ...(csvImport ? { csvImport } : {}),
      ...(csvSourcePath === undefined ? {} : { csvSourcePath }),
      ...(csvSourceSha === undefined ? {} : { csvSourceSha }),
      ...(importTempDir === undefined ? {} : { importTempDir }),
      ...(restoreTarget === undefined ? {} : { restoreTarget }),
      ...(restoreTargetSha === undefined ? {} : { restoreTargetSha }),
      ...(emptySource ? { emptySource } : {}),
      ...(unsavedNew ? { unsavedNew } : {}),
    })
    unsavedNewWorkbooks.delete(path)
    return workbookFileSchema.parse({
      ...opened,
      // The renderer-facing path is what the user opened: for a restored
      // recovery copy that is the original file, not the copy under userData.
      path: restoreTarget ?? path,
      sha256: digest,
      fileBytes: snapshotStat.size,
      readOnly: false,
      needsSaveAs: suggestSaveAs !== undefined,
      ...(unsavedNew ? { unsavedNew: true } : {}),
      ...(csvSourcePath === undefined ? {} : { csvPath: csvSourcePath }),
      ...(emptyCsv ? { emptyCsv: true } : {}),
      restoredFromRecovery: restoreTarget !== undefined,
      automaticRecoveryDisabled: !allowsAutomaticWorkbookRecovery(opened.sheets),
    })
  } catch (error) {
    await rm(snapshotPath, { force: true }).catch(() => undefined)
    unsavedNewWorkbooks.delete(path)
    if (importTempDir !== undefined) {
      await cleanupImportTempDirectory(app.getPath('temp'), importTempDir)
    }
    throw error
  }
}

/// A 0-byte .xlsx is an empty workbook, not a corrupt one. The sidecar needs a
/// real archive, so a blank one opens from the temp dir with the original as
/// the write-back target (the restore path: Save writes it even with no edits).
async function openEmptyXlsx(path: string): Promise<{
  openPath: string
  importTempDir: string
  restoreTarget: string
  emptySource: true
}> {
  const directory = join(app.getPath('temp'), 'genoffice-imports', randomUUID())
  await mkdir(directory, { recursive: true })
  const openPath = join(directory, basename(path))
  try {
    await writeFile(openPath, (await csvToXlsxBufferForOpen('', 'Sheet1')).buffer)
  } catch (error) {
    await cleanupImportTempDirectory(app.getPath('temp'), directory)
    throw error
  }
  return { openPath, importTempDir: directory, restoreTarget: path, emptySource: true }
}

/** which legacy charset an Excel CSV most likely uses, judged by the UI language */
function legacyCsvCharset(): string | undefined {
  const byLang: Partial<Record<Lang, string>> = {
    zh: 'gb18030',
    'zh-TW': 'big5',
    ja: 'shift_jis',
    ko: 'euc-kr',
  }
  return byLang[getUiLang()]
}

/// .xls and .tsv open as a converted copy in the temp dir; the session
/// remembers the original's .xlsx sibling as the Save As default. .csv converts
/// the same way but keeps its file identity, so Save writes values back to it.
async function prepareWorkbookForOpen(
  client: XlsxSidecarClient,
  path: string,
  contents?: WebContents | undefined,
  parent?: BrowserWindow | undefined,
  options?: { skipRecoveryPrompt?: boolean },
): Promise<{
  openPath: string
  suggestSaveAs?: string
  csvImport?: boolean
  csvSourcePath?: string
  emptyCsv?: boolean
  importTempDir?: string
  restoreTarget?: string
  emptySource?: boolean
  unsavedNew?: boolean
}> {
  // A shell-created "New spreadsheet" opens from a temp directory and has no
  // user-visible file yet: Save As produces it, and closing unsaved discards
  // the directory. The mark is consumed once the session exists (see
  // openWorkbookSession) so the name stays reserved through the open and a
  // reopened path is a normal file.
  const unsavedNew = unsavedNewWorkbooks.get(path)
  if (unsavedNew !== undefined) {
    const base = {
      suggestSaveAs: unsavedNew.suggestSaveAs,
      importTempDir: unsavedNew.tempDir,
      unsavedNew: true,
    }
    // a crash left work under this would-be name: offer it like a file's
    // recovery copy, opening the copy itself as the (still unsaved) workbook
    const recovery = pendingUnsavedNewRecoveryFor(unsavedNew.suggestSaveAs)
    if (recovery) {
      const choice =
        contents && !contents.isDestroyed()
          ? await promptRecoveryRestore(contents, unsavedNew.suggestSaveAs, recovery)
          : await promptRecoveryRestoreNative(parent)
      // the copy is what the recovery writer keeps overwriting, so the session
      // streams from the backing file, not from the copy
      if (choice === 'restore') {
        try {
          await copyFile(recovery, path)
        } catch (error) {
          await discardPreparedOpen({ openPath: path, importTempDir: unsavedNew.tempDir })
          throw error
        }
      } else if (choice === 'discard') forgetUnsavedNewRecovery(unsavedNew.suggestSaveAs)
    }
    return { openPath: path, ...base }
  }
  const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase()
  if (extension !== 'csv' && extension !== 'tsv' && extension !== 'xls') {
    // Unsaved work from a lost session: offer the recovery copy. Restoring
    // opens it with restoreTarget pointing back at the original, so a plain
    // Save writes straight back over the file the user opened — the restore
    // prompt (which spells out the overwrite) was the confirmation.
    // Merge sources are read-only picks: never surface (or worse, discard)
    // another workbook's crash recovery from here.
    const recovery = options?.skipRecoveryPrompt ? undefined : pendingRecoveryFor(path)
    if (recovery) {
      const choice =
        contents && !contents.isDestroyed()
          ? await promptRecoveryRestore(contents, path, recovery)
          : await promptRecoveryRestoreNative(parent)
      if (choice === 'restore') return { openPath: recovery, restoreTarget: path }
      if (choice === 'discard') clearWorkbookRecovery(path)
    }
    if (extension === 'xlsx' && (await stat(path)).size === 0) return openEmptyXlsx(path)
    return { openPath: path }
  }
  const stem = basename(path).replace(/\.[^.]+$/, '')
  const directory = join(app.getPath('temp'), 'genoffice-imports', randomUUID())
  await mkdir(directory, { recursive: true })
  const openPath = join(directory, `${stem}.xlsx`)
  let emptyCsv = false
  try {
    if (extension === 'csv' || extension === 'tsv') {
      const csvStat = await stat(path)
      if (csvStat.size > MAX_DELIMITED_IMPORT_BYTES) throw new Error(tm('errFileTooLarge'))
      const converted = await csvToXlsxBufferForOpen(
        decodeCsvBuffer(await readFile(path), legacyCsvCharset()),
        'Sheet1',
        // A .tsv's delimiter is declared by its extension, and sniffing by
        // frequency can get it wrong: annotation-heavy exports (gene
        // descriptions, database cross-references) hold comma-separated
        // lists, so a narrow table ends up with more commas than tabs and
        // every row shatters on the comma. A .csv keeps the sniff — the prose
        // guard in resolveImportDelimiter is what that path needs.
        extension === 'tsv' ? '\t' : undefined,
      )
      emptyCsv = converted.empty
      await writeFile(openPath, converted.buffer)
    } else {
      await client.convertWorkbook({ path, targetPath: openPath })
    }
  } catch (error) {
    await cleanupImportTempDirectory(app.getPath('temp'), directory)
    throw error
  }
  // CSV keeps its file identity: Save writes the values back to the original
  // .csv (Excel's behavior), so no Save As detour is suggested. Legacy .xls and
  // view-only .tsv route the first save through Save As to a fresh .xlsx —
  // writing values back to a .tsv would need a tab serializer this path lacks.
  return extension === 'csv'
    ? {
        openPath,
        importTempDir: directory,
        csvImport: true,
        csvSourcePath: path,
        ...(emptyCsv ? { emptyCsv: true } : {}),
      }
    : { openPath, importTempDir: directory, suggestSaveAs: path.replace(/\.[^.]+$/, '.xlsx') }
}

/** shell-injected items appended to the File menu (e.g. Back to Home) */
let extraFileMenuItems: MenuItemConstructorOptions[] = []

export function setSheetsExtraFileMenuItems(items: MenuItemConstructorOptions[]): void {
  extraFileMenuItems = items
}

/** tab mode: closes the sheets tab instead of the whole shell window (Cmd+W / role:'close') */
let closeActiveTabHook: (() => void) | null = null
export function setSheetsCloseTabHook(fn: (() => void) | null): void {
  closeActiveTabHook = fn
}

/// The ribbon has no File tab; file commands live in
/// the application menu and are forwarded to the renderer.
function installApplicationMenu(): void {
  const sendMenuAction = sendSheetsMenuAction
  const labels = appMenuLabels(getUiLang())
  Menu.setApplicationMenu(
    Menu.buildFromTemplate([
      ...(process.platform === 'darwin' ? [{ role: 'appMenu' as const }] : []),
      {
        label: tm('menuFile'),
        submenu: [
          {
            label: tm('menuOpenWorkbook'),
            accelerator: 'CmdOrCtrl+O',
            click: () => sendMenuAction('open'),
          },
          ...(extraFileMenuItems.length > 0
            ? [{ type: 'separator' as const }, ...extraFileMenuItems]
            : []),
          { type: 'separator' },
          {
            label: tm('menuSave'),
            accelerator: 'CmdOrCtrl+S',
            click: () => sendMenuAction('save'),
          },
          {
            label: tm('menuSaveAs'),
            accelerator: 'Shift+CmdOrCtrl+S',
            click: () => sendMenuAction('save-as'),
          },
          { type: 'separator' },
          {
            label: tm('menuPrint'),
            accelerator: 'CmdOrCtrl+P',
            click: () => sendMenuAction('print'),
          },
          { type: 'separator' },
          {
            label: tm('menuExportPdf'),
            click: () => sendMenuAction('export-pdf'),
          },
          {
            label: tm('menuExportCsv'),
            click: () => sendMenuAction('export-csv'),
          },
          { type: 'separator' },
          closeActiveTabHook
            ? {
                label: process.platform === 'darwin' ? tm('menuClose') : tm('menuQuit'),
                accelerator: process.platform === 'darwin' ? 'CmdOrCtrl+W' : 'CmdOrCtrl+Q',
                click: () => closeActiveTabHook?.(),
              }
            : process.platform === 'darwin'
              ? { role: 'close' as const, label: tm('menuClose') }
              : { role: 'quit' as const, label: tm('menuQuit') },
        ],
      },
      {
        label: tm('menuEdit'),
        submenu: [
          // role: 'editMenu' would bind ⌘Z to webContents.undo(), a text-editing
          // no-op that starves Univer of the shortcut — forward it instead.
          {
            label: tm('menuUndo'),
            accelerator: 'CmdOrCtrl+Z',
            click: () => sendMenuAction('undo'),
          },
          {
            label: tm('menuRedo'),
            accelerator: 'Shift+CmdOrCtrl+Z',
            click: () => sendMenuAction('redo'),
          },
          { type: 'separator' },
          { role: 'cut', label: labels.cut },
          { role: 'copy', label: labels.copy },
          { role: 'paste', label: labels.paste },
          { type: 'separator' },
          { role: 'selectAll', label: labels.selectAll },
        ],
      },
      // ⌘R is Fill Right and ⌘0/⌘+/⌘- hide columns / zoom the sheet; menu
      // accelerators would swallow them on macOS.
      viewMenuTemplate(labels, { devItems: !app.isPackaged, pageZoom: false }),
      windowMenuTemplate(process.platform, labels),
      helpMenuTemplate(labels),
    ]),
  )
}

/** stop the Rust sidecar (shell calls this from its own before-quit hook) */
export function stopSheetsSidecar(): void {
  sidecar?.stop()
  sidecar = null
}

export {
  installApplicationMenu as installSheetsMenu,
  startCaptureServer as startSheetsCaptureServer,
}

export function startSheetsStandalone(): void {
  registerRendererScheme()
  installNavigationGuard(app)
  installContextMenu(app, () => contextMenuLabels(getUiLang()))
  // GENOFFICE_USER_DATA: test drivers point this at a scratch dir so automated
  // instances get their own userData AND single-instance lock (the lock is scoped
  // to userData), allowing parallel instances alongside a normal dev run.
  // Same dev-only hook as apps/slides/src/main/slides-main.ts.
  if (!app.isPackaged && process.env.GENOFFICE_USER_DATA) {
    app.setPath('userData', process.env.GENOFFICE_USER_DATA)
  }
  app.whenReady().then(() => {
    installRendererProtocol({ sheets: join(__dirname, '../renderer') })
    setUiLang(normalizeLang(process.env.GENOFFICE_LANG ?? app.getLocale()))
    app.setAccessibilitySupportEnabled(true)
    installApplicationMenu()
    startCaptureServer()
    return createSheetsWindow()
  })
  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit()
  })
  app.on('before-quit', () => {
    stopSheetsSidecar()
  })
  app.on('activate', () => {
    if (!mainWindow) void createSheetsWindow()
  })
}

function resolveSidecarPath(): string {
  const executable = process.platform === 'win32' ? 'xlsx-sidecar.exe' : 'xlsx-sidecar'
  if (runtime.sidecarPath) return runtime.sidecarPath
  if (process.env.XLSX_SIDECAR_PATH) return process.env.XLSX_SIDECAR_PATH
  if (app.isPackaged) return join(process.resourcesPath, 'native', executable)
  return join(app.getAppPath(), 'native', 'xlsx-engine', 'target', 'release', executable)
}

async function sha256File(path: string): Promise<string> {
  return new Promise((resolve, reject) => {
    const hash = createHash('sha256')
    const stream = createReadStream(path)
    stream.on('data', (chunk) => hash.update(chunk))
    stream.once('error', reject)
    stream.once('end', () => resolve(hash.digest('hex')))
  })
}

async function closeAllSessions(entry: {
  client: XlsxSidecarClient
  sessions: Map<string, SessionInfo>
}): Promise<void> {
  const sessions = [...entry.sessions.entries()]
  entry.sessions.clear()
  await Promise.allSettled(
    sessions.map(async ([sessionId, session]) => {
      await cleanupSessionResources({
        tempRoot: app.getPath('temp'),
        snapshotPath: session.snapshotPath,
        importTempDir: session.importTempDir,
        closeSidecar: () => entry.client.close(sessionId),
      })
    }),
  )
}
