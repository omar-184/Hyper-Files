; Hypercube Office NSIS customizations (included by electron-builder):
;   - "New > Word/Excel/PowerPoint document" Explorer templates
;   - Default Programs registration (Capabilities + RegisteredApplications),
;     which is what makes Hypercube Office selectable as the default app for PDF and
;     Office files in Windows 10/11 Settings
;   - a finish-page checkbox that opens that Settings page. Windows does not
;     let installers silently take over file types; the user confirms there.

!define HYPERFILES_CAPABILITIES "Software\Hypercube Office\Capabilities"
!define HYPERFILES_REGISTERED_NAME "HypercubeOffice"

; Scope templates to our ProgIDs (electron-builder uses fileAssociations.name).
; A shared .ext\ShellNew would overwrite Office/WPS templates. OOXML files
; must be copied from valid packages, never created with NullFile.
!macro HyperFilesRegisterShellNew EXT PROGID
  WriteRegStr SHELL_CONTEXT "Software\Classes\.${EXT}\${PROGID}\ShellNew" "FileName" "$INSTDIR\resources\shell-new\blank.${EXT}"
!macroend

!macro HyperFilesUnregisterShellNew EXT PROGID
  ; Only remove our own registration, including when uninstalling for an update.
  ReadRegStr $0 SHELL_CONTEXT "Software\Classes\.${EXT}\${PROGID}\ShellNew" "FileName"
  ${If} $0 == "$INSTDIR\resources\shell-new\blank.${EXT}"
    DeleteRegKey SHELL_CONTEXT "Software\Classes\.${EXT}\${PROGID}\ShellNew"
    DeleteRegKey /ifempty SHELL_CONTEXT "Software\Classes\.${EXT}\${PROGID}"
  ${EndIf}
!macroend

; ProgIDs must match fileAssociations[].name (progId() in electron-builder.cjs)
!macro HyperFilesCapability EXT
  WriteRegStr SHELL_CONTEXT "${HYPERFILES_CAPABILITIES}\FileAssociations" ".${EXT}" "HyperFiles.${EXT}"
!macroend

!macro HyperFilesRegisterCapabilities
  WriteRegStr SHELL_CONTEXT "${HYPERFILES_CAPABILITIES}" "ApplicationName" "Hypercube Office"
  WriteRegStr SHELL_CONTEXT "${HYPERFILES_CAPABILITIES}" "ApplicationDescription" "Offline PDF and office suite"
  WriteRegStr SHELL_CONTEXT "${HYPERFILES_CAPABILITIES}" "ApplicationIcon" "$appExe,0"
  !insertmacro HyperFilesCapability "pdf"
  !insertmacro HyperFilesCapability "docx"
  !insertmacro HyperFilesCapability "xlsx"
  !insertmacro HyperFilesCapability "xlsm"
  !insertmacro HyperFilesCapability "xls"
  !insertmacro HyperFilesCapability "pptx"
  !insertmacro HyperFilesCapability "csv"
  !insertmacro HyperFilesCapability "tsv"
  !insertmacro HyperFilesCapability "md"
  !insertmacro HyperFilesCapability "markdown"
  !insertmacro HyperFilesCapability "html"
  !insertmacro HyperFilesCapability "htm"
  WriteRegStr SHELL_CONTEXT "Software\RegisteredApplications" "${HYPERFILES_REGISTERED_NAME}" "${HYPERFILES_CAPABILITIES}"
!macroend

!macro HyperFilesUnregisterCapabilities
  DeleteRegValue SHELL_CONTEXT "Software\RegisteredApplications" "${HYPERFILES_REGISTERED_NAME}"
  DeleteRegKey SHELL_CONTEXT "${HYPERFILES_CAPABILITIES}"
  DeleteRegKey /ifempty SHELL_CONTEXT "Software\Hypercube Office"
!macroend

!macro customInstall
  !insertmacro HyperFilesRegisterShellNew "docx" "HyperFiles.docx"
  !insertmacro HyperFilesRegisterShellNew "xlsx" "HyperFiles.xlsx"
  !insertmacro HyperFilesRegisterShellNew "pptx" "HyperFiles.pptx"
  !insertmacro HyperFilesRegisterCapabilities
  !insertmacro UPDATEFILEASSOC
!macroend

!macro customUnInstall
  Push $0
  !insertmacro HyperFilesUnregisterShellNew "docx" "HyperFiles.docx"
  !insertmacro HyperFilesUnregisterShellNew "xlsx" "HyperFiles.xlsx"
  !insertmacro HyperFilesUnregisterShellNew "pptx" "HyperFiles.pptx"
  Pop $0
  !insertmacro HyperFilesUnregisterCapabilities
  !insertmacro UPDATEFILEASSOC
!macroend

; Same finish page electron-builder builds by default (run the app), plus a
; second checkbox for the default-app offer.
!macro customFinishPage
  Function HyperFilesStartApp
    ${if} ${isUpdated}
      StrCpy $1 "--updated"
    ${else}
      StrCpy $1 ""
    ${endif}
    ${StdUtils.ExecShellAsUser} $0 "$launchLink" "open" "$1"
  FunctionEnd

  ; Windows 11 opens Hypercube Office's own page under Settings > Apps > Default
  ; apps; Windows 10 ignores the parameter and opens Default apps.
  Function HyperFilesOpenDefaultApps
    ${if} $installMode == "all"
      ExecShell "open" "ms-settings:defaultapps?registeredAppMachine=${HYPERFILES_REGISTERED_NAME}"
    ${else}
      ExecShell "open" "ms-settings:defaultapps?registeredAppUser=${HYPERFILES_REGISTERED_NAME}"
    ${endif}
  FunctionEnd

  !define MUI_FINISHPAGE_RUN
  !define MUI_FINISHPAGE_RUN_FUNCTION "HyperFilesStartApp"
  !define MUI_FINISHPAGE_SHOWREADME
  !define MUI_FINISHPAGE_SHOWREADME_TEXT "Make Hypercube Office the default app for PDF and Office files"
  !define MUI_FINISHPAGE_SHOWREADME_FUNCTION "HyperFilesOpenDefaultApps"
  !insertmacro MUI_PAGE_FINISH
!macroend
