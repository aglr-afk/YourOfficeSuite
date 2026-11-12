Unicode true
!include "MUI2.nsh"
!include "FileFunc.nsh"
!define APP "YourOfficeSuite"
!define EXE "YourOfficeSuite.exe"
!define VER "2.3.0"
!define UNKEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\YourOfficeSuite"
Name "${APP}"
OutFile "../release/YourOfficeSuite-Setup-${VER}.exe"
InstallDir "$LOCALAPPDATA\Programs\${APP}"
InstallDirRegKey HKCU "Software\YourOfficeSuite" "InstallDir"
RequestExecutionLevel user
SetCompressor /SOLID lzma
SetCompressorDictSize 64
ManifestDPIAware true
BrandingText "${APP} ${VER}"
VIProductVersion "${VER}.0"
VIAddVersionKey "ProductName" "${APP}"
VIAddVersionKey "FileDescription" "${APP} Setup"
VIAddVersionKey "FileVersion" "${VER}"
VIAddVersionKey "LegalCopyright" "AGPL-3.0"
!define MUI_ICON "icon.ico"
!define MUI_UNICON "icon.ico"
!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\${EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "Start ${APP}"
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

!macro CheckRunning
  IfFileExists "$INSTDIR\${EXE}" 0 +7
  ClearErrors
  FileOpen $0 "$INSTDIR\${EXE}" a
  IfErrors 0 +3
    MessageBox MB_OK|MB_ICONEXCLAMATION "${APP} is running. Close it, then run this again."
    Abort
  FileClose $0
!macroend

!macro Assoc EXT DESC
  WriteRegStr HKCU "Software\Classes\YourOfficeSuite.${EXT}" "" "${DESC}"
  WriteRegStr HKCU "Software\Classes\YourOfficeSuite.${EXT}\DefaultIcon" "" "$INSTDIR\${EXE},0"
  WriteRegStr HKCU "Software\Classes\YourOfficeSuite.${EXT}\shell\open\command" "" '"$INSTDIR\${EXE}" "%1"'
  WriteRegStr HKCU "Software\Classes\.${EXT}\OpenWithProgids" "YourOfficeSuite.${EXT}" ""
  WriteRegStr HKCU "Software\Classes\Applications\${EXE}\SupportedTypes" ".${EXT}" ""
!macroend
!macro UnAssoc EXT
  DeleteRegKey HKCU "Software\Classes\YourOfficeSuite.${EXT}"
  DeleteRegValue HKCU "Software\Classes\.${EXT}\OpenWithProgids" "YourOfficeSuite.${EXT}"
!macroend

Section "Install"
  !insertmacro CheckRunning
  RMDir /r "$INSTDIR\resources"
  RMDir /r "$INSTDIR\locales"
  SetOutPath "$INSTDIR"
  File /r "../release/win-unpacked/*.*"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateShortcut "$SMPROGRAMS\${APP}.lnk" "$INSTDIR\${EXE}"
  CreateShortcut "$DESKTOP\${APP}.lnk" "$INSTDIR\${EXE}"
  WriteRegStr HKCU "Software\YourOfficeSuite" "InstallDir" "$INSTDIR"
  WriteRegStr HKCU "${UNKEY}" "DisplayName" "${APP}"
  WriteRegStr HKCU "${UNKEY}" "DisplayVersion" "${VER}"
  WriteRegStr HKCU "${UNKEY}" "Publisher" "YourOfficeSuite"
  WriteRegStr HKCU "${UNKEY}" "DisplayIcon" "$INSTDIR\${EXE}"
  WriteRegStr HKCU "${UNKEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNKEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegDWORD HKCU "${UNKEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNKEY}" "NoRepair" 1
  ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
  IntFmt $0 "0x%08X" $0
  WriteRegDWORD HKCU "${UNKEY}" "EstimatedSize" "$0"
  WriteRegStr HKCU "Software\Classes\Applications\${EXE}" "FriendlyAppName" "${APP}"
  WriteRegStr HKCU "Software\Classes\Applications\${EXE}\shell\open\command" "" '"$INSTDIR\${EXE}" "%1"'
  !insertmacro Assoc "docx" "Word Document"
  !insertmacro Assoc "xlsx" "Excel Workbook"
  !insertmacro Assoc "pptx" "PowerPoint Presentation"
  !insertmacro Assoc "pdf" "PDF Document"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
SectionEnd

Section "Uninstall"
  !insertmacro CheckRunning
  Delete "$SMPROGRAMS\${APP}.lnk"
  Delete "$DESKTOP\${APP}.lnk"
  RMDir /r "$INSTDIR"
  DeleteRegKey HKCU "${UNKEY}"
  DeleteRegKey HKCU "Software\YourOfficeSuite"
  DeleteRegKey HKCU "Software\Classes\Applications\${EXE}"
  !insertmacro UnAssoc "docx"
  !insertmacro UnAssoc "xlsx"
  !insertmacro UnAssoc "pptx"
  !insertmacro UnAssoc "pdf"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
SectionEnd
