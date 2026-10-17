Unicode true
!include "MUI2.nsh"
!include "LogicLib.nsh"
!include "WinVer.nsh"
!include "x64.nsh"

!define APP_VERSION "1.0.0"
!define APP_EXE "build\installer-app\KeyFlip.exe"
!define UNINSTALL_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\KeyFlip"

Name "KeyFlip"
OutFile "KeyFlipSetup.exe"
InstallDir "$LOCALAPPDATA\Programs\KeyFlip"
InstallDirRegKey HKCU "${UNINSTALL_KEY}" "InstallLocation"
RequestExecutionLevel user
SetCompressor /SOLID lzma
ShowInstDetails show
ShowUninstDetails show
Icon "resources\keyflip.ico"
UninstallIcon "resources\keyflip.ico"
VIProductVersion "${APP_VERSION}.0"
VIAddVersionKey /LANG=1049 "ProductName" "KeyFlip"
VIAddVersionKey /LANG=1049 "FileDescription" "Установщик KeyFlip"
VIAddVersionKey /LANG=1049 "FileVersion" "${APP_VERSION}"
VIAddVersionKey /LANG=1049 "LegalCopyright" "Copyright 2026 KeyFlip"

!define MUI_ABORTWARNING
!define MUI_ICON "resources\keyflip.ico"
!define MUI_UNICON "resources\keyflip.ico"
!define MUI_WELCOMEPAGE_TEXT "Мастер установит KeyFlip для вашей учётной записи.$\r$\n$\r$\nPython и права администратора на этом компьютере не нужны.$\r$\n$\r$\nПеред установкой закройте работающий KeyFlip через меню трея → «Выход»."
!define MUI_FINISHPAGE_RUN "$INSTDIR\KeyFlip.exe"
!define MUI_FINISHPAGE_RUN_TEXT "Запустить KeyFlip"
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_UNPAGE_FINISH
!insertmacro MUI_LANGUAGE "Russian"

!macro CHECK_RUNNING
  check_running:
  System::Call 'kernel32::CreateMutexW(p 0, i 0, w "Global\KeyFlipMutex") p .r2 ?e'
  Pop $3
  ${If} $2 != 0
    System::Call 'kernel32::CloseHandle(p r2)'
  ${EndIf}
  ${If} $3 == 183
    MessageBox MB_RETRYCANCEL|MB_ICONEXCLAMATION "KeyFlip уже работает. Выберите «Выход» в меню трея и нажмите «Повторить»." /SD IDCANCEL IDRETRY check_running
    Abort
  ${EndIf}
  ${If} $2 == 0
    Abort "Не удалось проверить состояние KeyFlip."
  ${EndIf}
!macroend

Function .onInit
  SetShellVarContext current
  ${IfNot} ${RunningX64}
    MessageBox MB_OK|MB_ICONSTOP "Нужна 64-разрядная Windows 10 или новее."
    Abort
  ${EndIf}
  ${IfNot} ${AtLeastWin10}
    MessageBox MB_OK|MB_ICONSTOP "Нужна Windows 10 или новее."
    Abort
  ${EndIf}
FunctionEnd

Section "Установка"
  ReadRegStr $1 HKCU "${UNINSTALL_KEY}" "InstallLocation"
  ${If} $1 != ""
    ${If} $1 != $INSTDIR
      Abort "KeyFlip уже установлен в $1. Для смены папки сначала удалите старую установку."
    ${EndIf}
  ${Else}
    IfFileExists "$INSTDIR\KeyFlip.exe" 0 +2
      Abort "В выбранной папке уже есть KeyFlip.exe. Выберите другую папку."
  ${EndIf}

  !insertmacro CHECK_RUNNING

  SetOutPath "$INSTDIR"
  File /oname=KeyFlip.exe "${APP_EXE}"
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateDirectory "$SMPROGRAMS\KeyFlip"
  CreateShortcut "$SMPROGRAMS\KeyFlip\KeyFlip.lnk" "$INSTDIR\KeyFlip.exe"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayName" "KeyFlip"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "DisplayIcon" "$INSTDIR\KeyFlip.exe,0"
  WriteRegStr HKCU "${UNINSTALL_KEY}" "UninstallString" '$\"$INSTDIR\Uninstall.exe$\"'
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoModify" 1
  WriteRegDWORD HKCU "${UNINSTALL_KEY}" "NoRepair" 1
SectionEnd

Function un.onInit
  SetShellVarContext current
FunctionEnd

Section "Uninstall"
  !insertmacro CHECK_RUNNING
  ClearErrors
  ExecWait '"$INSTDIR\KeyFlip.exe" --uninstall-cleanup' $0
  ${If} ${Errors}
    Abort "Не удалось проверить работающий KeyFlip. Файлы сохранены."
  ${EndIf}
  ${If} $0 != 0
    Abort "Не удалось безопасно удалить автозапуск. Файлы сохранены."
  ${EndIf}

  Delete "$INSTDIR\KeyFlip.exe"
  ${If} ${Errors}
    Abort "Не удалось удалить KeyFlip.exe. Закройте приложение и повторите."
  ${EndIf}
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"
  Delete "$SMPROGRAMS\KeyFlip\KeyFlip.lnk"
  RMDir "$SMPROGRAMS\KeyFlip"
  DeleteRegKey HKCU "${UNINSTALL_KEY}"
SectionEnd
