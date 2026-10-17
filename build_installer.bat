@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set "NSIS_COMPILER=%~1"
if not defined NSIS_COMPILER if exist "%ProgramFiles(x86)%\NSIS\makensis.exe" set "NSIS_COMPILER=%ProgramFiles(x86)%\NSIS\makensis.exe"
if not defined NSIS_COMPILER if exist "%ProgramFiles%\NSIS\makensis.exe" set "NSIS_COMPILER=%ProgramFiles%\NSIS\makensis.exe"
if not defined NSIS_COMPILER (
    echo Не найден NSIS. Передайте путь к makensis.exe первым аргументом.
    exit /b 1
)
if not exist "%NSIS_COMPILER%" (
    echo Не найден компилятор: %NSIS_COMPILER%
    exit /b 1
)
call build.bat build\installer-app
if errorlevel 1 exit /b %errorlevel%
"%NSIS_COMPILER%" /INPUTCHARSET UTF8 /V3 installer.nsi
if errorlevel 1 exit /b %errorlevel%
echo Готово: %CD%\KeyFlipSetup.exe
