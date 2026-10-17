@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo Не найден Python проекта. Создайте .venv и установите requirements-build.txt.
    exit /b 1
)
set "DIST_DIR=%~dp0dist"
if not "%~1"=="" set "DIST_DIR=%~1"
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --windowed --onefile --name KeyFlip --icon "%CD%\resources\keyflip.ico" --add-data "%CD%\resources;resources" --distpath "%DIST_DIR%" --specpath "%CD%\build" --workpath "%CD%\build" "%CD%\app.pyw"
if errorlevel 1 exit /b %errorlevel%
echo Готово: %DIST_DIR%\KeyFlip.exe
