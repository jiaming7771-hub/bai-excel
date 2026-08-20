@echo off
setlocal EnableExtensions
title Excel Toolbox Windows Build
color 0A
chcp 65001 >nul

if "%~1"=="" (
  cmd /k "%~f0" _inner
  exit /b
)

cd /d "%~dp0.."
if errorlevel 1 (
  echo [ERROR] cannot enter project folder
  goto :end
)

echo ========================================
echo   Excel Toolbox - Windows Build
echo   Project: %CD%
echo ========================================
echo.

where py >nul 2>nul
if %errorlevel%==0 (
  set "PY=py -3"
  goto :have_py
)
where python >nul 2>nul
if %errorlevel%==0 (
  set "PY=python"
  goto :have_py
)
where python3 >nul 2>nul
if %errorlevel%==0 (
  set "PY=python3"
  goto :have_py
)

echo [ERROR] Python not found. Install 3.10+ and check Add to PATH.
goto :end

:have_py
echo [1/5] Python OK: %PY%
%PY% --version
if errorlevel 1 goto :fail

echo.
echo [2/5] Prepare Windows venv...
if exist ".venv\Scripts\python.exe" goto :venv_ok

echo Removing old/Mac .venv...
if exist ".venv" (
  attrib -R ".venv\*.*" /S /D >nul 2>nul
  rmdir /s /q ".venv" 2>nul
)
if exist ".venv" (
  echo [ERROR] Cannot delete .venv - delete it manually then retry
  goto :fail
)
%PY% -m venv .venv
if errorlevel 1 goto :fail

:venv_ok
if not exist ".venv\Scripts\python.exe" (
  echo [ERROR] venv missing Scripts\python.exe
  goto :fail
)
set "VPY=%CD%\.venv\Scripts\python.exe"

echo.
echo [3/5] Install dependencies...
"%VPY%" -m pip install --upgrade pip
if errorlevel 1 goto :fail
"%VPY%" -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo [4/5] Building exe (name: ExcelToolbox.exe)...
"%VPY%" -m PyInstaller --noconfirm --clean excel_toolbox_windows.spec
if errorlevel 1 goto :fail

set "EXE="
if exist "dist\ExcelToolbox.exe" set "EXE=dist\ExcelToolbox.exe"
if exist "dist\Excel小工具箱.exe" set "EXE=dist\Excel小工具箱.exe"
if "%EXE%"=="" (
  echo [ERROR] No exe found in dist\
  dir /b dist
  goto :fail
)
echo Found: %EXE%

echo.
echo [5/5] Copy to Desktop...
set "OUT=%USERPROFILE%\Desktop\ExcelToolbox-Release"
if not exist "%OUT%" mkdir "%OUT%"
copy /Y "%EXE%" "%OUT%\ExcelToolbox.exe"
if errorlevel 1 goto :fail

echo.
echo ========================================
echo   BUILD OK
echo   File: %OUT%\ExcelToolbox.exe
echo   Send this exe to Windows users.
echo ========================================
goto :end

:fail
echo.
echo ========================================
echo   BUILD FAILED
echo ========================================
echo If dist already has an exe, you can use it directly:
echo   %CD%\dist\
dir /b "dist\*.exe" 2>nul

:end
echo.
echo Press any key to close...
pause >nul
exit /b 0
