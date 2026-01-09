@echo off
REM build-installer.cmd - Wrapper to call PowerShell build script
REM Usage: build-installer.cmd [options]
REM   No options    - Full build (publish + installer)
REM   -SkipPublish  - Only compile installer

setlocal EnableDelayedExpansion

set "PSCMD="

REM Try PowerShell Core first
where pwsh >nul 2>&1
if !errorlevel!==0 (
    set "PSCMD=pwsh"
    goto :runps
)

REM Try Windows PowerShell
if exist "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" (
    set "PSCMD=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
    goto :runps
)

echo ERROR: PowerShell not found!
echo Please install PowerShell or run the commands manually.
exit /b 1

:runps
"!PSCMD!" -ExecutionPolicy Bypass -File "%~dp0build-installer.ps1" %*
set "EXITCODE=!errorlevel!"
exit /b !EXITCODE!
