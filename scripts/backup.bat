@echo off
REM ============================================================
REM  Backup script — creates timestamped copies of index.html and
REM  SiteAcceptanceTracker.gs in the ../backups/ folder.
REM
REM  Usage:
REM    cd site-acceptance
REM    scripts\backup.bat
REM
REM  Or double-click this file from Windows Explorer.
REM ============================================================
cd /d "%~dp0.."

set BACKUP_DIR=%CD%\backups
set TIMESTAMP=%date:~-4%-%date:~4,2%-%date:~7,2%

if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"

if exist index.html (
    copy /y index.html "%BACKUP_DIR%\index_%TIMESTAMP%.html" >nul
    echo [backup] index_%TIMESTAMP%.html
)

if exist SiteAcceptanceTracker.gs (
    copy /y SiteAcceptanceTracker.gs "%BACKUP_DIR%\SiteAcceptanceTracker_%TIMESTAMP%.gs" >nul
    echo [backup] SiteAcceptanceTracker_%TIMESTAMP%.gs
)

echo [backup] Done. Files saved to %BACKUP_DIR%
echo.
dir /b /o-d "%BACKUP_DIR%" | findstr /i "\.html \.gs$" | head -10
pause
