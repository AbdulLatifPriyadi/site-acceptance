@echo off
REM ============================================================
REM  Build index-test.html from index.html
REM  Injects mock fetch/data interceptors so it runs standalone.
REM
REM  Usage: scripts\build-test.bat
REM ============================================================
cd /d "%~dp0.."

echo [build-test] Reading index.html...

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0build-test.ps1"

pause
