@echo off
REM ─────────────────────────────────────────────────────────────────────
REM  Hot-reload dev server. Watches Razor / C# files and rebuilds + reloads
REM  on change. Use this for quick iteration without F5/restart cycles.
REM ─────────────────────────────────────────────────────────────────────

setlocal
set ASPNETCORE_ENVIRONMENT=Development
set ASPNETCORE_URLS=http://*:5050
cd /d "%~dp0JinoSupporter.Web"
dotnet watch run --no-launch-profile
endlocal
