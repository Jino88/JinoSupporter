@echo off
REM ─────────────────────────────────────────────────────────────────────
REM  Runs the already-built JinoSupporter.Web on port 5050 without
REM  re-compiling. F5 (Web Server) launches a debugger AND rebuilds; this
REM  script just launches the last-compiled DLL.
REM
REM  Use F5 in VSCode/VS when you want fresh code; use this when you only
REM  want to spin the server back up after a clean compile already exists.
REM ─────────────────────────────────────────────────────────────────────

setlocal
set ASPNETCORE_ENVIRONMENT=Development
set ASPNETCORE_URLS=http://*:5050

set EXE=%~dp0JinoSupporter.Web\bin\Debug\net8.0\JinoSupporter.Web.exe

if not exist "%EXE%" (
    echo [run-web] %EXE% not found. Build the project first ^(F5 once or `dotnet build`^).
    pause
    exit /b 1
)

echo [run-web] Launching %EXE% on http://localhost:5050 ...
"%EXE%"
endlocal
