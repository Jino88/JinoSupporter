@echo off
setlocal
chcp 65001 >nul

set "ROOT=D:\000. MyWorks\005. Program\Repository\JinoSupporter"
set "WEB_DIR=%ROOT%\JinoSupporter.Web"
set "PROJECT=%WEB_DIR%\JinoSupporter.Web.csproj"
set "EXE=%WEB_DIR%\bin\Debug\net8.0\JinoSupporter.Web.exe"
set "PORT=5050"
set "URL=http://localhost:%PORT%"

title JinoSupporter Web Restart

echo [1/4] Stop existing JinoSupporter Web server...
powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command ^
  "$ErrorActionPreference='SilentlyContinue';" ^
  "$port=%PORT%;" ^
  "$pids=@();" ^
  "$pids += Get-NetTCPConnection -State Listen -LocalPort $port | Select-Object -ExpandProperty OwningProcess -Unique;" ^
  "$pids += Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'JinoSupporter.Web.exe' -or $_.CommandLine -like '*JinoSupporter.Web.dll*' -or $_.CommandLine -like '*JinoSupporter.Web.csproj*' } | Select-Object -ExpandProperty ProcessId;" ^
  "$pids = $pids | Where-Object { $_ -and $_ -ne $PID } | Sort-Object -Unique;" ^
  "if ($pids) { Stop-Process -Id $pids -Force; Write-Host ('Stopped PID: ' + ($pids -join ', ')) } else { Write-Host 'No running server found.' }"

echo.
echo [2/4] Build JinoSupporter.Web...
dotnet build "%PROJECT%" -c Debug
if errorlevel 1 (
    echo.
    echo Build failed. Server was not restarted.
    pause
    exit /b 1
)

if not exist "%EXE%" (
    echo.
    echo Build completed, but executable was not found:
    echo %EXE%
    pause
    exit /b 1
)

echo.
echo [3/4] Open browser: %URL%
start "" "%URL%"

echo.
echo [4/4] Start JinoSupporter.Web on port %PORT%...
echo Stop this server with Ctrl+C or by closing this window.
echo.
set "ASPNETCORE_ENVIRONMENT=Development"
set "ASPNETCORE_URLS=http://0.0.0.0:%PORT%"
cd /d "%WEB_DIR%"
"%EXE%"

echo.
echo Server stopped.
pause
