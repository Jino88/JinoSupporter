@echo off
REM ─────────────────────────────────────────────────────────────────────
REM  JinoSupporter Web — server toolbox (menu).
REM  Double-click. Pick a number. Window stays open until you press Q.
REM ─────────────────────────────────────────────────────────────────────

setlocal enabledelayedexpansion

set "ROOT=%~dp0"
set "CSPROJ=%ROOT%JinoSupporter.Web\JinoSupporter.Web.csproj"
set "DIST=%ROOT%dist\JinoSupporterWeb"
set "EXE=%DIST%\JinoSupporter.Web.exe"
set "TASK=JinoSupporterWebAutoStart"
set "PORT=5050"


:menu
cls
echo.
echo  ╔══════════════════════════════════════════════════╗
echo  ║   JinoSupporter Web — Server Toolbox             ║
echo  ╚══════════════════════════════════════════════════╝
echo   Project root : %ROOT%
echo   Publish dir  : %DIST%
if exist "%EXE%" (echo   Build status : OK ^(exe found^)) else (echo   Build status : NOT BUILT yet)
echo.
echo   [1] Publish        - build deployable copy
echo   [2] Start server   - launch published exe ^(port %PORT%^)
echo   [3] User shortcut  - generate JinoSupporter.url
echo   [4] Open firewall  - allow inbound TCP %PORT% ^(Admin needed^)
echo   [5] Install autostart   - run on logon ^(Admin needed^)
echo   [6] Remove  autostart
echo   [Q] Quit
echo.
set "choice="
set /p "choice=Choice: "
if "%choice%"=="" goto menu
if /I "%choice%"=="1" call :do_publish    & goto pause_back
if /I "%choice%"=="2" call :do_start      & goto pause_back
if /I "%choice%"=="3" call :do_shortcut   & goto pause_back
if /I "%choice%"=="4" call :do_firewall   & goto pause_back
if /I "%choice%"=="5" call :do_autostart  & goto pause_back
if /I "%choice%"=="6" call :do_rmautostart & goto pause_back
if /I "%choice%"=="q" goto bye
echo Unknown choice: %choice%
timeout /t 2 >nul
goto menu


:pause_back
echo.
echo Press any key to return to menu...
pause >nul
goto menu


:bye
endlocal
exit /b 0


REM ════════════════════ Action subroutines ═════════════════════════════

:do_publish
echo.
where dotnet >nul 2>&1
if errorlevel 1 (
    echo [publish] dotnet not found in PATH. Install .NET 8 SDK first.
    exit /b 1
)
if exist "%DIST%" (
    echo [publish] Cleaning %DIST% ...
    rd /s /q "%DIST%"
)
echo [publish] dotnet publish %CSPROJ%
echo                  to %DIST%
echo.
dotnet publish "%CSPROJ%" -c Release -o "%DIST%" --self-contained false -r win-x64
if errorlevel 1 (
    echo [publish] FAILED.
    exit /b 1
)
> "%DIST%\start-server.cmd" (
    echo @echo off
    echo setlocal
    echo set ASPNETCORE_ENVIRONMENT=Production
    echo set ASPNETCORE_URLS=http://*:%PORT%
    echo "%%~dp0JinoSupporter.Web.exe"
)
echo.
echo [publish] Done.
echo            Folder : %DIST%
echo            Launch : double-click start-server.cmd inside it.
echo            Target PC needs .NET 8 Desktop Runtime
echo            ^(https://dotnet.microsoft.com/download/dotnet/8.0^)
exit /b 0


:do_start
echo.
if not exist "%EXE%" (
    echo [start] %EXE% not found.
    echo         Run option [1] Publish first.
    exit /b 1
)
echo [start] Launching server on http://localhost:%PORT%
echo         Stop with Ctrl+C in this window.
echo.
set ASPNETCORE_ENVIRONMENT=Production
set ASPNETCORE_URLS=http://*:%PORT%
"%EXE%"
exit /b 0


:do_shortcut
echo.
set "DEFAULT_HOST=jinosupport.local"
set "HOST="
set /p "HOST=Server host or IP [%DEFAULT_HOST%]: "
if "!HOST!"=="" set "HOST=%DEFAULT_HOST%"

set "PORT_INPUT="
set /p "PORT_INPUT=Port [%PORT%]: "
if "!PORT_INPUT!"=="" set "PORT_INPUT=%PORT%"

set "OUT=%ROOT%JinoSupporter.url"
(
    echo [InternetShortcut]
    echo URL=http://!HOST!:!PORT_INPUT!
    echo IconIndex=0
) > "%OUT%"
echo.
echo [shortcut] Created: %OUT%
echo            URL    : http://!HOST!:!PORT_INPUT!
echo            Share this .url with end users.
exit /b 0


:do_firewall
echo.
echo [firewall] Adding inbound TCP %PORT% rule "JinoSupporter Web" ...
netsh advfirewall firewall add rule name="JinoSupporter Web" dir=in action=allow protocol=TCP localport=%PORT%
if errorlevel 1 (
    echo [firewall] FAILED. Re-run this cmd as Administrator.
    exit /b 1
)
echo [firewall] OK.
exit /b 0


:do_autostart
echo.
set "STARTCMD=%DIST%\start-server.cmd"
if not exist "%STARTCMD%" (
    echo [autostart] %STARTCMD% not found. Run [1] Publish first.
    exit /b 1
)
schtasks /Create /TN "%TASK%" /TR "\"%STARTCMD%\"" /SC ONLOGON /RL HIGHEST /F
if errorlevel 1 (
    echo [autostart] FAILED. Re-run this cmd as Administrator.
    exit /b 1
)
echo.
echo [autostart] Done. Server starts on next user logon.
echo            Run now without logout :  schtasks /Run /TN "%TASK%"
exit /b 0


:do_rmautostart
echo.
schtasks /Delete /TN "%TASK%" /F
exit /b 0
