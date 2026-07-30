# BMES NG Rate Standalone

WPF + BlazorWebView desktop app for the BMES and NG Rate screens.

This app does not host or connect to the JinoSupporter web server. It builds as its own local executable. Before each build, `tools/SyncBmesFromWeb.ps1` extracts the BMES/NG Rate Razor and service source files from `JinoSupporter.Web`, rewrites the namespace to `BmesNgRateStandalone`, removes server render-mode directives, and compiles those files directly into this project.

## Source Sync

The build target in `BmesNgRateStandalone.csproj` runs:

```cmd
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\SyncBmesFromWeb.ps1
```

So changes made in the JinoSupporter BMES/NG Rate source are reflected in the standalone project on the next build, without referencing `JinoSupporter.Web.dll`.

## Build

From the repository root:

```cmd
dotnet build .\BmesNgRateStandalone\BmesNgRateStandalone.csproj
```

The project is also included in `JinoSupporter.sln`.

## Installer

The installer is built with Inno Setup. Install Inno Setup 6, then run:

```cmd
powershell -NoProfile -ExecutionPolicy Bypass -File .\BmesNgRateStandalone\tools\BuildStandaloneInstaller.ps1
```

The script publishes the app as a self-contained win-x64 build to
`bin/Release/net8.0-windows/win-x64/publish` and creates:

```txt
BmesNgRateStandalone/dist/BmesNgRateStandalone_Setup-<version>.exe
```

The same flow is available from VS Code task `build-standalone-installer`.

`PublishStandaloneUpdate.ps1` calls this script with `-SkipPublish` and copies the installer
into `JinoSupporter.Web/standalone-updates/`, where the web app's **Tools > PC Download** page
serves it. If Inno Setup is missing the publish still succeeds and the page offers the zip only.

## Auto Update

The standalone executable checks this manifest on startup:

```txt
http://10.6.4.54:5050/standalone/update.json
```

Normal app usage is still local and does not require the web server. The server is used only to check for and download a newer standalone package.

To publish an update package into `JinoSupporter.Web/standalone-updates`, run:

```cmd
powershell -NoProfile -ExecutionPolicy Bypass -File .\BmesNgRateStandalone\tools\PublishStandaloneUpdate.ps1 -Version 1.0.1
```

The script creates `update.json` and a zip package. `JinoSupporter.Web` serves them from `/standalone/update.json` and `/standalone/download/{fileName}`.
