# JinoSupporter Migration Plan

## Goal

Move only the required module files from `DataGraphHost` into this repository and run everything from a single WPF host.

`WebView2` is no longer part of the target architecture in this repository.

## Current Result

- Main shell is WPF-based.
- Left navigation selects modules.
- Right content area hosts each module's main screen.
- `Translator` and `FileTransfer` now use WPF replacement views in this project.
- Old `WebView2` host code copied into this repo for those modules has been removed.

## Source Mapping

- `DataMaker` -> `WorkbenchHost/Modules/DataMaker`
- `GraphMaker:ScatterPlot` -> `WorkbenchHost/Modules/GraphMaker/ScatterPlot` plus shared `GraphMaker/Common`, `Themes`, viewers, and supporting graph windows
- `GraphMaker:ProcessTrend` -> `WorkbenchHost/Modules/GraphMaker/ProcessTrend` plus shared `GraphMaker/Common`, `Themes`, viewers, and supporting graph windows
- `GraphMaker:UnifiedMultiY` -> `WorkbenchHost/Modules/GraphMaker/Common/UnifiedMultiYView*` and `GraphMaker/ValuePlot/*` plus shared graph files
- `GraphMaker:DailyDataTrendExtra` -> `WorkbenchHost/Modules/GraphMaker/MultiXMultiY` plus shared graph files
- `GraphMaker:HeatMap` -> `WorkbenchHost/Modules/GraphMaker/HeatMap` plus shared graph files
- `GraphMaker:AudioBusData` -> `WorkbenchHost/Modules/GraphMaker/AudioBus` plus shared graph files
- `DiskTree` -> `WorkbenchHost/Modules/DiskTree`
- `Memo` -> `WorkbenchHost/Modules/Memo`
- `VideoConverter` -> `WorkbenchHost/Modules/VideoConverter`
- `JsonEditor` -> `WorkbenchHost/Modules/JsonEditor`
- `ScreenCapture` -> `WorkbenchHost/Modules/ScreenCapture`
- `FileTransfer` -> migrated to local WPF replacement view
- `Translator` -> migrated to local WPF replacement view
- `Settings` -> `WorkbenchHost/Modules/AppSettings`

## Shared Files Required

- `WorkbenchHost/Infrastructure/AppSettingsPathManager.cs`
- `WorkbenchHost/Infrastructure/WorkbenchSettingsStore.cs`

These are directly referenced by:

- `DataMaker`
- `DiskTree`
- `Memo`
- `VideoConverter`
- `ScreenCapture`
- `Settings`
- part of `GraphMaker` through BMES/settings flow

## New Host Strategy

- keep the current final module lineup in a single WPF shell
- use left-side navigation for module selection
- host WPF controls or adapted WPF content directly in the right pane
- keep existing secondary windows where modules already depend on them
- avoid carrying browser-based routing or message passing into the new project

## Module Dependency Notes

### DataMaker

Required:

- full `Modules/DataMaker`
- `Infrastructure/*Settings*`

Direct external packages:

- `System.Data.SQLite`
- `Newtonsoft.Json`

Strong coupling:

- settings store
- BMES credential window path from shared settings
- app-level storage assumptions

### GraphMaker Suite

Required:

- nearly all `Modules/GraphMaker` shared code, not only the named subfolders

Why:

- the listed tools share helpers, dialogs, result windows, themes, and plotting helpers across folders
- `UnifiedMultiY` and other graph modes reuse common result/detail infrastructure

Direct external packages:

- `OxyPlot.Wpf`
- `ExcelDataReader`
- `ExcelDataReader.DataSet`
- `EPPlus`
- `System.Text.Encoding.CodePages`

Strong coupling:

- shared `ModernTheme.xaml`
- shared helper windows in `Common`
- some modes depend on graph viewers outside their own subfolder

### DiskTree

Required:

- full `Modules/DiskTree`
- `Infrastructure/*Settings*`

Direct external packages:

- `Microsoft.Data.Sqlite`

Strong coupling:

- default DB path conventions
- shared storage root

### Memo

Required:

- full `Modules/Memo`
- `Infrastructure/*Settings*`

Direct external packages:

- `Microsoft.Data.Sqlite`

Strong coupling:

- memo DB path stored in shared settings JSON

### VideoConverter

Required:

- full `Modules/VideoConverter`
- `Infrastructure/*Settings*`

Strong coupling:

- settings persistence through shared settings store
- `Settings` module references `VideoConverter.MainWindow.SettingsFilePath`
- still uses `System.Windows.Forms.FolderBrowserDialog`, so project keeps `UseWindowsForms=true`

### JsonEditor

Required:

- full `Modules/JsonEditor`

Direct external packages:

- `Newtonsoft.Json`

Coupling level:

- low to medium

### ScreenCapture

Required:

- full `Modules/ScreenCapture`
- `Infrastructure/*Settings*`

Strong coupling:

- global hotkey registration
- overlay windows
- settings persistence

### FileTransfer

Current state:

- old browser-based host code has been removed from this repository
- current module entry in the shell is a WPF replacement view

Risk:

- if feature parity is required later, this module still needs a deeper native WPF implementation

### Translator

Current state:

- old browser-based host code has been removed from this repository
- current module entry in the shell is a WPF replacement view

Risk:

- if feature parity is required later, this module still needs a deeper native WPF implementation

### Settings

Required:

- full `Modules/AppSettings`
- `Infrastructure/*Settings*`

Indirect dependencies:

- `DataMaker.R6.FetchDataBMES.FormSettingBMESWindow`
- `VideoConverter.MainWindow`
- `ScreenCapture.CaptureHotkeyManager`
- `DiskTree.MainWindow`

Risk:

- even the settings screen is coupled to concrete module classes to expose their storage paths

## Execution Status

1. Required module folders identified.
2. Direct dependencies and shared infrastructure traced.
3. Files copied into the new WPF project.
4. `csproj` created and package references aligned.
5. Main shell converted to WPF navigation and right-pane module hosting.
6. Browser-based module remnants removed from this repo where they would cause confusion.
7. Build revalidated successfully.

## Remaining Hard Dependencies

- `Settings` still references concrete classes from multiple modules
- several module namespaces are top-level (`DataMaker`, `GraphMaker`, `DiskTree`, `VideoConverter`) rather than `WorkbenchHost.*`, so namespace normalization still needs care if full unification is desired
- some modules still rely on older window-centric patterns and can be hosted, but are not yet clean MVVM-style components
