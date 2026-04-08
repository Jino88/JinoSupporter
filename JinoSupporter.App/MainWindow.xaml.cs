using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Windows.Media;
using JinoSupporter.App.Infrastructure.Shell;
using JinoSupporter.App.Modules.FileTransfer;
using JinoSupporter.App.Modules.Translator;
using WorkbenchHost.Modules.AppSettings;
using WorkbenchHost.Modules.JsonEditor;
using WorkbenchHost.Modules.Memo;
using WorkbenchHost.Modules.ScreenCapture;

namespace JinoSupporter.App;

public partial class MainWindow : Window
{
    private const string GraphMakerGroupKey = "GraphMaker";

    private static readonly Uri GraphMakerThemeUri =
        new("pack://application:,,,/Modules/GraphMaker/Themes/ModernTheme.xaml", UriKind.Absolute);

    private readonly Dictionary<string, ShellModuleDefinition> _modules = new(StringComparer.Ordinal);
    private readonly Dictionary<string, Button> _menuButtons = new(StringComparer.Ordinal);
    private readonly HashSet<string> _graphMakerTargets = new(StringComparer.Ordinal)
    {
        "GraphMaker:ScatterPlot",
        "GraphMaker:ProcessTrend",
        "GraphMaker:UnifiedMultiY",
        "GraphMaker:DailyDataTrendExtra",
        "GraphMaker:HeatMap",
        "GraphMaker:AudioBusData"
    };

    private global::DataMaker.MainWindow? _dataMakerWindow;
    private global::DiskTree.MainWindow? _diskTreeWindow;
    private global::VideoConverter.MainWindow? _videoConverterWindow;
    private global::GraphMaker.ScatterPlotView? _scatterPlotView;
    private global::GraphMaker.ProcessFlowTrendView? _processTrendView;
    private global::GraphMaker.UnifiedMultiYView? _unifiedMultiYView;
    private global::GraphMaker.DailyDataTrendExtraView? _dailyTrendExtraView;
    private global::GraphMaker.HeatMapView? _heatMapView;
    private global::GraphMaker.AudioBusDataView? _audioBusDataView;
    private MemoView? _memoView;
    private JsonEditorView? _jsonEditorView;
    private ScreenCaptureView? _screenCaptureView;
    private AppSettingsView? _appSettingsView;
    private FileTransferView? _fileTransferView;
    private TranslatorView? _translatorView;

    private FrameworkElement? _dataMakerContent;
    private FrameworkElement? _diskTreeContent;
    private FrameworkElement? _videoConverterContent;
    private StackPanel? _graphMakerChildrenHost;
    private TextBlock? _graphMakerChevron;
    private bool _isGraphMakerExpanded;
    private string? _activeTarget;

    public MainWindow()
    {
        InitializeComponent();
        RegisterModules();
        BuildMenu();
        SelectModule("DataMaker");
    }

    private void RegisterModules()
    {
        RegisterModule(new ShellModuleDefinition(
            "DataMaker",
            string.Empty,
            "DataMaker",
            "DB load, mapping, grouping, and report setup.",
            "Core",
            "DataMaker owns core data workflows.",
            new[] { "Mapping and AI helper actions are surfaced in the shell.", "The hosted workspace still runs the native DataMaker logic." },
            () => GetHostedWindowContent(GetDataMakerWindow(), ref _dataMakerContent),
            new DelegateModuleBackend<global::DataMaker.MainWindow>(
                "DataMaker",
                new[]
                {
                    new ShellActionDefinition("show-get-data", "Get Data"),
                    new ShellActionDefinition("load-db", "Load DB...", true),
                    new ShellActionDefinition("load-group-json", "Load Group JSON"),
                    new ShellActionDefinition("scan-mapping", "Scan Mapping"),
                    new ShellActionDefinition("show-model-groups", "Model Groups"),
                    new ShellActionDefinition("get-report", "Get Report"),
                    new ShellActionDefinition("show-settings", "Settings"),
                    new ShellActionDefinition("bmes-settings", "BMES Setting"),
                    new ShellActionDefinition("get-bmes-data", "Get BMES Data", true),
                    new ShellActionDefinition("get-routing", "Get Routing"),
                    new ShellActionDefinition("check-debug-rate", "Check Debug Rate"),
                    new ShellActionDefinition("show-routing-mapping", "Routing Mapping"),
                    new ShellActionDefinition("show-reason-mapping", "Reason Mapping"),
                    new ShellActionDefinition("show-ai-process-classifier", "AI Process Classifier"),
                    new ShellActionDefinition("ai-select-training", "Select Training"),
                    new ShellActionDefinition("ai-train-model", "Train Model"),
                    new ShellActionDefinition("ai-run-inference", "Run Inference"),
                    new ShellActionDefinition("ai-export-output", "Export Output"),
                    new ShellActionDefinition("ai-import-feedback", "Import Feedback"),
                    new ShellActionDefinition("ai-retrain-feedback", "Retrain With Feedback"),
                    new ShellActionDefinition("ai-reason-select-training", "Select Reason Training"),
                    new ShellActionDefinition("ai-reason-train-model", "Train Reason Model"),
                    new ShellActionDefinition("ai-reason-run-inference", "Run Reason Inference"),
                    new ShellActionDefinition("ai-reason-export-output", "Export Reason Output"),
                    new ShellActionDefinition("ai-reason-import-feedback", "Import Reason Feedback"),
                    new ShellActionDefinition("ai-reason-retrain-feedback", "Retrain Reason Feedback"),
                    new ShellActionDefinition("mapping-export", "Export Mapping"),
                    new ShellActionDefinition("mapping-import", "Import Mapping"),
                    new ShellActionDefinition("clear-log", "Clear Log")
                },
                GetDataMakerWindow,
                window => window.GetWebModuleSnapshot(),
                async (window, action) => await window.InvokeWebModuleActionAsync(action),
                (window, handler) => window.WebModuleSnapshotChanged += handler,
                (window, handler) => window.WebModuleSnapshotChanged -= handler)));

        RegisterGraphModules();
        RegisterUtilityModules();

        foreach (ShellModuleDefinition module in _modules.Values)
        {
            if (module.Backend is not null)
            {
                module.Backend.SnapshotChanged += HandleModuleSnapshotChanged;
            }
        }
    }

    private void RegisterGraphModules()
    {
        RegisterModule(BuildGraphModule("GraphMaker:ScatterPlot", "SPL Graph Plot", "Scatter plot workspace.", "Scatter", "Scatter plot workflow.", new[] { "Graph setup and report actions are surfaced in the shell." }, GetScatterPlotView, new[] { new ShellActionDefinition("browse-files", "Browse"), new ShellActionDefinition("remove-selected-file", "Remove Selected"), new ShellActionDefinition("save-report", "Save Report"), new ShellActionDefinition("load-report", "Load Report"), new ShellActionDefinition("generate-graph", "Generate Graph", true) }, view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler));
        RegisterModule(BuildGraphModule("GraphMaker:ProcessTrend", "Process Trend", "Trend analysis workspace.", "Trend", "Process trend workflow.", new[] { "Selected files, ranges, and generation flow are reflected in the shell." }, GetProcessTrendView, new[] { new ShellActionDefinition("browse-files", "Browse"), new ShellActionDefinition("remove-selected-file", "Remove Selected"), new ShellActionDefinition("save-report", "Save Report"), new ShellActionDefinition("load-report", "Load Report"), new ShellActionDefinition("generate-graph", "Generate Graph", true) }, view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler));
        RegisterModule(BuildGraphModule("GraphMaker:UnifiedMultiY", "Single X", "Unified multi-series graph workflow.", "MultiY", "Single X graph workflow.", new[] { "File, header, column, and output setup are reflected in the shell." }, GetUnifiedMultiYView, new[] { new ShellActionDefinition("browse-file", "Browse"), new ShellActionDefinition("save-config", "Save Config"), new ShellActionDefinition("load-config", "Load Config"), new ShellActionDefinition("generate-graph", "Generate Graph", true) }, view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler));
        RegisterModule(BuildGraphModule("GraphMaker:DailyDataTrendExtra", "Multi X", "Expanded daily trend visualization.", "Trend", "Multi X graph workflow.", new[] { "File list, metric selection, and chart generation are reflected in the shell." }, GetDailyTrendExtraView, new[] { new ShellActionDefinition("browse-files", "Browse"), new ShellActionDefinition("remove-selected-file", "Remove Selected"), new ShellActionDefinition("save-config", "Save Config"), new ShellActionDefinition("load-config", "Load Config"), new ShellActionDefinition("generate-graph", "Generate Graph", true) }, view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler));
        RegisterModule(BuildGraphModule("GraphMaker:HeatMap", "Heat Map", "Heat map generation workflow.", "HeatMap", "Heat map workflow.", new[] { "Source path and generation readiness are reflected in the shell." }, GetHeatMapView, new[] { new ShellActionDefinition("browse-file", "Browse"), new ShellActionDefinition("load-template", "Load Template"), new ShellActionDefinition("generate-heatmap", "Generate Heat Map", true) }, view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler));
        RegisterModule(BuildGraphModule("GraphMaker:AudioBusData", "Audio Bus Data", "Audio bus visualization workflow.", "Audio", "Audio bus data workflow.", new[] { "Audio-bus file state and generation commands are reflected in the shell." }, GetAudioBusDataView, new[] { new ShellActionDefinition("browse-file", "Browse"), new ShellActionDefinition("load-config", "Load Config"), new ShellActionDefinition("generate-graph", "Generate Graph", true) }, view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler));
    }

    private void RegisterUtilityModules()
    {
        RegisterModule(new ShellModuleDefinition("DiskTree", string.Empty, "DiskTree", "File tree inspection and folder analysis.", "Storage", "DiskTree scan and compare workflow.", new[] { "Root folder, duplicate state, and progress are reflected in the shell." }, () => GetHostedWindowContent(GetDiskTreeWindow(), ref _diskTreeContent), new DelegateModuleBackend<global::DiskTree.MainWindow>("DiskTree", new[] { new ShellActionDefinition("browse-root-folder", "Browse Root"), new ShellActionDefinition("scan-and-update", "Scan + Update", true), new ShellActionDefinition("browse-compare-file", "Browse Compare"), new ShellActionDefinition("find-identical", "Find Identical"), new ShellActionDefinition("collect-matches", "Collect To Temp") }, GetDiskTreeWindow, window => window.GetWebModuleSnapshot(), (window, action) => Task.FromResult<object?>(window.InvokeWebModuleAction(action)), (window, handler) => window.WebModuleSnapshotChanged += handler, (window, handler) => window.WebModuleSnapshotChanged -= handler)));
        RegisterModule(new ShellModuleDefinition("Memo", string.Empty, "Memo", "Quick notes and image-backed memo flow.", "Notes", "Memo workspace.", new[] { "Memo row count and selected row status are visible from the shell." }, () => _memoView ??= new MemoView(), new DelegateModuleBackend<MemoView>("Memo", new[] { new ShellActionDefinition("add-row", "Add Row", true), new ShellActionDefinition("delete-row", "Delete Row"), new ShellActionDefinition("save-memo", "Save") }, () => _memoView ??= new MemoView(), view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler)));
        RegisterModule(new ShellModuleDefinition("VideoConverter", string.Empty, "VideoConverter", "Conversion workflow for media files.", "Media", "Video conversion workflow.", new[] { "Queue, output, and encoding status are reflected in the shell." }, () => GetHostedWindowContent(GetVideoConverterWindow(), ref _videoConverterContent), new DelegateModuleBackend<global::VideoConverter.MainWindow>("VideoConverter", new[] { new ShellActionDefinition("add-files", "Add Files"), new ShellActionDefinition("add-folder", "Add Folder"), new ShellActionDefinition("remove-selected", "Remove Selected"), new ShellActionDefinition("clear-queue", "Clear"), new ShellActionDefinition("start-encoding", "Start Encoding", true), new ShellActionDefinition("cancel-encoding", "Cancel"), new ShellActionDefinition("browse-ffmpeg", "Browse FFmpeg"), new ShellActionDefinition("browse-output", "Browse Output") }, GetVideoConverterWindow, window => window.GetWebModuleSnapshot(), (window, action) => Task.FromResult<object?>(window.InvokeWebModuleAction(action)), (window, handler) => window.WebModuleSnapshotChanged += handler, (window, handler) => window.WebModuleSnapshotChanged -= handler)));
        RegisterModule(new ShellModuleDefinition("JsonEditor", string.Empty, "JsonEditor", "Structured editing for JSON content.", "Editor", "Json editor workspace.", new[] { "Open, save, save-as, and reload commands are available from the side panel." }, () => _jsonEditorView ??= new JsonEditorView(), new DelegateModuleBackend<JsonEditorView>("JsonEditor", new[] { new ShellActionDefinition("open-json", "Open JSON", true), new ShellActionDefinition("save-json", "Save"), new ShellActionDefinition("save-json-as", "Save As"), new ShellActionDefinition("reload-json", "Reload") }, () => _jsonEditorView ??= new JsonEditorView(), view => view.GetWebModuleSnapshot(), async (view, action) => await view.InvokeWebModuleActionAsync(action), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler)));
        RegisterModule(new ShellModuleDefinition("ScreenCapture", string.Empty, "ScreenCapture", "Capture flows and hotkey-based entry.", "Capture", "Screen capture workspace.", new[] { "Capture state and commands can be launched from the shell panel." }, () => _screenCaptureView ??= new ScreenCaptureView(), new DelegateModuleBackend<ScreenCaptureView>("ScreenCapture", new[] { new ShellActionDefinition("capture-full-screen", "Capture Full Screen", true), new ShellActionDefinition("capture-active-window", "Capture Active Window"), new ShellActionDefinition("capture-region", "Capture Region"), new ShellActionDefinition("copy-capture", "Copy"), new ShellActionDefinition("save-capture", "Save PNG"), new ShellActionDefinition("clear-ink", "Clear Ink") }, () => _screenCaptureView ??= new ScreenCaptureView(), view => view.GetWebModuleSnapshot(), async (view, action) => await view.InvokeWebModuleActionAsync(action), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler)));
        RegisterModule(new ShellModuleDefinition("FileTransfer", string.Empty, "File Transfer", "Transfer queue and activity dashboard.", "Transfer", "File transfer workspace.", new[] { "Queue count, transfer state, progress, and recent activity are exposed through the shell." }, () => _fileTransferView ??= new FileTransferView(), new DelegateModuleBackend<FileTransferView>("FileTransfer", new[] { new ShellActionDefinition("add-file", "Add File"), new ShellActionDefinition("remove-selected", "Remove Selected"), new ShellActionDefinition("start-transfer", "Start Transfer", true) }, () => _fileTransferView ??= new FileTransferView(), view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler)));
        RegisterModule(new ShellModuleDefinition("Translator", string.Empty, "Translator", "Translation workspace with source, result, glossary, and history panes.", "Language", "Translator workspace.", new[] { "Source text, result text, glossary count, and history count are reflected in the shell." }, () => _translatorView ??= new TranslatorView(), new DelegateModuleBackend<TranslatorView>("Translator", new[] { new ShellActionDefinition("translate-text", "Translate", true), new ShellActionDefinition("swap-text", "Swap"), new ShellActionDefinition("clear-text", "Clear") }, () => _translatorView ??= new TranslatorView(), view => view.GetWebModuleSnapshot(), (view, action) => Task.FromResult<object?>(view.InvokeWebModuleAction(action)), (view, handler) => view.WebModuleSnapshotChanged += handler, (view, handler) => view.WebModuleSnapshotChanged -= handler)));
        RegisterModule(new ShellModuleDefinition("Settings", string.Empty, "Settings", "Host-level app configuration.", "Config", "Shared settings workspace.", new[] { "Settings remains strongly coupled to several concrete modules." }, () => _appSettingsView ??= new AppSettingsView()));
    }

    private void RegisterModule(ShellModuleDefinition definition)
    {
        _modules[definition.Target] = definition;
    }

    private ShellModuleDefinition BuildGraphModule<TView>(
        string target,
        string title,
        string summary,
        string tag,
        string detail,
        IReadOnlyList<string> notes,
        Func<TView> getView,
        IReadOnlyList<ShellActionDefinition> actions,
        Func<TView, object?> getSnapshot,
        Func<TView, string, Task<object?>> invokeAction,
        Action<TView, Action> subscribe,
        Action<TView, Action> unsubscribe)
        where TView : FrameworkElement
    {
        return new ShellModuleDefinition(
            target,
            string.Empty,
            title,
            summary,
            tag,
            detail,
            notes,
            () =>
            {
                EnsureGraphMakerResourcesLoaded();
                return getView();
            },
            new DelegateModuleBackend<TView>(
                target,
                actions,
                () =>
                {
                    EnsureGraphMakerResourcesLoaded();
                    return getView();
                },
                getSnapshot,
                invokeAction,
                subscribe,
                unsubscribe));
    }

    private void BuildMenu()
    {
        MenuHost.Children.Clear();
        _menuButtons.Clear();
        AddMenuButton("DataMaker");
        AddGraphMakerGroup();
        AddMenuButton("DiskTree");
        AddMenuButton("Memo");
        AddMenuButton("VideoConverter");
        AddMenuButton("JsonEditor");
        AddMenuButton("ScreenCapture");
        AddMenuButton("FileTransfer");
        AddMenuButton("Translator");
        AddMenuButton("Settings");
    }

    private void SelectModule(string target)
    {
        if (!_modules.TryGetValue(target, out ShellModuleDefinition? module))
        {
            return;
        }

        if (IsGraphMakerTarget(target))
        {
            SetGraphMakerExpanded(true);
        }

        _activeTarget = target;
        CurrentModuleTitle.Text = module.Title;
        CurrentModuleDescription.Text = module.Detail;
        EmbeddedModuleHost.Content = module.CreateContent();

        UpdateMenuVisualState();
    }

    private void UpdateMenuVisualState()
    {
        foreach ((string target, Button button) in _menuButtons)
        {
            bool selected = string.Equals(target, _activeTarget, StringComparison.Ordinal) ||
                            (string.Equals(target, GraphMakerGroupKey, StringComparison.Ordinal) && IsGraphMakerTarget(_activeTarget));
            button.Background = selected ? new SolidColorBrush(Color.FromRgb(255, 229, 204)) : Brushes.Transparent;
            button.BorderBrush = selected ? new SolidColorBrush(Color.FromRgb(255, 191, 138)) : Brushes.Transparent;
        }
    }

    private void AddMenuButton(string target, Thickness? margin = null, bool isChild = false)
    {
        ShellModuleDefinition module = _modules[target];
        string icon = GetMenuIcon(target);
        Button button = new()
        {
            Content = BuildMenuButtonContent(icon, module.Title, isChild),
            Margin = margin ?? new Thickness(0, 0, 0, 4),
            HorizontalContentAlignment = HorizontalAlignment.Left,
            VerticalContentAlignment = VerticalAlignment.Center,
            Background = Brushes.Transparent,
            Foreground = new SolidColorBrush(Color.FromRgb(31, 41, 55)),
            BorderBrush = Brushes.Transparent,
            BorderThickness = new Thickness(1),
            ToolTip = module.Summary,
            Style = (Style)FindResource("MenuNavButtonStyle")
        };

        button.Click += (_, _) => SelectModule(target);
        _menuButtons[target] = button;

        if (isChild && _graphMakerChildrenHost is not null)
        {
            _graphMakerChildrenHost.Children.Add(button);
            return;
        }

        MenuHost.Children.Add(button);
    }

    private void AddGraphMakerGroup()
    {
        Button parentButton = new()
        {
            Content = BuildGraphMakerGroupContent(),
            Margin = new Thickness(0, 0, 0, 4),
            HorizontalContentAlignment = HorizontalAlignment.Left,
            VerticalContentAlignment = VerticalAlignment.Center,
            Background = Brushes.Transparent,
            Foreground = new SolidColorBrush(Color.FromRgb(31, 41, 55)),
            BorderBrush = Brushes.Transparent,
            BorderThickness = new Thickness(1),
            ToolTip = "GraphMaker modules",
            Style = (Style)FindResource("MenuNavButtonStyle")
        };

        parentButton.Click += (_, _) => SetGraphMakerExpanded(!_isGraphMakerExpanded);
        MenuHost.Children.Add(parentButton);
        _menuButtons[GraphMakerGroupKey] = parentButton;

        _graphMakerChildrenHost = new StackPanel
        {
            Margin = new Thickness(14, 0, 0, 4)
        };

        MenuHost.Children.Add(_graphMakerChildrenHost);

        foreach (string target in _graphMakerTargets)
        {
            AddMenuButton(target, new Thickness(0, 0, 0, 3), isChild: true);
        }

        SetGraphMakerExpanded(false);
    }

    private FrameworkElement BuildGraphMakerGroupContent()
    {
        Grid grid = new();
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
        grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

        FrameworkElement mainContent = BuildMenuButtonContent(GetMenuIcon(GraphMakerGroupKey), "GraphMaker", isChild: false);
        Grid.SetColumn(mainContent, 0);
        Grid.SetColumnSpan(mainContent, 2);
        grid.Children.Add(mainContent);

        _graphMakerChevron = new TextBlock
        {
            Text = "▸",
            Margin = new Thickness(10, 0, 2, 0),
            VerticalAlignment = VerticalAlignment.Center,
            Foreground = new SolidColorBrush(Color.FromRgb(107, 124, 147)),
            FontSize = 14,
            FontWeight = FontWeights.SemiBold
        };

        Grid.SetColumn(_graphMakerChevron, 2);
        grid.Children.Add(_graphMakerChevron);
        return grid;
    }

    private static FrameworkElement BuildMenuButtonContent(string icon, string title, bool isChild)
    {
        StackPanel panel = new()
        {
            Orientation = Orientation.Horizontal
        };

        Border iconBadge = new()
        {
            Width = isChild ? 22 : 24,
            Height = isChild ? 22 : 24,
            CornerRadius = new CornerRadius(isChild ? 7 : 8),
            Background = new SolidColorBrush(Color.FromRgb(229, 239, 252)),
            BorderBrush = new SolidColorBrush(Color.FromRgb(201, 218, 241)),
            BorderThickness = new Thickness(1),
            VerticalAlignment = VerticalAlignment.Center,
            Child = BuildMenuIcon(icon)
        };

        TextBlock titleBlock = new()
        {
            Text = title,
            Margin = new Thickness(isChild ? 7 : 8, 0, 0, 0),
            TextWrapping = TextWrapping.Wrap,
            Foreground = new SolidColorBrush(Color.FromRgb(31, 41, 55)),
            FontSize = isChild ? 12 : 13,
            FontWeight = isChild ? FontWeights.Medium : FontWeights.SemiBold,
            VerticalAlignment = VerticalAlignment.Center
        };

        panel.Children.Add(iconBadge);
        panel.Children.Add(titleBlock);
        return panel;
    }

    private static FrameworkElement BuildMenuIcon(string icon)
    {
        Geometry geometry = Geometry.Parse(icon);
        return new Viewbox
        {
            Width = 13,
            Height = 13,
            Stretch = Stretch.Uniform,
            Child = new System.Windows.Shapes.Path
            {
                Data = geometry,
                Fill = new SolidColorBrush(Color.FromRgb(45, 85, 127)),
                Stretch = Stretch.Uniform,
                Width = 13,
                Height = 13,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            }
        };
    }

    private static string GetMenuIcon(string target)
    {
        return target switch
        {
            "DataMaker" => "F1 M 3,2 L 11,2 L 11,4 L 3,4 Z M 4,5 L 10,5 L 10,6.5 L 4,6.5 Z M 4,7.5 L 10,7.5 L 10,9 L 4,9 Z M 4,10 L 8,10 L 8,11.5 L 4,11.5 Z",
            GraphMakerGroupKey => "F1 M 2,10 L 3.2,10 L 3.2,11.5 L 2,11.5 Z M 5,7 L 6.2,7 L 6.2,8.2 L 5,8.2 Z M 8.2,4.2 L 9.4,4.2 L 9.4,5.4 L 8.2,5.4 Z M 10.2,8.5 A 1.2,1.2 0 1 1 10.21,8.5 Z",
            "GraphMaker:ScatterPlot" => "F1 M 2,10 L 3.2,10 L 3.2,11.5 L 2,11.5 Z M 5,7 L 6.2,7 L 6.2,8.2 L 5,8.2 Z M 8.2,4.2 L 9.4,4.2 L 9.4,5.4 L 8.2,5.4 Z M 10.2,8.5 A 1.2,1.2 0 1 1 10.21,8.5 Z",
            "GraphMaker:ProcessTrend" => "F1 M 2,9.5 L 4.5,7 L 6.5,8 L 9.5,4.5 L 10.5,5.5 L 6.7,9.3 L 4.4,8.2 L 2.8,9.8 Z",
            "GraphMaker:UnifiedMultiY" => "F1 M 2,10 L 5,6.5 L 7,8.2 L 10,3.5 L 11,4.3 L 7.1,10 L 5,8.3 L 2.7,10.8 Z M 2,3 L 2,4.2 L 4.8,4.2 L 4.8,3 Z",
            "GraphMaker:DailyDataTrendExtra" => "F1 M 2,9 L 4,6 L 6,7.2 L 8,4.5 L 10,6.2 L 11,5.1 L 8.1,3 L 6,5.7 L 4,4.5 L 2,7.4 Z",
            "GraphMaker:HeatMap" => "F1 M 2,2 L 5,2 L 5,5 L 2,5 Z M 5.5,2 L 8.5,2 L 8.5,5 L 5.5,5 Z M 9,2 L 12,2 L 12,5 L 9,5 Z M 2,5.5 L 5,5.5 L 5,8.5 L 2,8.5 Z M 5.5,5.5 L 8.5,5.5 L 8.5,8.5 L 5.5,8.5 Z M 9,5.5 L 12,5.5 L 12,8.5 L 9,8.5 Z M 2,9 L 5,9 L 5,12 L 2,12 Z M 5.5,9 L 8.5,9 L 8.5,12 L 5.5,12 Z M 9,9 L 12,9 L 12,12 L 9,12 Z",
            "GraphMaker:AudioBusData" => "F1 M 2,8.5 L 3.2,8.5 L 3.2,10.5 L 2,10.5 Z M 4.2,6.5 L 5.4,6.5 L 5.4,10.5 L 4.2,10.5 Z M 6.4,4.5 L 7.6,4.5 L 7.6,10.5 L 6.4,10.5 Z M 8.6,6 L 9.8,6 L 9.8,10.5 L 8.6,10.5 Z",
            "DiskTree" => "F1 M 2,3 L 6,3 L 7,4.2 L 12,4.2 L 12,10.8 L 2,10.8 Z",
            "Memo" => "F1 M 3,2 L 10,2 L 10,12 L 3,12 Z M 4.2,4 L 8.8,4 L 8.8,5 L 4.2,5 Z M 4.2,6 L 8.8,6 L 8.8,7 L 4.2,7 Z M 4.2,8 L 7.4,8 L 7.4,9 L 4.2,9 Z",
            "VideoConverter" => "F1 M 3,3 L 9.5,6.8 L 3,10.6 Z M 9.7,4.2 L 12,4.2 L 12,9.4 L 9.7,9.4 Z",
            "JsonEditor" => "F1 M 4.2,2.5 L 2,7 L 4.2,11.5 L 5.4,10.9 L 3.6,7 L 5.4,3.1 Z M 9.8,2.5 L 8.6,3.1 L 10.4,7 L 8.6,10.9 L 9.8,11.5 L 12,7 Z",
            "ScreenCapture" => "F1 M 2,4 L 4.2,4 L 4.2,5.2 L 3.2,5.2 L 3.2,10.8 L 8.8,10.8 L 8.8,9.8 L 10,9.8 L 10,12 L 2,12 Z M 6,2 L 12,2 L 12,8 L 10.8,8 L 10.8,4.1 L 6,4.1 Z",
            "FileTransfer" => "F1 M 2,6 L 7.5,6 L 7.5,4.2 L 11.5,7.1 L 7.5,10 L 7.5,8 L 2,8 Z",
            "Translator" => "F1 M 3,3 L 9,3 L 9,4.2 L 6.8,4.2 L 6.8,5.5 L 8.7,5.5 L 8.7,6.7 L 6.8,6.7 L 6.8,9.5 L 5.3,9.5 L 5.3,6.7 L 3,6.7 L 3,5.5 L 5.3,5.5 L 5.3,4.2 L 3,4.2 Z M 9.4,7.2 L 12,10.8 L 10.8,10.8 L 10.2,10 L 8.8,10 L 8.2,10.8 L 7,10.8 L 9.6,7.2 Z",
            "Settings" => "F1 M 7,2.2 L 8,2.2 L 8.3,3.5 L 9.4,3.9 L 10.5,3.2 L 11.2,3.9 L 10.5,5 L 10.9,6.1 L 12.2,6.4 L 12.2,7.4 L 10.9,7.7 L 10.5,8.8 L 11.2,9.9 L 10.5,10.6 L 9.4,9.9 L 8.3,10.3 L 8,11.6 L 7,11.6 L 6.7,10.3 L 5.6,9.9 L 4.5,10.6 L 3.8,9.9 L 4.5,8.8 L 4.1,7.7 L 2.8,7.4 L 2.8,6.4 L 4.1,6.1 L 4.5,5 L 3.8,3.9 L 4.5,3.2 L 5.6,3.9 L 6.7,3.5 Z M 7.5,5.3 A 1.7,1.7 0 1 1 7.51,5.3 Z",
            _ => "F1 M 3,3 L 11,3 L 11,11 L 3,11 Z"
        };
    }

    private void SetGraphMakerExpanded(bool expanded)
    {
        _isGraphMakerExpanded = expanded;

        if (_graphMakerChildrenHost is not null)
        {
            _graphMakerChildrenHost.Visibility = expanded ? Visibility.Visible : Visibility.Collapsed;
        }

        if (_graphMakerChevron is not null)
        {
            _graphMakerChevron.Text = expanded ? "▾" : "▸";
        }
    }

    private bool IsGraphMakerTarget(string? target) =>
        !string.IsNullOrWhiteSpace(target) && _graphMakerTargets.Contains(target);

    private void HandleModuleSnapshotChanged()
    {
    }

    protected override void OnClosed(EventArgs e)
    {
        CloseAuxiliaryWindows();
        base.OnClosed(e);

        if (Application.Current is not null)
        {
            Application.Current.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new Action(() =>
                {
                    if (Application.Current is not null)
                    {
                        Application.Current.Shutdown();
                    }
                }));
        }
    }

    private global::DataMaker.MainWindow GetDataMakerWindow() => _dataMakerWindow ??= new global::DataMaker.MainWindow();
    private global::DiskTree.MainWindow GetDiskTreeWindow() => _diskTreeWindow ??= new global::DiskTree.MainWindow();
    private global::VideoConverter.MainWindow GetVideoConverterWindow() => _videoConverterWindow ??= new global::VideoConverter.MainWindow();

    private global::GraphMaker.ScatterPlotView GetScatterPlotView()
    {
        EnsureGraphMakerResourcesLoaded();
        return _scatterPlotView ??= new global::GraphMaker.ScatterPlotView();
    }

    private global::GraphMaker.ProcessFlowTrendView GetProcessTrendView()
    {
        EnsureGraphMakerResourcesLoaded();
        return _processTrendView ??= new global::GraphMaker.ProcessFlowTrendView();
    }

    private global::GraphMaker.UnifiedMultiYView GetUnifiedMultiYView()
    {
        EnsureGraphMakerResourcesLoaded();
        return _unifiedMultiYView ??= new global::GraphMaker.UnifiedMultiYView();
    }

    private global::GraphMaker.DailyDataTrendExtraView GetDailyTrendExtraView()
    {
        EnsureGraphMakerResourcesLoaded();
        return _dailyTrendExtraView ??= new global::GraphMaker.DailyDataTrendExtraView();
    }

    private global::GraphMaker.HeatMapView GetHeatMapView()
    {
        EnsureGraphMakerResourcesLoaded();
        return _heatMapView ??= new global::GraphMaker.HeatMapView();
    }

    private global::GraphMaker.AudioBusDataView GetAudioBusDataView()
    {
        EnsureGraphMakerResourcesLoaded();
        return _audioBusDataView ??= new global::GraphMaker.AudioBusDataView();
    }

    private static FrameworkElement GetHostedWindowContent(Window window, ref FrameworkElement? cache)
    {
        if (cache is not null)
        {
            return cache;
        }

        if (window.Content is not FrameworkElement content)
        {
            throw new InvalidOperationException($"Window '{window.GetType().FullName}' does not expose a FrameworkElement content root.");
        }

        window.Content = null;
        content.HorizontalAlignment = HorizontalAlignment.Stretch;
        content.VerticalAlignment = VerticalAlignment.Stretch;
        cache = content;
        return cache;
    }

    private static void EnsureGraphMakerResourcesLoaded()
    {
        if (Application.Current is null)
        {
            return;
        }

        bool alreadyLoaded = Application.Current.Resources.MergedDictionaries.Any(dictionary =>
            dictionary.Source is not null &&
            string.Equals(dictionary.Source.OriginalString, GraphMakerThemeUri.OriginalString, StringComparison.OrdinalIgnoreCase));

        if (!alreadyLoaded)
        {
            Application.Current.Resources.MergedDictionaries.Add(new ResourceDictionary
            {
                Source = GraphMakerThemeUri
            });
        }
    }

    private void CloseAuxiliaryWindows()
    {
        CloseWindowInstance(_dataMakerWindow);
        CloseWindowInstance(_diskTreeWindow);
        CloseWindowInstance(_videoConverterWindow);

        if (Application.Current is null)
        {
            return;
        }

        foreach (Window window in Application.Current.Windows.OfType<Window>().ToList())
        {
            if (ReferenceEquals(window, this))
            {
                continue;
            }

            CloseWindowInstance(window);
        }
    }

    private static void CloseWindowInstance(Window? window)
    {
        if (window is null)
        {
            return;
        }

        try
        {
            window.Close();
        }
        catch
        {
            // Best-effort shutdown cleanup for hosted/hidden windows.
        }
    }
}
