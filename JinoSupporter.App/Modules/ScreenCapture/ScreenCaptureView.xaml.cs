using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Ink;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WorkbenchHost.Modules.ScreenCapture;

public partial class ScreenCaptureView : UserControl, INotifyPropertyChanged
{
    private readonly List<CaptureColorOption> _colors =
    [
        new("Red", Colors.IndianRed),
        new("Blue", Colors.SteelBlue),
        new("Green", Colors.SeaGreen),
        new("Yellow", Colors.Goldenrod),
        new("White", Colors.White),
        new("Black", Colors.Black)
    ];

    private BitmapSource? _currentCapture;
    private string _statusMessage = "Ready to capture.";
    private string _captureInfoText = "No image loaded";
    private string _strokeWidthText = "3 px";
    private IReadOnlyDictionary<CaptureCommand, CaptureHotkey> _hotkeys = ScreenCaptureCoordinator.Hotkeys;
    private string _selectedTool = "pen";

    public event PropertyChangedEventHandler? PropertyChanged;
    public event Action? WebModuleSnapshotChanged;

    public string StatusMessage
    {
        get => _statusMessage;
        private set
        {
            if (_statusMessage == value)
            {
                return;
            }

            _statusMessage = value;
            OnPropertyChanged();
        }
    }

    public string CaptureInfoText
    {
        get => _captureInfoText;
        private set
        {
            if (_captureInfoText == value)
            {
                return;
            }

            _captureInfoText = value;
            OnPropertyChanged();
        }
    }

    public string StrokeWidthText
    {
        get => _strokeWidthText;
        private set
        {
            if (_strokeWidthText == value)
            {
                return;
            }

            _strokeWidthText = value;
            OnPropertyChanged();
        }
    }

    public string FullScreenHotkeyText => GetHotkeyText(CaptureCommand.FullScreen);

    public string ActiveWindowHotkeyText => GetHotkeyText(CaptureCommand.ActiveWindow);

    public string RegionHotkeyText => GetHotkeyText(CaptureCommand.Region);

    public string HotkeySummaryText =>
        $"Hotkeys: Full {FullScreenHotkeyText} | Window {ActiveWindowHotkeyText} | Region {RegionHotkeyText}";

    public ScreenCaptureView()
    {
        InitializeComponent();
        DataContext = this;

        ColorComboBox.ItemsSource = _colors;
        ColorComboBox.DisplayMemberPath = nameof(CaptureColorOption.Name);
        ColorComboBox.SelectedIndex = 0;

        AnnotationCanvas.DefaultDrawingAttributes = CreateDrawingAttributes(_colors[0].Color, (float)StrokeWidthSlider.Value, isHighlighter: false);
        AnnotationCanvas.EditingMode = InkCanvasEditingMode.Ink;
        PenToolToggle.IsChecked = true;
        UpdateSurfaceSize();
        ConfigureHotkeyContextMenus();
        ScreenCaptureCoordinator.CaptureRequested += OnCaptureRequested;
        ScreenCaptureCoordinator.HotkeysChanged += OnHotkeysChanged;
    }

    ~ScreenCaptureView()
    {
        ScreenCaptureCoordinator.CaptureRequested -= OnCaptureRequested;
        ScreenCaptureCoordinator.HotkeysChanged -= OnHotkeysChanged;
    }

    private void CaptureFullScreenButton_Click(object sender, RoutedEventArgs e)
    {
        RunCapture(CaptureCommand.FullScreen);
    }

    private void CaptureActiveWindowButton_Click(object sender, RoutedEventArgs e)
    {
        RunCapture(CaptureCommand.ActiveWindow);
    }

    private async void CaptureRegionButton_Click(object sender, RoutedEventArgs e)
    {
        await RunRegionCaptureAsync();
    }

    private void CopyButton_Click(object sender, RoutedEventArgs e)
    {
        BitmapSource? output = RenderOutputBitmap();
        if (output is null)
        {
            StatusMessage = "Nothing to copy yet.";
            return;
        }

        Clipboard.SetImage(output);
        StatusMessage = "Capture copied to clipboard.";
        NotifyWebModuleSnapshotChanged();
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        BitmapSource? output = RenderOutputBitmap();
        if (output is null)
        {
            StatusMessage = "Nothing to save yet.";
            return;
        }

        var dialog = new SaveFileDialog
        {
            Filter = "PNG image (*.png)|*.png",
            DefaultExt = ".png",
            AddExtension = true,
            FileName = $"capture_{DateTime.Now:yyyyMMdd_HHmmss}.png"
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        using var stream = File.Create(dialog.FileName);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(output));
        encoder.Save(stream);
        StatusMessage = $"Saved {Path.GetFileName(dialog.FileName)}";
        NotifyWebModuleSnapshotChanged();
    }

    private async System.Threading.Tasks.Task RunRegionCaptureAsync()
    {
        Window? hostWindow = Window.GetWindow(this);
        if (hostWindow is not null)
        {
            hostWindow.WindowState = WindowState.Minimized;
        }

        await System.Threading.Tasks.Task.Delay(180);

        Rectangle? selectedBounds = null;

        try
        {
            var overlay = new RegionCaptureOverlayWindow();
            bool? result = overlay.ShowDialog();
            if (result == true)
            {
                selectedBounds = overlay.SelectedBounds;
            }
        }
        finally
        {
            if (hostWindow is not null)
            {
                hostWindow.WindowState = WindowState.Normal;
                hostWindow.Activate();
            }
        }

        if (selectedBounds is not Rectangle bounds || bounds.Width <= 0 || bounds.Height <= 0)
        {
            StatusMessage = "Region capture canceled.";
            NotifyWebModuleSnapshotChanged();
            return;
        }

        CaptureFromRectangle(bounds, "Custom region");
    }

    private void ClearInkButton_Click(object sender, RoutedEventArgs e)
    {
        AnnotationCanvas.Strokes.Clear();
        StatusMessage = "Annotations cleared.";
        NotifyWebModuleSnapshotChanged();
    }

    private void PenToolToggle_Click(object sender, RoutedEventArgs e)
    {
        SetEditingMode(InkCanvasEditingMode.Ink, PenToolToggle);
    }

    private void HighlighterToolToggle_Click(object sender, RoutedEventArgs e)
    {
        SetEditingMode(InkCanvasEditingMode.Ink, HighlighterToolToggle, isHighlighter: true);
    }

    private void EraserToolToggle_Click(object sender, RoutedEventArgs e)
    {
        SetEditingMode(InkCanvasEditingMode.EraseByStroke, EraserToolToggle);
    }

    private void ColorComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        UpdateDrawingAttributes();
        NotifyWebModuleSnapshotChanged();
    }

    private void StrokeWidthSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
    {
        StrokeWidthText = $"{StrokeWidthSlider.Value:F0} px";
        UpdateDrawingAttributes();
        NotifyWebModuleSnapshotChanged();
    }

    private void CaptureFromRectangle(Rectangle bounds, string label)
    {
        if (bounds.Width <= 0 || bounds.Height <= 0)
        {
            StatusMessage = "Invalid capture bounds.";
            NotifyWebModuleSnapshotChanged();
            return;
        }

        try
        {
            using Bitmap bitmap = new(bounds.Width, bounds.Height, System.Drawing.Imaging.PixelFormat.Format32bppPArgb);
            using Graphics graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(bounds.Location, System.Drawing.Point.Empty, bounds.Size, CopyPixelOperation.SourceCopy);

            BitmapSource source = ConvertBitmap(bitmap);
            ApplyCapture(source, label);
        }
        catch (Exception ex)
        {
            StatusMessage = $"Capture failed: {ex.Message}";
            NotifyWebModuleSnapshotChanged();
        }
    }

    private void ApplyCapture(BitmapSource bitmap, string label)
    {
        _currentCapture = bitmap;
        CaptureImage.Source = bitmap;
        AnnotationCanvas.Strokes.Clear();
        EmptyStateText.Visibility = Visibility.Collapsed;
        UpdateSurfaceSize();
        Clipboard.SetImage(bitmap);
        CaptureInfoText = $"{label} · {bitmap.PixelWidth:N0} x {bitmap.PixelHeight:N0}";
        StatusMessage = $"{label} captured at {DateTime.Now:HH:mm:ss} and copied to clipboard.";
        NotifyWebModuleSnapshotChanged();
    }

    private void UpdateSurfaceSize()
    {
        double width = _currentCapture?.PixelWidth ?? 960;
        double height = _currentCapture?.PixelHeight ?? 540;

        CaptureSurface.Width = width;
        CaptureSurface.Height = height;
        AnnotationCanvas.Width = width;
        AnnotationCanvas.Height = height;
    }

    private void UpdateDrawingAttributes()
    {
        if (ColorComboBox.SelectedItem is not CaptureColorOption option)
        {
            return;
        }

        bool isHighlighter = HighlighterToolToggle.IsChecked == true && EraserToolToggle.IsChecked != true;
        AnnotationCanvas.DefaultDrawingAttributes = CreateDrawingAttributes(option.Color, (float)StrokeWidthSlider.Value, isHighlighter);
    }

    private void SetEditingMode(InkCanvasEditingMode mode, ToggleButton selectedButton, bool isHighlighter = false)
    {
        PenToolToggle.IsChecked = ReferenceEquals(selectedButton, PenToolToggle);
        HighlighterToolToggle.IsChecked = ReferenceEquals(selectedButton, HighlighterToolToggle);
        EraserToolToggle.IsChecked = ReferenceEquals(selectedButton, EraserToolToggle);
        _selectedTool = mode == InkCanvasEditingMode.EraseByStroke
            ? "eraser"
            : isHighlighter
                ? "highlighter"
                : "pen";

        AnnotationCanvas.EditingMode = mode;
        if (mode == InkCanvasEditingMode.Ink)
        {
            if (ColorComboBox.SelectedItem is CaptureColorOption option)
            {
                AnnotationCanvas.DefaultDrawingAttributes = CreateDrawingAttributes(option.Color, (float)StrokeWidthSlider.Value, isHighlighter);
            }
        }

        NotifyWebModuleSnapshotChanged();
    }

    private BitmapSource? RenderOutputBitmap()
    {
        if (_currentCapture is null)
        {
            return null;
        }

        CaptureSurface.Measure(new System.Windows.Size(CaptureSurface.Width, CaptureSurface.Height));
        CaptureSurface.Arrange(new Rect(0, 0, CaptureSurface.Width, CaptureSurface.Height));
        CaptureSurface.UpdateLayout();

        var renderBitmap = new RenderTargetBitmap(
            (int)CaptureSurface.Width,
            (int)CaptureSurface.Height,
            96,
            96,
            PixelFormats.Pbgra32);

        renderBitmap.Render(CaptureSurface);
        renderBitmap.Freeze();
        return renderBitmap;
    }

    private static DrawingAttributes CreateDrawingAttributes(System.Windows.Media.Color color, float width, bool isHighlighter)
    {
        return new DrawingAttributes
        {
            Color = color,
            Width = width,
            Height = width,
            FitToCurve = true,
            IgnorePressure = true,
            IsHighlighter = isHighlighter
        };
    }

    private static Rectangle GetVirtualScreenBounds()
    {
        int left = (int)SystemParameters.VirtualScreenLeft;
        int top = (int)SystemParameters.VirtualScreenTop;
        int width = (int)SystemParameters.VirtualScreenWidth;
        int height = (int)SystemParameters.VirtualScreenHeight;
        return new Rectangle(left, top, width, height);
    }

    private static bool TryGetForegroundWindowBounds(out Rectangle bounds)
    {
        IntPtr windowHandle = GetForegroundWindow();
        if (windowHandle == IntPtr.Zero || !GetWindowRect(windowHandle, out RECT rect))
        {
            bounds = Rectangle.Empty;
            return false;
        }

        bounds = Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
        return bounds.Width > 0 && bounds.Height > 0;
    }

    private static BitmapSource ConvertBitmap(Bitmap bitmap)
    {
        IntPtr hBitmap = bitmap.GetHbitmap();
        try
        {
            BitmapSource source = Imaging.CreateBitmapSourceFromHBitmap(
                hBitmap,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            source.Freeze();
            return source;
        }
        finally
        {
            DeleteObject(hBitmap);
        }
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private void RunCapture(CaptureCommand command)
    {
        switch (command)
        {
            case CaptureCommand.FullScreen:
                CaptureFromRectangle(GetVirtualScreenBounds(), "Full screen");
                break;
            case CaptureCommand.ActiveWindow:
                if (!TryGetForegroundWindowBounds(out Rectangle bounds))
                {
                    StatusMessage = "Could not detect the active window.";
                    return;
                }

                CaptureFromRectangle(bounds, "Active window");
                break;
            case CaptureCommand.Region:
                _ = RunRegionCaptureAsync();
                break;
        }
    }

    private void OnCaptureRequested(CaptureCommand command)
    {
        Dispatcher.BeginInvoke(new Action(() => RunCapture(command)));
    }

    private void OnHotkeysChanged(IReadOnlyDictionary<CaptureCommand, CaptureHotkey> hotkeys)
    {
        _hotkeys = hotkeys;
        OnPropertyChanged(nameof(FullScreenHotkeyText));
        OnPropertyChanged(nameof(ActiveWindowHotkeyText));
        OnPropertyChanged(nameof(RegionHotkeyText));
        OnPropertyChanged(nameof(HotkeySummaryText));
        NotifyWebModuleSnapshotChanged();
    }

    private void ConfigureHotkeyContextMenus()
    {
        AttachHotkeyMenu(CaptureFullScreenButton, CaptureCommand.FullScreen);
        AttachHotkeyMenu(CaptureActiveWindowButton, CaptureCommand.ActiveWindow);
        AttachHotkeyMenu(CaptureRegionButton, CaptureCommand.Region);
    }

    private void AttachHotkeyMenu(Button button, CaptureCommand command)
    {
        var menu = new System.Windows.Controls.ContextMenu();

        var setItem = new MenuItem { Header = "Set Shortcut..." };
        setItem.Click += (_, _) => ConfigureHotkey(command);
        menu.Items.Add(setItem);

        var disableItem = new MenuItem { Header = "Disable Shortcut" };
        disableItem.Click += (_, _) => DisableHotkey(command);
        menu.Items.Add(disableItem);

        button.ContextMenu = menu;
    }

    private void ConfigureHotkey(CaptureCommand command)
    {
        CaptureHotkey current = GetHotkey(command);
        var dialog = new CaptureHotkeyDialog(command, current)
        {
            Owner = Window.GetWindow(this)
        };

        if (dialog.ShowDialog() != true || dialog.SelectedHotkey is not CaptureHotkey selectedHotkey)
        {
            return;
        }

        if (!selectedHotkey.IsEmpty)
        {
            foreach ((CaptureCommand existingCommand, CaptureHotkey existingHotkey) in _hotkeys)
            {
                if (existingCommand != command && existingHotkey.Equals(selectedHotkey))
                {
                    MessageBox.Show(
                        Window.GetWindow(this),
                        $"'{selectedHotkey.ToDisplayString()}' is already assigned to {existingCommand.GetDisplayName()}.",
                        "Capture Shortcut",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
            }
        }

        var updated = new Dictionary<CaptureCommand, CaptureHotkey>(_hotkeys)
        {
            [command] = selectedHotkey
        };

        ScreenCaptureCoordinator.UpdateHotkeys(updated);
        StatusMessage = $"{command.GetDisplayName()} shortcut set to {selectedHotkey.ToDisplayString()}.";
        NotifyWebModuleSnapshotChanged();
    }

    private void DisableHotkey(CaptureCommand command)
    {
        var updated = new Dictionary<CaptureCommand, CaptureHotkey>(_hotkeys)
        {
            [command] = CaptureHotkey.None
        };

        ScreenCaptureCoordinator.UpdateHotkeys(updated);
        StatusMessage = $"{command.GetDisplayName()} shortcut disabled.";
        NotifyWebModuleSnapshotChanged();
    }

    private CaptureHotkey GetHotkey(CaptureCommand command)
    {
        return _hotkeys.TryGetValue(command, out CaptureHotkey hotkey) ? hotkey : CaptureHotkey.None;
    }

    private string GetHotkeyText(CaptureCommand command)
    {
        return GetHotkey(command).ToDisplayString();
    }

    public object GetWebModuleSnapshot()
    {
        string selectedColor = ColorComboBox.SelectedItem is CaptureColorOption option
            ? option.Name
            : _colors[0].Name;

        BitmapSource? previewBitmap = RenderOutputBitmap();

        return new
        {
            moduleType = "ScreenCapture",
            status = StatusMessage,
            captureInfo = CaptureInfoText,
            hotkeySummary = HotkeySummaryText,
            fullScreenHotkey = FullScreenHotkeyText,
            activeWindowHotkey = ActiveWindowHotkeyText,
            regionHotkey = RegionHotkeyText,
            strokeWidth = Math.Round(StrokeWidthSlider.Value),
            strokeWidthText = StrokeWidthText,
            selectedColor,
            selectedTool = _selectedTool,
            availableColors = _colors.Select(color => color.Name).ToArray(),
            hasCapture = previewBitmap is not null,
            imageDataUrl = previewBitmap is null ? string.Empty : ToDataUrl(previewBitmap)
        };
    }

    public object UpdateWebModuleState(JsonElement payload)
    {
        if (payload.ValueKind != JsonValueKind.Object)
        {
            return GetWebModuleSnapshot();
        }

        if (payload.TryGetProperty("selectedColor", out JsonElement selectedColorElement))
        {
            string requestedColor = selectedColorElement.GetString() ?? string.Empty;
            CaptureColorOption? requestedOption = _colors.FirstOrDefault(color =>
                string.Equals(color.Name, requestedColor, StringComparison.OrdinalIgnoreCase));
            if (requestedOption is not null)
            {
                ColorComboBox.SelectedItem = requestedOption;
            }
        }

        if (payload.TryGetProperty("strokeWidth", out JsonElement strokeWidthElement)
            && strokeWidthElement.TryGetDouble(out double strokeWidth))
        {
            StrokeWidthSlider.Value = Math.Clamp(strokeWidth, StrokeWidthSlider.Minimum, StrokeWidthSlider.Maximum);
        }

        if (payload.TryGetProperty("selectedTool", out JsonElement selectedToolElement))
        {
            switch ((selectedToolElement.GetString() ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "pen":
                    SetEditingMode(InkCanvasEditingMode.Ink, PenToolToggle);
                    break;
                case "highlighter":
                    SetEditingMode(InkCanvasEditingMode.Ink, HighlighterToolToggle, isHighlighter: true);
                    break;
                case "eraser":
                    SetEditingMode(InkCanvasEditingMode.EraseByStroke, EraserToolToggle);
                    break;
            }
        }

        NotifyWebModuleSnapshotChanged();
        return GetWebModuleSnapshot();
    }

    public async Task<object> InvokeWebModuleActionAsync(string action)
    {
        switch (action)
        {
            case "capture-full-screen":
                RunCapture(CaptureCommand.FullScreen);
                break;
            case "capture-active-window":
                RunCapture(CaptureCommand.ActiveWindow);
                break;
            case "capture-region":
                await RunRegionCaptureAsync();
                break;
            case "copy-capture":
                CopyButton_Click(this, new RoutedEventArgs());
                break;
            case "save-capture":
                SaveButton_Click(this, new RoutedEventArgs());
                break;
            case "clear-ink":
                ClearInkButton_Click(this, new RoutedEventArgs());
                break;
        }

        return GetWebModuleSnapshot();
    }

    private void NotifyWebModuleSnapshotChanged()
    {
        WebModuleSnapshotChanged?.Invoke();
    }

    private static string ToDataUrl(BitmapSource bitmapSource)
    {
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmapSource));
        using MemoryStream stream = new();
        encoder.Save(stream);
        return $"data:image/png;base64,{Convert.ToBase64String(stream.ToArray())}";
    }

    [DllImport("gdi32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DeleteObject(IntPtr hObject);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    private sealed record CaptureColorOption(string Name, System.Windows.Media.Color Color);
}
