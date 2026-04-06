using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Diagnostics;

namespace WorkbenchHost.Modules.ScreenCapture;

public partial class CaptureHotkeyDialog : Window
{
    private readonly DispatcherTimer _modifierMonitorTimer;
    private readonly string _baseTitle;
    private readonly string _defaultHintText;
    private bool _hasCompleteGesture;

    public CaptureHotkey? SelectedHotkey { get; private set; }

    public CaptureHotkeyDialog(CaptureCommand command, CaptureHotkey currentHotkey)
    {
        InitializeComponent();
        _baseTitle = $"Set shortcut for {command.GetDisplayName()}";
        _defaultHintText = "Press a modifier key combination like Ctrl + Shift + R";
        TitleTextBlock.Text = _baseTitle;
        SelectedHotkey = currentHotkey;
        SetGestureDisplay(currentHotkey.ToDisplayString());
        HintTextBlock.Text = currentHotkey.IsEmpty
            ? _defaultHintText
            : $"Current shortcut: {currentHotkey.ToDisplayString()}";

        _modifierMonitorTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(50)
        };
        _modifierMonitorTimer.Tick += ModifierMonitorTimer_Tick;
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        Debug.WriteLine("[CaptureHotkeyDialog] Window loaded");
        Dispatcher.BeginInvoke(new Action(() =>
        {
            Debug.WriteLine("[CaptureHotkeyDialog] Applying initial focus to GestureTextBox");
            Activate();
            Focus();
            GestureDisplayBorder.Focus();
            Keyboard.Focus(GestureDisplayBorder);
            _modifierMonitorTimer.Start();
        }), DispatcherPriority.Input);
    }

    private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        Debug.WriteLine(
            $"[CaptureHotkeyDialog] PreviewKeyDown key={e.Key}, systemKey={e.SystemKey}, " +
            $"keyboardModifiers={Keyboard.Modifiers}, source={e.OriginalSource?.GetType().Name}");

        if (e.Key == Key.Escape)
        {
            DialogResult = false;
            Close();
            return;
        }

        if (e.Key == Key.Back)
        {
            SelectedHotkey = CaptureHotkey.None;
            SetGestureDisplay("Disabled");
            HintTextBlock.Text = "Shortcut disabled. Press a new combination or save.";
            _hasCompleteGesture = false;
            e.Handled = true;
            return;
        }

        Key key = e.Key == Key.System ? e.SystemKey : e.Key;
        ModifierKeys modifiers = GetEffectiveModifiers(key, Keyboard.Modifiers);
        Debug.WriteLine(
            $"[CaptureHotkeyDialog] Resolved key={key}, effectiveModifiers={modifiers}, " +
            $"isModifierOnly={IsModifierKey(key)}");

        if (IsModifierKey(key))
        {
            string modifierText = GetModifierDisplayText(modifiers);
            SetGestureDisplay(string.IsNullOrWhiteSpace(modifierText) ? "Waiting..." : modifierText);
            HintTextBlock.Text = $"Modifiers: {modifierText}";
            Debug.WriteLine($"[CaptureHotkeyDialog] Modifier-only input displayed as '{GestureDisplayTextBlock.Text}'");
            e.Handled = true;
            return;
        }

        if (modifiers == ModifierKeys.None)
        {
            SetGestureDisplay("Use Ctrl, Alt, Shift, or Win");
            HintTextBlock.Text = "Single keys are not supported for global capture shortcuts.";
            _hasCompleteGesture = false;
            Debug.WriteLine("[CaptureHotkeyDialog] Rejected single key without modifiers");
            e.Handled = true;
            return;
        }

        SelectedHotkey = new CaptureHotkey(modifiers, key);
        SetGestureDisplay(SelectedHotkey.Value.ToDisplayString());
        HintTextBlock.Text = "Press Save to apply this shortcut.";
        _hasCompleteGesture = true;
        Debug.WriteLine($"[CaptureHotkeyDialog] Captured complete gesture '{GestureDisplayTextBlock.Text}'");
        e.Handled = true;
    }

    private void Window_PreviewKeyUp(object sender, KeyEventArgs e)
    {
        Debug.WriteLine(
            $"[CaptureHotkeyDialog] PreviewKeyUp key={e.Key}, systemKey={e.SystemKey}, " +
            $"keyboardModifiers={Keyboard.Modifiers}, source={e.OriginalSource?.GetType().Name}");

        if (Keyboard.Modifiers == ModifierKeys.None && !_hasCompleteGesture)
        {
            SetGestureDisplay("Waiting...");
            HintTextBlock.Text = _defaultHintText;
            Debug.WriteLine("[CaptureHotkeyDialog] Reset display to Waiting...");
        }
    }

    private void Window_Closed(object sender, System.EventArgs e)
    {
        Debug.WriteLine("[CaptureHotkeyDialog] Window closed");
        _modifierMonitorTimer.Stop();
        _modifierMonitorTimer.Tick -= ModifierMonitorTimer_Tick;
    }

    private void GestureDisplayBorder_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        Debug.WriteLine("[CaptureHotkeyDialog] Gesture display clicked, forcing focus");
        GestureDisplayBorder.Focus();
        Keyboard.Focus(GestureDisplayBorder);
        e.Handled = true;
    }

    private void ClearButton_Click(object sender, RoutedEventArgs e)
    {
        SelectedHotkey = CaptureHotkey.None;
        SetGestureDisplay("Disabled");
        HintTextBlock.Text = "Shortcut disabled. Press a new combination or save.";
        _hasCompleteGesture = false;
    }

    private void SaveButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }

    private static bool IsModifierKey(Key key)
    {
        return key is Key.LeftCtrl or Key.RightCtrl
            or Key.LeftAlt or Key.RightAlt
            or Key.LeftShift or Key.RightShift
            or Key.LWin or Key.RWin;
    }

    private static string GetModifierDisplayText(ModifierKeys modifiers)
    {
        if (modifiers == ModifierKeys.None)
        {
            return string.Empty;
        }

        var parts = new System.Collections.Generic.List<string>();
        if (modifiers.HasFlag(ModifierKeys.Control))
        {
            parts.Add("Ctrl");
        }

        if (modifiers.HasFlag(ModifierKeys.Shift))
        {
            parts.Add("Shift");
        }

        if (modifiers.HasFlag(ModifierKeys.Alt))
        {
            parts.Add("Alt");
        }

        if (modifiers.HasFlag(ModifierKeys.Windows))
        {
            parts.Add("Win");
        }

        return string.Join(" + ", parts);
    }

    private static ModifierKeys GetEffectiveModifiers(Key key, ModifierKeys currentModifiers)
    {
        ModifierKeys effective = currentModifiers;

        switch (key)
        {
            case Key.LeftCtrl:
            case Key.RightCtrl:
                effective |= ModifierKeys.Control;
                break;
            case Key.LeftShift:
            case Key.RightShift:
                effective |= ModifierKeys.Shift;
                break;
            case Key.LeftAlt:
            case Key.RightAlt:
                effective |= ModifierKeys.Alt;
                break;
            case Key.LWin:
            case Key.RWin:
                effective |= ModifierKeys.Windows;
                break;
        }

        return effective;
    }

    private void ModifierMonitorTimer_Tick(object? sender, System.EventArgs e)
    {
        Debug.WriteLine(
            $"[CaptureHotkeyDialog] ModifierTimer modifiers={Keyboard.Modifiers}, " +
            $"hasCompleteGesture={_hasCompleteGesture}, display='{GestureDisplayTextBlock.Text}'");

        if (_hasCompleteGesture)
        {
            return;
        }

        string modifierText = GetModifierDisplayText(Keyboard.Modifiers);
        if (string.IsNullOrWhiteSpace(modifierText))
        {
            if (GestureDisplayTextBlock.Text != "Waiting..." && GestureDisplayTextBlock.Text != "Disabled")
            {
                SetGestureDisplay("Waiting...");
                HintTextBlock.Text = _defaultHintText;
            }

            return;
        }

        SetGestureDisplay(modifierText);
        HintTextBlock.Text = $"Modifiers: {modifierText}";
    }

    private void SetGestureDisplay(string text)
    {
        GestureDisplayTextBlock.Text = text;
        Title = string.IsNullOrWhiteSpace(text) || text == "Waiting..."
            ? "Set Capture Shortcut"
            : $"Set Capture Shortcut [{text}]";
    }
}
