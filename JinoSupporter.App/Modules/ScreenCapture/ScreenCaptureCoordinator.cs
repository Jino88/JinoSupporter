using System;
using System.Collections.Generic;
using System.Windows.Input;

namespace WorkbenchHost.Modules.ScreenCapture;

public static class ScreenCaptureCoordinator
{
    private static IReadOnlyDictionary<CaptureCommand, CaptureHotkey> _hotkeys = CaptureHotkeyManager.Load();

    public static event Action<CaptureCommand>? CaptureRequested;
    public static event Action<IReadOnlyDictionary<CaptureCommand, CaptureHotkey>>? HotkeysChanged;

    public static IReadOnlyDictionary<CaptureCommand, CaptureHotkey> Hotkeys => _hotkeys;

    public static void RequestCapture(CaptureCommand command)
    {
        CaptureRequested?.Invoke(command);
    }

    public static void UpdateHotkeys(IReadOnlyDictionary<CaptureCommand, CaptureHotkey> hotkeys)
    {
        _hotkeys = hotkeys;
        CaptureHotkeyManager.Save(hotkeys);
        HotkeysChanged?.Invoke(_hotkeys);
    }
}
