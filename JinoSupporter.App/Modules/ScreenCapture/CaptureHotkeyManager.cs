using System;
using System.Collections.Generic;
using System.Windows.Input;
using WorkbenchHost.Infrastructure;

namespace WorkbenchHost.Modules.ScreenCapture;

public static class CaptureHotkeyManager
{
    public static string SettingsPath => WorkbenchSettingsStore.SettingsFilePath;

    public static IReadOnlyDictionary<CaptureCommand, CaptureHotkey> Load()
    {
        try
        {
            WorkbenchSettingsStore.WorkbenchSettings settings = WorkbenchSettingsStore.GetSettings();
            if (settings.CaptureHotkeys.Count == 0)
            {
                return GetDefaultHotkeys();
            }

            var result = new Dictionary<CaptureCommand, CaptureHotkey>();
            foreach ((string key, WorkbenchSettingsStore.CaptureHotkeySetting dto) in settings.CaptureHotkeys)
            {
                if (!Enum.TryParse(key, ignoreCase: true, out CaptureCommand command))
                {
                    continue;
                }

                if (Enum.TryParse(dto.Key, ignoreCase: true, out Key parsedKey))
                {
                    result[command] = new CaptureHotkey(dto.Modifiers, parsedKey);
                }
            }

            foreach ((CaptureCommand command, CaptureHotkey hotkey) in GetDefaultHotkeys())
            {
                result.TryAdd(command, hotkey);
            }

            return result;
        }
        catch
        {
            return GetDefaultHotkeys();
        }
    }

    public static void Save(IReadOnlyDictionary<CaptureCommand, CaptureHotkey> hotkeys)
    {
        WorkbenchSettingsStore.UpdateSettings(settings =>
        {
            settings.CaptureHotkeys.Clear();
            foreach ((CaptureCommand command, CaptureHotkey hotkey) in hotkeys)
            {
                settings.CaptureHotkeys[command.ToString()] = new WorkbenchSettingsStore.CaptureHotkeySetting
                {
                    Modifiers = hotkey.Modifiers,
                    Key = hotkey.Key.ToString()
                };
            }
        });
    }

    public static Dictionary<CaptureCommand, CaptureHotkey> GetDefaultHotkeys()
    {
        return new Dictionary<CaptureCommand, CaptureHotkey>
        {
            [CaptureCommand.FullScreen] = new(ModifierKeys.Control | ModifierKeys.Shift, Key.D1),
            [CaptureCommand.ActiveWindow] = new(ModifierKeys.Control | ModifierKeys.Shift, Key.D2),
            [CaptureCommand.Region] = new(ModifierKeys.Control | ModifierKeys.Shift, Key.D3)
        };
    }
}
