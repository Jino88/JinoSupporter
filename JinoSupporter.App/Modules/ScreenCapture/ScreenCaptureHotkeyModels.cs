using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;

namespace WorkbenchHost.Modules.ScreenCapture;

public enum CaptureCommand
{
    FullScreen,
    ActiveWindow,
    Region
}

public readonly record struct CaptureHotkey(ModifierKeys Modifiers, Key Key)
{
    public static CaptureHotkey None => new(ModifierKeys.None, Key.None);

    public bool IsEmpty => Key == Key.None;

    public string ToDisplayString()
    {
        if (IsEmpty)
        {
            return "Disabled";
        }

        List<string> parts = new();
        if (Modifiers.HasFlag(ModifierKeys.Control))
        {
            parts.Add("Ctrl");
        }

        if (Modifiers.HasFlag(ModifierKeys.Shift))
        {
            parts.Add("Shift");
        }

        if (Modifiers.HasFlag(ModifierKeys.Alt))
        {
            parts.Add("Alt");
        }

        if (Modifiers.HasFlag(ModifierKeys.Windows))
        {
            parts.Add("Win");
        }

        parts.Add(GetKeyDisplayName(Key));
        return string.Join(" + ", parts.Where(part => !string.IsNullOrWhiteSpace(part)));
    }

    private static string GetKeyDisplayName(Key key)
    {
        if (key is >= Key.D0 and <= Key.D9)
        {
            return ((char)('0' + (key - Key.D0))).ToString();
        }

        if (key is >= Key.NumPad0 and <= Key.NumPad9)
        {
            return $"Num {(char)('0' + (key - Key.NumPad0))}";
        }

        if (key is >= Key.A and <= Key.Z)
        {
            return key.ToString();
        }

        return key switch
        {
            Key.PrintScreen => "PrintScreen",
            Key.Escape => "Esc",
            Key.Return => "Enter",
            Key.Prior => "PageUp",
            Key.Next => "PageDown",
            Key.Back => "Backspace",
            Key.OemPlus => "+",
            Key.OemMinus => "-",
            Key.OemComma => ",",
            Key.OemPeriod => ".",
            _ => key.ToString()
        };
    }
}   

public static class CaptureCommandExtensions
{
    public static string GetDisplayName(this CaptureCommand command)
    {
        return command switch
        {
            CaptureCommand.FullScreen => "Capture Full Screen",
            CaptureCommand.ActiveWindow => "Capture Active Window",
            CaptureCommand.Region => "Capture Region",
            _ => command.ToString()
        };
    }
}
