using System.IO;
using System.Text;
using System.Text.Json;
using Microsoft.Win32;

namespace GraphMaker;

public static class GraphReportFileDialogHelper
{
    private const string DefaultFilter = "Graph report (*.graphreport.json)|*.graphreport.json|JSON files (*.json)|*.json|All files (*.*)|*.*";

    public static JsonSerializerOptions JsonOptions { get; } = new() { WriteIndented = true };

    public static bool SaveState<T>(string title, string fileName, T state)
    {
        var dialog = new SaveFileDialog
        {
            Title = title,
            Filter = DefaultFilter,
            FileName = fileName
        };

        if (dialog.ShowDialog() != true)
        {
            return false;
        }

        File.WriteAllText(dialog.FileName, JsonSerializer.Serialize(state, JsonOptions), Encoding.UTF8);
        return true;
    }

    public static T? LoadState<T>(string title) where T : class
    {
        var dialog = new OpenFileDialog
        {
            Title = title,
            Filter = DefaultFilter,
            Multiselect = false
        };

        if (dialog.ShowDialog() != true)
        {
            return null;
        }

        return JsonSerializer.Deserialize<T>(File.ReadAllText(dialog.FileName), JsonOptions);
    }
}
