using System.IO;
using System.Text;
using System.Text.Json;

namespace InferenceDataAIService.Wpf;

internal static class ExcelFolderHistoryStore
{
    private static readonly UTF8Encoding Utf8WithoutBom =
        new(encoderShouldEmitUTF8Identifier: false);

    internal static string HistoryFilePath => Path.Combine(
        Environment.GetFolderPath(
            Environment.SpecialFolder.LocalApplicationData),
        "InferenceDataAIService",
        "excel-search-folders.json");

    internal static IReadOnlyList<string> Load()
    {
        if (!File.Exists(HistoryFilePath))
            return [];

        try
        {
            var history = JsonSerializer.Deserialize<ExcelFolderHistory>(
                File.ReadAllText(HistoryFilePath, Encoding.UTF8));
            return Normalize(history?.Folders ?? []);
        }
        catch (
            Exception exception
        ) when (
            exception is IOException
            or UnauthorizedAccessException
            or JsonException
            or ArgumentException
            or NotSupportedException)
        {
            return [];
        }
    }

    internal static void Save(IEnumerable<string> folders)
    {
        var normalized = Normalize(folders);
        var directory = Path.GetDirectoryName(HistoryFilePath)
            ?? throw new InvalidOperationException(
                "검색 폴더 이력 저장 위치를 확인할 수 없습니다.");
        Directory.CreateDirectory(directory);
        var temporary = HistoryFilePath
            + "."
            + Guid.NewGuid().ToString("N")
            + ".tmp";
        try
        {
            File.WriteAllText(
                temporary,
                JsonSerializer.Serialize(
                    new ExcelFolderHistory
                    {
                        Folders = [.. normalized],
                    },
                    new JsonSerializerOptions
                    {
                        WriteIndented = true,
                    }) + Environment.NewLine,
                Utf8WithoutBom);
            File.Move(
                temporary,
                HistoryFilePath,
                overwrite: true);
        }
        finally
        {
            if (File.Exists(temporary))
                File.Delete(temporary);
        }
    }

    private static IReadOnlyList<string> Normalize(
        IEnumerable<string> folders)
    {
        var result = new List<string>();
        var seen = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var value in folders)
        {
            if (string.IsNullOrWhiteSpace(value))
                continue;
            try
            {
                var path = Path.GetFullPath(value.Trim());
                if (seen.Add(path))
                    result.Add(path);
            }
            catch (
                Exception exception
            ) when (
                exception is ArgumentException
                or NotSupportedException
                or PathTooLongException)
            {
                // Ignore malformed entries from an older history file.
            }
        }
        return result;
    }

    private sealed class ExcelFolderHistory
    {
        public List<string> Folders { get; set; } = [];
    }
}
