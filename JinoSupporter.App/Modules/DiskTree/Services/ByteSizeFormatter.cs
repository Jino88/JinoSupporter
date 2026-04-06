namespace DiskTree.Services;

public static class ByteSizeFormatter
{
    private static readonly string[] UnitNames = ["B", "KB", "MB", "GB", "TB", "PB"];

    public static string Format(long sizeBytes)
    {
        if (sizeBytes < 0)
        {
            return "0 B";
        }

        double value = sizeBytes;
        int unitIndex = 0;

        while (value >= 1024.0 && unitIndex < UnitNames.Length - 1)
        {
            value /= 1024.0;
            unitIndex++;
        }

        return $"{value:0.##} {UnitNames[unitIndex]}";
    }
}
