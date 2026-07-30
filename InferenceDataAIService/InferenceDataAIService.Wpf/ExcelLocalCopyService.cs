using System.IO;
using System.Security.Cryptography;

namespace InferenceDataAIService.Wpf;

internal static class ExcelLocalCopyService
{
    internal static string GetLocalCopyBase(
        string configuredArchiveDirectory) =>
        Path.GetFullPath(configuredArchiveDirectory);

    internal static Task<ExcelLocalCopyResult> CopyAsync(
        IReadOnlyList<ExcelFolderSearchRow> sourceRows,
        string archiveDirectory,
        IProgress<ExcelLocalCopyProgress>? progress = null) =>
        Task.Run(() => Copy(
            sourceRows,
            archiveDirectory,
            progress));

    private static ExcelLocalCopyResult Copy(
        IReadOnlyList<ExcelFolderSearchRow> sourceRows,
        string archiveDirectory,
        IProgress<ExcelLocalCopyProgress>? progress)
    {
        var rows = sourceRows
            .Where(row => !row.ExistsInDatabase)
            .ToArray();
        if (rows.Length == 0)
            throw new InvalidOperationException(
                "로컬로 복사할 DB 없음 파일이 없습니다.");

        var copyBase = Path.GetFullPath(
            GetLocalCopyBase(archiveDirectory));
        Directory.CreateDirectory(copyBase);

        var copied = new List<string>(rows.Length);
        for (var index = 0; index < rows.Length; index++)
        {
            var row = rows[index];
            progress?.Report(new ExcelLocalCopyProgress(
                index + 1,
                rows.Length,
                row.FileName));

            var sourcePath = Path.GetFullPath(row.FullPath);
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException(
                    "복사할 원본 Excel을 찾을 수 없습니다.",
                    sourcePath);

            var destinationPath = ResolveDestinationPath(
                copyBase,
                sourcePath);
            EnsureChildPath(copyBase, destinationPath);
            if (!string.Equals(
                    sourcePath,
                    destinationPath,
                    StringComparison.OrdinalIgnoreCase)
                && !FilesMatch(sourcePath, destinationPath))
            {
                File.Copy(
                    sourcePath,
                    destinationPath,
                    overwrite: false);
            }

            var sourceLength = new FileInfo(sourcePath).Length;
            var destinationLength =
                new FileInfo(destinationPath).Length;
            if (sourceLength != destinationLength)
            {
                throw new IOException(
                    "로컬 복사 크기 검증에 실패했습니다. "
                    + $"원본 {sourceLength:N0} bytes, "
                    + $"복사본 {destinationLength:N0} bytes");
            }
            copied.Add(destinationPath);
        }

        return new ExcelLocalCopyResult(
            copyBase,
            copied);
    }

    private static string ResolveDestinationPath(
        string archiveDirectory,
        string sourcePath)
    {
        var fileName = Path.GetFileName(sourcePath);
        var candidate = Path.GetFullPath(
            Path.Combine(archiveDirectory, fileName));
        if (!File.Exists(candidate)
            || FilesMatch(sourcePath, candidate))
        {
            return candidate;
        }

        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        for (var suffix = 2; suffix < int.MaxValue; suffix++)
        {
            candidate = Path.GetFullPath(
                Path.Combine(
                    archiveDirectory,
                    $"{stem} ({suffix}){extension}"));
            if (!File.Exists(candidate)
                || FilesMatch(sourcePath, candidate))
            {
                return candidate;
            }
        }

        throw new IOException(
            $"보관함에서 사용할 파일명을 만들 수 없습니다: {fileName}");
    }

    private static bool FilesMatch(
        string sourcePath,
        string destinationPath)
    {
        if (!File.Exists(destinationPath))
            return false;
        if (string.Equals(
                Path.GetFullPath(sourcePath),
                Path.GetFullPath(destinationPath),
                StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var sourceInfo = new FileInfo(sourcePath);
        var destinationInfo = new FileInfo(destinationPath);
        if (sourceInfo.Length != destinationInfo.Length)
            return false;

        using var source = File.OpenRead(sourcePath);
        using var destination = File.OpenRead(destinationPath);
        var sourceHash = SHA256.HashData(source);
        var destinationHash = SHA256.HashData(destination);
        return sourceHash.AsSpan().SequenceEqual(destinationHash);
    }

    private static void EnsureChildPath(
        string rootPath,
        string candidatePath)
    {
        var normalizedRoot = Path.GetFullPath(rootPath)
            .TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        var normalizedCandidate = Path.GetFullPath(candidatePath);
        if (!normalizedCandidate.StartsWith(
                normalizedRoot,
                StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"로컬 복사 범위를 벗어난 경로입니다: {candidatePath}");
        }
    }
}

internal sealed record ExcelLocalCopyResult(
    string LocalRoot,
    IReadOnlyList<string> CopiedPaths);

internal sealed record ExcelLocalCopyProgress(
    int Current,
    int Total,
    string FileName);
