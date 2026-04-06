using System.Security.Cryptography;
using System.IO;

namespace DiskTree.Services;

public static class FileHasher
{
    public static string ComputeHeadTailHash(string filePath, long fileSize)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var hasher = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

        if (fileSize <= 0)
        {
            return Convert.ToHexString(hasher.GetHashAndReset());
        }

        long sampleLength = Math.Max(1, (long)Math.Ceiling(fileSize * 0.10));

        if (sampleLength * 2 >= fileSize)
        {
            AppendRange(stream, hasher, 0, fileSize);
        }
        else
        {
            AppendRange(stream, hasher, 0, sampleLength);
            AppendRange(stream, hasher, fileSize - sampleLength, sampleLength);
        }

        return Convert.ToHexString(hasher.GetHashAndReset());
    }

    public static string ComputeFullHash(string filePath)
    {
        using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var sha = SHA256.Create();
        byte[] hash = sha.ComputeHash(stream);
        return Convert.ToHexString(hash);
    }

    private static void AppendRange(FileStream stream, IncrementalHash hasher, long start, long count)
    {
        stream.Seek(start, SeekOrigin.Begin);
        byte[] buffer = new byte[81920];
        long remaining = count;

        while (remaining > 0)
        {
            int toRead = remaining > buffer.Length ? buffer.Length : (int)remaining;
            int read = stream.Read(buffer, 0, toRead);
            if (read <= 0)
            {
                break;
            }

            hasher.AppendData(buffer, 0, read);
            remaining -= read;
        }
    }
}
