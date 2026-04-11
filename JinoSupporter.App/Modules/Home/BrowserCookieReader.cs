using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;

namespace JinoSupporter.App.Modules.Home;

internal static class BrowserCookieReader
{
    private sealed record BrowserSource(string Name, string UserDataPath);

    private static readonly BrowserSource[] Sources =
    {
        new("Chrome", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Google", "Chrome", "User Data")),
        new("Edge", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Edge", "User Data"))
    };

    public static string? TryGetChatGptCookieHeader(out string statusMessage)
    {
        List<string> errors = new();

        foreach (BrowserSource source in Sources)
        {
            try
            {
                string? header = TryReadFromBrowser(source, out string browserStatus);
                if (!string.IsNullOrWhiteSpace(header))
                {
                    statusMessage = $"{source.Name} 로그인 쿠키를 자동으로 읽었습니다.";
                    return header;
                }

                errors.Add($"{source.Name}: {browserStatus}");
            }
            catch (Exception ex)
            {
                errors.Add($"{source.Name}: {ex.Message}");
            }
        }

        statusMessage = errors.Count > 0
            ? string.Join(Environment.NewLine, errors)
            : "Chrome/Edge에서 chatgpt.com 쿠키를 찾지 못했습니다.";
        return null;
    }

    public static string? TryExtractCookieHeaderFromRawRequest(string? rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText))
        {
            return null;
        }

        Match match = Regex.Match(
            rawText,
            @"(?im)^\s*cookie\s*[:=]\s*(.+?)\s*$");

        return match.Success ? match.Groups[1].Value.Trim() : null;
    }

    private static string? TryReadFromBrowser(BrowserSource source, out string statusMessage)
    {
        if (!Directory.Exists(source.UserDataPath))
        {
            statusMessage = "브라우저 사용자 데이터 폴더가 없습니다.";
            return null;
        }

        string localStatePath = Path.Combine(source.UserDataPath, "Local State");
        if (!File.Exists(localStatePath))
        {
            statusMessage = "Local State 파일이 없습니다.";
            return null;
        }

        byte[] masterKey = GetMasterKey(localStatePath);
        if (masterKey.Length == 0)
        {
            statusMessage = "브라우저 master key를 복호화하지 못했습니다.";
            return null;
        }

        foreach (string profileDir in EnumerateProfiles(source.UserDataPath))
        {
            string cookiesPath = Path.Combine(profileDir, "Network", "Cookies");
            if (!File.Exists(cookiesPath))
            {
                continue;
            }

            string tempPath = Path.Combine(Path.GetTempPath(), $"js_cookie_{Guid.NewGuid():N}.db");
            try
            {
                File.Copy(cookiesPath, tempPath, true);
                string? cookieHeader = ReadChatGptCookiesFromDb(tempPath, masterKey);
                if (!string.IsNullOrWhiteSpace(cookieHeader))
                {
                    statusMessage = $"{Path.GetFileName(profileDir)} 프로필 사용";
                    return cookieHeader;
                }
            }
            finally
            {
                TryDelete(tempPath);
            }
        }

        statusMessage = "chatgpt.com 쿠키를 찾지 못했습니다.";
        return null;
    }

    private static IEnumerable<string> EnumerateProfiles(string userDataPath)
    {
        string[] candidates = { "Default", "Profile 1", "Profile 2", "Profile 3", "Profile 4", "Profile 5" };
        return candidates
            .Select(name => Path.Combine(userDataPath, name))
            .Where(Directory.Exists);
    }

    private static byte[] GetMasterKey(string localStatePath)
    {
        string json = File.ReadAllText(localStatePath);
        using JsonDocument document = JsonDocument.Parse(json);

        if (!document.RootElement.TryGetProperty("os_crypt", out JsonElement osCrypt))
        {
            return Array.Empty<byte>();
        }

        if (!osCrypt.TryGetProperty("encrypted_key", out JsonElement encryptedKeyElement))
        {
            return Array.Empty<byte>();
        }

        string? encryptedKeyBase64 = encryptedKeyElement.GetString();
        if (string.IsNullOrWhiteSpace(encryptedKeyBase64))
        {
            return Array.Empty<byte>();
        }

        byte[] encryptedKey = Convert.FromBase64String(encryptedKeyBase64);
        byte[] keyPayload = encryptedKey.Skip(5).ToArray();
        return ProtectedData.Unprotect(keyPayload, null, DataProtectionScope.CurrentUser);
    }

    private static string? ReadChatGptCookiesFromDb(string dbPath, byte[] masterKey)
    {
        List<string> parts = new();

        using SqliteConnection connection = new($"Data Source={dbPath};Mode=ReadOnly");
        connection.Open();

        using SqliteCommand command = connection.CreateCommand();
        command.CommandText = """
            SELECT name, value, encrypted_value
            FROM cookies
            WHERE host_key LIKE '%chatgpt.com%'
            ORDER BY name
            """;

        using SqliteDataReader reader = command.ExecuteReader();
        while (reader.Read())
        {
            string name = reader.GetString(0);
            string plainValue = reader.IsDBNull(1) ? string.Empty : reader.GetString(1);
            byte[] encryptedValue = reader.IsDBNull(2) ? Array.Empty<byte>() : (byte[])reader["encrypted_value"];

            string value = !string.IsNullOrEmpty(plainValue)
                ? plainValue
                : DecryptCookieValue(encryptedValue, masterKey);

            if (!string.IsNullOrWhiteSpace(value))
            {
                parts.Add($"{name}={value}");
            }
        }

        return parts.Count > 0 ? string.Join("; ", parts) : null;
    }

    private static string DecryptCookieValue(byte[] encryptedValue, byte[] masterKey)
    {
        if (encryptedValue.Length == 0)
        {
            return string.Empty;
        }

        string prefix = Encoding.ASCII.GetString(encryptedValue, 0, Math.Min(3, encryptedValue.Length));
        if (prefix is "v10" or "v11")
        {
            return DecryptAesCookie(encryptedValue, masterKey);
        }

        if (prefix == "v20")
        {
            return string.Empty;
        }

        try
        {
            byte[] decrypted = ProtectedData.Unprotect(encryptedValue, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(decrypted);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string DecryptAesCookie(byte[] encryptedValue, byte[] masterKey)
    {
        if (encryptedValue.Length < 3 + 12 + 16)
        {
            return string.Empty;
        }

        byte[] nonce = encryptedValue.Skip(3).Take(12).ToArray();
        byte[] cipherAndTag = encryptedValue.Skip(15).ToArray();
        byte[] cipher = cipherAndTag.Take(cipherAndTag.Length - 16).ToArray();
        byte[] tag = cipherAndTag.Skip(cipher.Length).ToArray();
        byte[] plaintext = new byte[cipher.Length];

        try
        {
            using AesGcm aes = new(masterKey, 16);
            aes.Decrypt(nonce, cipher, tag, plaintext);
            return Encoding.UTF8.GetString(plaintext);
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }
}
