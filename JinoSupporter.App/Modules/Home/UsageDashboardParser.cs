using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;

namespace JinoSupporter.App.Modules.Home;

internal sealed class UsageDashboardParser
{
    private readonly string _providerName;
    private readonly CardDefinition[] _cardDefinitions;

    public UsageDashboardParser(string providerName, params CardDefinition[] cardDefinitions)
    {
        _providerName = providerName;
        _cardDefinitions = cardDefinitions;
    }

    public CodexUsageSnapshot CreateDefaultSnapshot()
    {
        CodexUsageSnapshot snapshot = new()
        {
            IsAuthenticated = false,
            StatusMessage = $"Refresh를 눌러 {_providerName} usage를 읽어오세요."
        };

        foreach (CardDefinition definition in _cardDefinitions)
        {
            snapshot.Cards.Add(new CodexUsageCard
            {
                Title = definition.Title,
                Value = "-",
                Detail = string.Empty
            });
        }

        return snapshot;
    }

    public CodexUsageSnapshot ParseUsageText(string? pageText)
    {
        CodexUsageSnapshot snapshot = CreateDefaultSnapshot();

        if (string.IsNullOrWhiteSpace(pageText))
        {
            snapshot.StatusMessage = "사용량 페이지에서 읽을 텍스트가 없습니다.";
            return snapshot;
        }

        string decoded = WebUtility.HtmlDecode(pageText);
        string[] lines = decoded
            .Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
            .Select(line => Regex.Replace(line, "\\s+", " ").Trim())
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToArray();

        string normalized = string.Join('\n', lines);
        snapshot.DebugText = string.Join(Environment.NewLine, lines.Take(120));

        if (LooksLikeUnauthenticated(normalized))
        {
            snapshot.StatusMessage = $"{_providerName} 로그인 세션이 아직 준비되지 않았습니다.";
            return snapshot;
        }

        snapshot.IsAuthenticated = true;
        snapshot.StatusMessage = $"로그인된 브라우저 세션에서 {_providerName} 사용량을 읽었습니다.";
        ApplyParsedValues(snapshot, lines, normalized);
        return snapshot;
    }

    private static bool LooksLikeUnauthenticated(string text)
    {
        string lower = text.ToLowerInvariant();
        return (lower.Contains("login") || lower.Contains("log in") || lower.Contains("sign in") || lower.Contains("continue with")) &&
               !lower.Contains("remaining credits") &&
               !lower.Contains("남은 크레딧") &&
               !lower.Contains("usage");
    }

    private void ApplyParsedValues(CodexUsageSnapshot snapshot, IReadOnlyList<string> lines, string normalized)
    {
        for (int i = 0; i < snapshot.Cards.Count; i++)
        {
            CodexUsageCard card = snapshot.Cards[i];
            CardDefinition definition = _cardDefinitions[i];

            if (TryExtractFromLines(lines, definition, out string? value, out string? detail) ||
                TryExtractFromNormalized(normalized, definition, out value, out detail))
            {
                card.Value = value ?? "-";
                card.Detail = detail ?? string.Empty;
            }
        }
    }

    private bool TryExtractFromLines(
        IReadOnlyList<string> lines,
        CardDefinition definition,
        out string? value,
        out string? detail)
    {
        value = null;
        detail = null;

        for (int i = 0; i < lines.Count; i++)
        {
            if (!MatchesAlias(lines[i], definition))
            {
                continue;
            }

            for (int offset = 1; offset <= 5 && i + offset < lines.Count; offset++)
            {
                string candidate = lines[i + offset];
                if (MatchesAnyTitle(candidate))
                {
                    break;
                }

                value ??= ExtractValue(candidate);
                if (detail is null && LooksLikeDetail(candidate))
                {
                    detail = candidate;
                }
            }

            return value is not null || detail is not null;
        }

        return false;
    }

    private static bool TryExtractFromNormalized(
        string normalized,
        CardDefinition definition,
        out string? value,
        out string? detail)
    {
        value = null;
        detail = null;

        foreach (string alias in definition.Aliases)
        {
            string escaped = Regex.Escape(alias);
            Match percentMatch = Regex.Match(
                normalized,
                $"{escaped}.{{0,150}}?(\\d{{1,3}}%\\s*(?:남음|remaining)?)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (percentMatch.Success)
            {
                value = percentMatch.Groups[1].Value.Trim();
            }

            if (value is null)
            {
                Match numberMatch = Regex.Match(
                    normalized,
                    $"{escaped}.{{0,150}}?([0-9]+)",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (numberMatch.Success)
                {
                    value = numberMatch.Groups[1].Value.Trim();
                }
            }

            Match detailMatch = Regex.Match(
                normalized,
                $"{escaped}.{{0,200}}?((?:초기화|reset|days left|hours left)[^.<>\\n]{{0,120}})",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (detailMatch.Success)
            {
                detail = detailMatch.Groups[1].Value.Trim();
            }

            if (value is not null || detail is not null)
            {
                return true;
            }
        }

        return false;
    }

    private bool MatchesAnyTitle(string line)
    {
        return _cardDefinitions.Any(definition => MatchesAlias(line, definition));
    }

    private static bool MatchesAlias(string line, CardDefinition definition)
    {
        return definition.Aliases.Any(alias => line.Contains(alias, StringComparison.OrdinalIgnoreCase));
    }

    private static string? ExtractValue(string line)
    {
        Match moneyMatch = Regex.Match(line, "(US\\$\\s*[0-9]+(?:\\.[0-9]+)?)", RegexOptions.IgnoreCase);
        if (moneyMatch.Success)
        {
            return moneyMatch.Groups[1].Value.Replace(" ", string.Empty).Trim();
        }

        Match percentMatch = Regex.Match(line, "(\\d{1,3}%\\s*(?:남음|remaining)?)", RegexOptions.IgnoreCase);
        if (percentMatch.Success)
        {
            return percentMatch.Groups[1].Value.Trim();
        }

        Match numberMatch = Regex.Match(line, "\\b(\\d+)\\b");
        return numberMatch.Success ? numberMatch.Groups[1].Value.Trim() : null;
    }

    private static bool LooksLikeDetail(string line)
    {
        return line.Contains("초기화", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("reset", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("days left", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("hours left", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("사용됨", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("재설정", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("현재 잔액", StringComparison.OrdinalIgnoreCase) ||
               line.Contains("last updated", StringComparison.OrdinalIgnoreCase);
    }

    internal sealed record CardDefinition(string Title, params string[] Aliases);
}
