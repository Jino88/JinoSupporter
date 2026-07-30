using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace InferenceDataAIService.Wpf;

internal static class IngestVerificationHtmlRenderer
{
    internal static string Render(
        IngestWorkbookResult result,
        RelatedStudiesDocument related)
    {
        if (string.IsNullOrWhiteSpace(result.ManifestPath)
            || !File.Exists(result.ManifestPath))
            return RenderMissingManifest(result);

        using var document = JsonDocument.Parse(
            File.ReadAllText(result.ManifestPath, Encoding.UTF8));
        var root = document.RootElement;
        var sourcePath = ReadPath(root, result.SourcePath);
        var studies = TryArray(root, "studies").ToList();

        var body = new StringBuilder();
        body.Append("<header><div><span class='eyebrow'>처리 결과</span><h1>")
            .Append(H(Path.GetFileName(sourcePath)))
            .Append("</h1><p>시험 조건별 원본 측정값</p></div><span class='status ")
            .Append(StatusClass(result.Status))
            .Append("'>")
            .Append(H(DisplayCode(result.Status)))
            .Append("</span></header>");

        body.Append("<main class='comparison-list'>");
        if (studies.Count == 0)
            body.Append("<p class='warning'>표로 표시할 시험 결과가 없습니다.</p>");
        for (var index = 0; index < studies.Count; index++)
            body.Append(RenderStudy(
                sourcePath,
                studies[index],
                index + 1));
        body.Append("</main>");

        body.Append("<footer>값을 누르면 원본 Excel 셀을 읽기 전용으로 엽니다.</footer>");
        return Page("처리 내용 검증", body.ToString());
    }

    private static string RenderStudy(
        string sourcePath,
        JsonElement study,
        int number)
    {
        var arms = TryArray(study, "arms").ToList();
        var outcomes = TryArray(study, "outcomes").ToList();
        var observedArmKeys = outcomes
            .SelectMany(outcome => TryArray(outcome, "observations"))
            .Select(observation => Text(observation, "arm"))
            .Where(key => !string.IsNullOrWhiteSpace(key))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var visibleArms = arms
            .Where(arm => observedArmKeys.Contains(Text(arm, "key")))
            .ToList();
        var columns = BuildOutcomeColumns(outcomes);
        var roles = BuildComparisonRoles(study, arms);
        var review = BuildStudyReviewContext(study, arms, outcomes);
        var reviewTitle = review.AllIssues.Count == 0
            ? string.Empty
            : "검토 사유:\n- " + string.Join(
                "\n- ",
                review.AllIssues);
        var builder = new StringBuilder("<section class='comparison-card'>")
            .Append("<div class='comparison-heading'><div><span>검토 ")
            .Append(number)
            .Append("</span><h2>")
            .Append(H(TranslateNarrative(Text(study, "title"))))
            .Append("</h2></div><span class='status ")
            .Append(StatusClass(Text(study, "verificationStatus")))
            .Append("'")
            .Append(string.IsNullOrWhiteSpace(reviewTitle)
                ? string.Empty
                : " title='" + HA(reviewTitle) + "'")
            .Append(">")
            .Append(H(DisplayCode(Text(study, "verificationStatus"))))
            .Append("</span></div>");
        if (visibleArms.Count == 0 || columns.Count == 0)
            return builder
                .Append("<p class='warning'>표로 표시할 조건별 측정값이 없습니다.</p></section>")
                .ToString();

        builder.Append("<div class='comparison-table-wrap'><table class='comparison-table'><thead><tr>")
            .Append("<th class='cohort-column'>구분</th>")
            .Append("<th class='condition-column'>조건</th>");
        foreach (var column in columns)
            builder.Append("<th>")
                .Append(H(column.Label))
                .Append("</th>");
        builder.Append("</tr></thead><tbody>");
        foreach (var arm in visibleArms)
        {
            var armKey = Text(arm, "key");
            var role = roles.GetValueOrDefault(armKey, "조건");
            var roleClass = role == "기준군" ? "control" : "comparison";
            var armIssues = review.ArmIssues.GetValueOrDefault(
                armKey,
                []);
            var armReviewTitle = armIssues.Count == 0
                ? string.Empty
                : "검토 사유:\n- " + string.Join("\n- ", armIssues);
            builder.Append("<tr><td class='cohort-column'><span class='cohort-badge ")
                .Append(roleClass)
                .Append(armIssues.Count == 0 ? string.Empty : " review-issue")
                .Append("'")
                .Append(string.IsNullOrWhiteSpace(armReviewTitle)
                    ? string.Empty
                    : " title='" + HA(armReviewTitle) + "'")
                .Append(">")
                .Append(H(role))
                .Append("</span></td><td class='condition-column'>")
                .Append(RenderArmLink(
                    sourcePath,
                    arm,
                    armIssues))
                .Append("</td>");
            foreach (var column in columns)
            {
                var cell = RenderOutcomeCell(
                        sourcePath,
                        armKey,
                        column.Outcomes,
                        review);
                builder.Append("<td class='metric-cell")
                    .Append(cell.HasReviewIssue ? " review-issue" : string.Empty)
                    .Append("'>")
                    .Append(cell.Html)
                    .Append("</td>");
            }
            builder.Append("</tr>");
        }
        builder.Append("</tbody></table></div></section>");
        return builder.ToString();
    }

    private static List<OutcomeColumn> BuildOutcomeColumns(
        IReadOnlyList<JsonElement> outcomes)
    {
        var columns = new List<OutcomeColumn>();
        foreach (var outcome in outcomes)
        {
            var originalLabel = Text(outcome, "originalLabel").Trim();
            if (string.IsNullOrWhiteSpace(originalLabel)
                || !TryArray(outcome, "observations").Any())
                continue;
            var key = Regex.Replace(
                originalLabel,
                @"\s+rate$",
                string.Empty,
                RegexOptions.IgnoreCase).Trim();
            var column = columns.FirstOrDefault(item =>
                string.Equals(
                    item.Key,
                    key,
                    StringComparison.OrdinalIgnoreCase));
            if (column is null)
            {
                column = new OutcomeColumn(
                    key,
                    DisplayOutcomeLabel(key),
                    []);
                columns.Add(column);
            }
            column.Outcomes.Add(outcome);
        }
        return columns;
    }

    private static Dictionary<string, string> BuildComparisonRoles(
        JsonElement study,
        IReadOnlyList<JsonElement> arms)
    {
        var roles = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var comparison in TryArray(study, "comparisons"))
        {
            var controlArm = Text(comparison, "controlArm");
            var comparedArm = Text(comparison, "comparedArm");
            if (!string.IsNullOrWhiteSpace(controlArm))
                roles.TryAdd(controlArm, "기준군");
            if (!string.IsNullOrWhiteSpace(comparedArm))
                roles.TryAdd(comparedArm, "비교군");
        }
        foreach (var arm in arms)
        {
            var key = Text(arm, "key");
            var role = Text(arm, "role").Trim().ToUpperInvariant();
            if (role is "CONTROL" or "REFERENCE")
                roles[key] = "기준군";
            else if (role is "TREATMENT" or "COMPARATOR")
                roles[key] = "비교군";
        }
        return roles;
    }

    private static string RenderArmLink(
        string sourcePath,
        JsonElement arm,
        IReadOnlyList<string> reviewIssues)
    {
        var label = TranslateArm(Coalesce(
            Text(arm, "label"),
            Text(arm, "condition")));
        var first = TryArray(arm, "evidence").FirstOrDefault();
        if (first.ValueKind == JsonValueKind.Undefined)
            return "<strong>" + H(label) + "</strong>";
        return ReplaceEvidenceLinkText(
            EvidenceLink(sourcePath, first),
            label,
            BuildEvidenceTooltip(
                first,
                label,
                reviewIssues));
    }

    private static OutcomeCell RenderOutcomeCell(
        string sourcePath,
        string armKey,
        IReadOnlyList<JsonElement> outcomes,
        StudyReviewContext review)
    {
        var values = new List<string>();
        var cellIssues = new List<string>();
        foreach (var outcome in outcomes)
        {
            var observation = TryArray(outcome, "observations")
                .FirstOrDefault(item => string.Equals(
                    Text(item, "arm"),
                    armKey,
                    StringComparison.OrdinalIgnoreCase));
            if (observation.ValueKind == JsonValueKind.Undefined)
                continue;
            var value = ObservationValue(observation);
            var issues = review.CellIssues.GetValueOrDefault(
                ReviewCellKey(armKey, outcome),
                []);
            foreach (var issue in issues)
            {
                if (!cellIssues.Contains(
                        issue,
                        StringComparer.Ordinal))
                    cellIssues.Add(issue);
            }
            var firstEvidence = TryArray(observation, "evidence")
                .FirstOrDefault();
            var renderedValue = firstEvidence.ValueKind == JsonValueKind.Undefined
                ? "<strong>" + H(value) + "</strong>"
                : ReplaceEvidenceLinkText(
                    EvidenceLink(sourcePath, firstEvidence),
                    value,
                    BuildEvidenceTooltip(
                        firstEvidence,
                        value,
                        issues));
            var metricType = Text(outcome, "metricType");
            values.Add(metricType.Contains(
                    "rate",
                    StringComparison.OrdinalIgnoreCase)
                ? "<span class='rate-value'>" + renderedValue + "</span>"
                : renderedValue);
        }
        if (values.Count == 0)
            return new OutcomeCell(
                "<span class='empty-value'>—</span>",
                false);
        var marker = cellIssues.Count == 0
            ? string.Empty
            : "<span class='review-marker' title='"
              + HA("검토 사유:\n- " + string.Join(
                  "\n- ",
                  cellIssues))
              + "'>!</span>";
        return new OutcomeCell(
            string.Join(
                "<span class='value-separator'>/</span>",
                values)
            + marker,
            cellIssues.Count > 0);
    }

    private static string ReplaceEvidenceLinkText(
        string link,
        string text,
        string title)
    {
        var start = link.IndexOf('>');
        var end = link.LastIndexOf("</a>", StringComparison.Ordinal);
        if (start < 0 || end <= start)
            return "<strong>" + H(text) + "</strong>";
        return link[..start]
            + " title='" + HA(title) + "'><strong>"
            + H(text)
            + "</strong></a>";
    }

    private static StudyReviewContext BuildStudyReviewContext(
        JsonElement study,
        IReadOnlyList<JsonElement> arms,
        IReadOnlyList<JsonElement> outcomes)
    {
        var limitations = TryArray(study, "limitations")
            .Select(item => item.ValueKind == JsonValueKind.String
                ? item.GetString() ?? string.Empty
                : item.ToString())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Where(item => !IsComparisonCaution(item))
            .Select(item => new ReviewLimitation(
                item,
                TranslateNarrative(item)))
            .ToList();
        var armIssues = new Dictionary<string, List<string>>(
            StringComparer.OrdinalIgnoreCase);
        var cellIssues = new Dictionary<string, List<string>>(
            StringComparer.OrdinalIgnoreCase);

        foreach (var limitation in limitations)
        {
            var mentionedArms = arms
                .Where(arm => LimitationMentionsArm(
                    limitation.Original,
                    arm))
                .Select(arm => Text(arm, "key"))
                .Where(key => !string.IsNullOrWhiteSpace(key))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var isRoleIssue = limitation.Original.Contains(
                    "does not label",
                    StringComparison.OrdinalIgnoreCase)
                && limitation.Original.Contains(
                    "Control",
                    StringComparison.OrdinalIgnoreCase);
            if (isRoleIssue)
            {
                foreach (var arm in arms)
                    AddReviewIssue(
                        armIssues,
                        Text(arm, "key"),
                        limitation.Display);
            }

            foreach (var outcome in outcomes)
            {
                if (!LimitationMentionsOutcome(
                        limitation.Original,
                        Text(outcome, "originalLabel")))
                    continue;
                foreach (var observation in TryArray(
                             outcome,
                             "observations"))
                {
                    var armKey = Text(observation, "arm");
                    if (mentionedArms.Count > 0
                        && !mentionedArms.Contains(armKey))
                        continue;
                    AddReviewIssue(
                        cellIssues,
                        ReviewCellKey(armKey, outcome),
                        limitation.Display);
                }
            }
        }

        return new StudyReviewContext(
            limitations.Select(item => item.Display).ToList(),
            armIssues,
            cellIssues);
    }

    private static bool IsComparisonCaution(string limitation)
    {
        if (Regex.IsMatch(
                limitation,
                @"(?:compared .*conditions|assembly arms).*different Input counts",
                RegexOptions.IgnoreCase))
            return true;
        if (limitation.StartsWith(
                "No randomization",
                StringComparison.OrdinalIgnoreCase))
            return true;
        return limitation.Contains(
                   "does not label",
                   StringComparison.OrdinalIgnoreCase)
               && limitation.Contains(
                   "Control",
                   StringComparison.OrdinalIgnoreCase);
    }

    private static bool LimitationMentionsArm(
        string limitation,
        JsonElement arm)
    {
        var labels = new[]
        {
            Text(arm, "label"),
            Text(arm, "condition"),
        };
        return labels.Any(label =>
            !string.IsNullOrWhiteSpace(label)
            && limitation.Contains(
                label,
                StringComparison.OrdinalIgnoreCase));
    }

    private static bool LimitationMentionsOutcome(
        string limitation,
        string outcomeLabel)
    {
        if (string.IsNullOrWhiteSpace(outcomeLabel))
            return false;
        var normalizedLabel = Regex.Replace(
            outcomeLabel,
            @"\s+rate$",
            string.Empty,
            RegexOptions.IgnoreCase).Trim();
        if (limitation.Contains(
                normalizedLabel,
                StringComparison.OrdinalIgnoreCase))
            return true;
        if (normalizedLabel.StartsWith(
                "NG visual",
                StringComparison.OrdinalIgnoreCase))
            return limitation.Contains(
                "NG visual",
                StringComparison.OrdinalIgnoreCase);
        if (normalizedLabel.StartsWith(
                "NG BAKO",
                StringComparison.OrdinalIgnoreCase))
            return limitation.Contains(
                "NG BAKO",
                StringComparison.OrdinalIgnoreCase);
        if (normalizedLabel.StartsWith(
                "HEARING",
                StringComparison.OrdinalIgnoreCase))
            return limitation.Contains(
                "HEARING",
                StringComparison.OrdinalIgnoreCase);
        return normalizedLabel switch
        {
            "Input" => Regex.IsMatch(
                limitation,
                @"\bInput\b",
                RegexOptions.IgnoreCase),
            "OK" => Regex.IsMatch(
                limitation,
                @"\bOK\b",
                RegexOptions.IgnoreCase),
            "Total NG" => limitation.Contains(
                "Total NG",
                StringComparison.OrdinalIgnoreCase),
            _ => false,
        };
    }

    private static string ReviewCellKey(
        string armKey,
        JsonElement outcome) =>
        armKey + "\u001f" + Coalesce(
            Text(outcome, "key"),
            Text(outcome, "originalLabel"));

    private static void AddReviewIssue(
        IDictionary<string, List<string>> target,
        string key,
        string issue)
    {
        if (string.IsNullOrWhiteSpace(key)
            || string.IsNullOrWhiteSpace(issue))
            return;
        if (!target.TryGetValue(key, out var issues))
        {
            issues = [];
            target[key] = issues;
        }
        if (!issues.Contains(issue, StringComparer.Ordinal))
            issues.Add(issue);
    }

    private static string BuildEvidenceTooltip(
        JsonElement evidence,
        string displayedValue,
        IReadOnlyList<string> reviewIssues)
    {
        var sheet = Text(evidence, "sheet");
        var range = Text(evidence, "range");
        var sourceText = Coalesce(
            Text(evidence, "sourceText"),
            displayedValue);
        var tooltip = new StringBuilder()
            .Append("원본 Excel: ")
            .Append(sheet)
            .Append("!")
            .Append(range)
            .Append("\n원본 표시값: ")
            .Append(sourceText)
            .Append("\n클릭하여 원본 셀 열기");
        if (reviewIssues.Count > 0)
            tooltip.Append("\n\n검토 사유:\n- ")
                .Append(string.Join("\n- ", reviewIssues));
        return tooltip.ToString();
    }

    private static string DisplayOutcomeLabel(string value) =>
        value switch
        {
            "Input" => "검사 수량",
            "Total NG" => "전체 NG",
            "HEARING - Input" => "Hearing 검사 수량",
            _ => value
                .Replace("NG visual - ", "외관 NG · ", StringComparison.OrdinalIgnoreCase)
                .Replace("NG BAKO - ", "NG BAKO · ", StringComparison.OrdinalIgnoreCase)
                .Replace("HEARING - ", "Hearing · ", StringComparison.OrdinalIgnoreCase),
        };

    private static string RenderArms(
        string sourcePath,
        IReadOnlyList<JsonElement> arms)
    {
        if (arms.Count == 0) return string.Empty;
        var rows = new StringBuilder();
        foreach (var arm in arms)
        {
            rows.Append("<tr><td>")
                .Append(H(DisplayCode(Text(arm, "role"))))
                .Append("</td><td><strong>")
                .Append(H(Coalesce(
                    Text(arm, "label"),
                    Text(arm, "condition"))))
                .Append("</strong></td><td>")
                .Append(H(Text(arm, "condition")))
                .Append("</td><td>")
                .Append(H(Scalar(arm, "sampleSize")))
                .Append("</td><td>")
                .Append(H(RenderFactorValues(arm)))
                .Append("</td><td>")
                .Append(FirstEvidenceLink(
                    sourcePath,
                    TryArray(arm, "evidence")))
                .Append("</td></tr>");
        }
        return "<h4>시험군·비교군·조건</h4><div class='table-wrap'><table>"
            + "<thead><tr><th>역할</th><th>표시명</th><th>조건</th>"
            + "<th>표본수</th><th>변경 요인</th><th>원본</th></tr></thead><tbody>"
            + rows
            + "</tbody></table></div>";
    }

    private static string RenderObservations(
        string sourcePath,
        IEnumerable<JsonElement> outcomes,
        IReadOnlyDictionary<string, string> armLabels)
    {
        var rows = new StringBuilder();
        var count = 0;
        foreach (var outcome in outcomes)
        {
            foreach (var observation in TryArray(outcome, "observations"))
            {
                count++;
                var armKey = Text(observation, "arm");
                armLabels.TryGetValue(armKey, out var armLabel);
                rows.Append("<tr><td>")
                    .Append(H(Text(outcome, "originalLabel")))
                    .Append("</td><td>")
                    .Append(H(Text(outcome, "metricType")))
                    .Append("</td><td>")
                    .Append(H(Coalesce(armLabel, armKey)))
                    .Append("</td><td class='value'>")
                    .Append(H(ObservationValue(observation)))
                    .Append("</td><td>")
                    .Append(H(Text(outcome, "unit")))
                    .Append("</td><td>")
                    .Append(FirstEvidenceLink(
                        sourcePath,
                        TryArray(observation, "evidence")))
                    .Append("</td></tr>");
            }
        }
        if (count == 0) return "<p class='warning'>저장된 관측값이 없습니다.</p>";
        return "<h4>측정값·관측값</h4><div class='table-wrap observation-wrap'><table>"
            + "<thead><tr><th>측정 항목</th><th>형식</th><th>시험군·조건</th>"
            + "<th>저장 값</th><th>단위</th><th>원본 셀</th></tr></thead><tbody>"
            + rows
            + "</tbody></table></div>";
    }

    private static string RenderComparisons(
        string sourcePath,
        IEnumerable<JsonElement> comparisons,
        IReadOnlyDictionary<string, string> armLabels)
    {
        var rows = new StringBuilder();
        var count = 0;
        foreach (var comparison in comparisons)
        {
            count++;
            var comparedKey = Text(comparison, "comparedArm");
            var controlKey = Text(comparison, "controlArm");
            armLabels.TryGetValue(comparedKey, out var compared);
            armLabels.TryGetValue(controlKey, out var control);
            rows.Append("<tr><td>")
                .Append(H(Coalesce(compared, comparedKey)))
                .Append("</td><td>")
                .Append(H(Coalesce(control, controlKey)))
                .Append("</td><td>")
                .Append(H(DisplayCode(Text(comparison, "validityStatus"))))
                .Append("</td><td>")
                .Append(H(DisplayCode(Text(comparison, "confoundingStatus"))))
                .Append("</td><td>")
                .Append(H(Scalar(comparison, "aggregationEligible")))
                .Append("</td><td>")
                .Append(FirstEvidenceLink(
                    sourcePath,
                    TryArray(comparison, "evidence")))
                .Append("</td></tr>");
        }
        if (count == 0) return string.Empty;
        return "<h4>비교 정의</h4><div class='table-wrap'><table>"
            + "<thead><tr><th>비교군</th><th>기준군</th><th>유효성</th>"
            + "<th>교란</th><th>집계 가능</th><th>원본</th></tr></thead><tbody>"
            + rows
            + "</tbody></table></div>";
    }

    private static string RenderEvidence(
        string sourcePath,
        IEnumerable<JsonElement> evidence,
        string title)
    {
        var items = evidence.ToList();
        if (items.Count == 0) return string.Empty;
        var links = string.Join(
            string.Empty,
            items.Select(item =>
                "<li>" + EvidenceLink(sourcePath, item)
                + (!string.IsNullOrWhiteSpace(Text(item, "sourceText"))
                    ? "<span>" + H(Text(item, "sourceText")) + "</span>"
                    : string.Empty)
                + "</li>"));
        return "<details><summary>" + H(title)
            + " · " + items.Count.ToString("N0") + "개</summary><ul class='evidence-list'>"
            + links + "</ul></details>";
    }

    private static string RenderLimitations(JsonElement element)
    {
        var limitations = TryArray(element, "limitations")
            .Select(item => item.ValueKind == JsonValueKind.String
                ? item.GetString() ?? string.Empty
                : item.ToString())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .ToList();
        if (limitations.Count == 0) return string.Empty;
        return "<details open><summary>검토 필요사항 · "
            + limitations.Count.ToString("N0")
            + "개</summary><ul class='limitations'>"
            + string.Join(string.Empty, limitations.Select(item =>
                "<li>" + H(TranslateNarrative(item)) + "</li>"))
            + "</ul></details>";
    }

    private static string RenderRelated(RelatedStudiesDocument related)
    {
        if (related.Studies.Count == 0)
            return "<p>관련 Study가 없습니다.</p>";
        return "<div class='table-wrap'><table><thead><tr><th>DATA ID</th>"
            + "<th>유사도</th><th>Study</th><th>원본</th></tr></thead><tbody>"
            + string.Join(string.Empty, related.Studies.Select(item =>
                "<tr><td><code>" + H(item.PublicDataId) + "</code></td><td>"
                + item.SimilarityScore.ToString("F3", CultureInfo.InvariantCulture)
                + "</td><td>" + H(item.Title) + "</td><td>"
                + H(Path.GetFileName(item.SourcePath)) + "</td></tr>"))
            + "</tbody></table></div>";
    }

    private static int CountStudyEvidence(JsonElement study)
    {
        var count = TryArray(study, "evidence").Count();
        count += TryArray(study, "arms").Sum(arm =>
            TryArray(arm, "evidence").Count());
        count += TryArray(study, "comparisons").Sum(comparison =>
            TryArray(comparison, "evidence").Count());
        count += TryArray(study, "outcomes").Sum(outcome =>
            TryArray(outcome, "evidence").Count()
            + TryArray(outcome, "observations").Sum(observation =>
                TryArray(observation, "evidence").Count()));
        return count;
    }

    private static string ObservationValue(JsonElement observation)
    {
        var valueText = Text(observation, "valueText");
        if (!string.IsNullOrWhiteSpace(valueText)) return valueText;
        var valueNumber = Scalar(observation, "valueNumber");
        if (!string.IsNullOrWhiteSpace(valueNumber)) return valueNumber;
        var average = Scalar(observation, "average");
        var min = Scalar(observation, "min");
        var max = Scalar(observation, "max");
        if (!string.IsNullOrWhiteSpace(average)
            || !string.IsNullOrWhiteSpace(min)
            || !string.IsNullOrWhiteSpace(max))
            return $"AVG {average} · MIN {min} · MAX {max}".Trim();
        var numerator = Scalar(observation, "numerator");
        var denominator = Scalar(observation, "denominator");
        if (!string.IsNullOrWhiteSpace(numerator)
            || !string.IsNullOrWhiteSpace(denominator))
            return $"{numerator}/{denominator}";
        return "(값 없음)";
    }

    private static string RenderFactorValues(JsonElement arm) =>
        string.Join(
            ", ",
            TryArray(arm, "factorValues").Select(item =>
            {
                var value = Coalesce(
                    Text(item, "value"),
                    Scalar(item, "valueNumber"));
                var unit = Text(item, "unit");
                return $"{Text(item, "factor")}: {value}{(
                    string.IsNullOrWhiteSpace(unit) ? string.Empty : " " + unit)}";
            }));

    private static string FirstEvidenceLink(
        string sourcePath,
        IEnumerable<JsonElement> evidence)
    {
        var first = evidence.FirstOrDefault();
        return first.ValueKind == JsonValueKind.Undefined
            ? string.Empty
            : EvidenceLink(sourcePath, first);
    }

    private static string EvidenceLink(string sourcePath, JsonElement evidence)
    {
        var sheet = Text(evidence, "sheet");
        var range = Text(evidence, "range");
        if (string.IsNullOrWhiteSpace(sheet)
            || string.IsNullOrWhiteSpace(range))
            return H(Coalesce(Text(evidence, "sourceText"), "(좌표 없음)"));
        var href = "inference-excel://open/?source="
            + WebUtility.UrlEncode(sourcePath)
            + "&sheet=" + WebUtility.UrlEncode(sheet)
            + "&range=" + WebUtility.UrlEncode(range);
        return "<a href='" + H(href) + "'>"
            + H($"{sheet}!{range}") + "</a>";
    }

    private static string ReadPath(JsonElement root, string fallback)
    {
        var source = TryObject(root, "source");
        return source.ValueKind == JsonValueKind.Object
            ? Coalesce(Text(source, "sourcePath"), fallback)
            : fallback;
    }

    private static JsonElement TryObject(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value)
        && value.ValueKind == JsonValueKind.Object
            ? value
            : default;

    private static IEnumerable<JsonElement> TryArray(
        JsonElement element,
        string name) =>
        element.ValueKind == JsonValueKind.Object
        && element.TryGetProperty(name, out var value)
        && value.ValueKind == JsonValueKind.Array
            ? value.EnumerateArray()
            : [];

    private static string Text(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object
        && element.TryGetProperty(name, out var value)
        && value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : string.Empty;

    private static string Scalar(JsonElement element, string name)
    {
        if (element.ValueKind != JsonValueKind.Object
            || !element.TryGetProperty(name, out var value)
            || value.ValueKind is JsonValueKind.Null
                or JsonValueKind.Undefined)
            return string.Empty;
        return value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : value.ToString();
    }

    private static string StatusLine(string status, string verification) =>
        "<div class='chips'>" + Chip("상태", status)
        + Chip("검증", verification) + "</div>";

    private static string Chip(string label, string value) =>
        string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : "<span class='chip'><small>" + H(label)
              + "</small>" + H(DisplayCode(value)) + "</span>";

    private static string LabelText(
        string label,
        string value,
        bool translate = false) =>
        string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : "<p><strong>" + H(label) + "</strong><span>"
              + H(translate ? TranslateNarrative(value) : value)
              + "</span></p>";

    private static string SummaryCard(string label, string value) =>
        "<div><span>" + H(label) + "</span><strong>"
        + H(value) + "</strong></div>";

    private static string DtDd(string label, string value) =>
        string.IsNullOrWhiteSpace(value)
            ? string.Empty
            : "<dt>" + H(label) + "</dt><dd><code>"
              + H(value) + "</code></dd>";

    private static string Coalesce(string? first, string? second) =>
        !string.IsNullOrWhiteSpace(first)
            ? first
            : second ?? string.Empty;

    private static string StatusClass(string value) =>
        value.Contains("REVIEW", StringComparison.OrdinalIgnoreCase)
            ? "review"
            : value.Contains("FAIL", StringComparison.OrdinalIgnoreCase)
                ? "failed"
                : "ok";

    private static string DisplayCode(string value) =>
        value.Trim().ToUpperInvariant() switch
        {
            "NEEDS_REVIEW" => "검토 필요 (NEEDS_REVIEW)",
            "UNASSESSED" => "미평가 (UNASSESSED)",
            "CAPTURED" => "원본 캡처 완료 (CAPTURED)",
            "COMPLETED" => "완료 (COMPLETED)",
            "FAILED" => "실패 (FAILED)",
            "VALID" => "유효 (VALID)",
            "INVALID" => "유효하지 않음 (INVALID)",
            "ELIGIBLE" => "집계 가능 (ELIGIBLE)",
            "INELIGIBLE" => "집계 불가 (INELIGIBLE)",
            "CONTROL" => "대조군 (CONTROL)",
            "REFERENCE" => "참조군 (REFERENCE)",
            "TREATMENT" => "시험군 (TREATMENT)",
            "COMPARATOR" => "비교군 (COMPARATOR)",
            "OTHER" => "기타 (OTHER)",
            _ => value,
        };

    private static string TranslateNarrative(string value)
    {
        if (string.IsNullOrWhiteSpace(value)
            || Regex.IsMatch(value, "[가-힣]"))
            return value;

        var exact = value switch
        {
            "TIU L5S3-01 [L] REPORT TEST FIND REASON NG BAKO AND VISUAL FINAL"
                => "TIU L5S3-01 [L] NG BAKO 원인 및 최종 외관 검사 보고서",
            "The workbook states a plan to separate Ass'y VP+Fr samples assembled by machine and by hand, check function and visual final results, and compare the results. It contains separate function-check and final-visual result tables with arm-specific Input, OK, NG-category, Total NG, and rate fields. A hidden 'After 1 day check again' row has no supplied input denominator, and the Decision section contains only its heading."
                => "이 문서는 기계 조립과 수작업 조립으로 나눈 Ass'y VP+Fr 시료의 기능 및 최종 외관 검사 결과를 비교하는 계획을 담고 있습니다. 기능 검사표와 최종 외관 검사표에는 조건별 Input, OK, NG 항목, 전체 NG 및 비율이 기록되어 있습니다. 숨김 처리된 '1일 후 재확인' 행에는 Input 분모가 없으며, Decision 영역에는 제목만 있습니다.",
            "1. Result check function: Ass'y by machine and Ass'y by hand"
                => "1. 기능 검사 결과: 기계 조립과 수작업 조립",
            "2. Result check visual final: Ass'y by machine and Ass'y by hand"
                => "2. 최종 외관 검사 결과: 기계 조립과 수작업 조립",
            "Find reaon NG bako and visual final"
                => "NG BAKO 및 최종 외관 불량 원인 확인",
            "Separate lot and check function, visual final; compare result."
                => "로트를 분리하여 기능과 최종 외관을 검사하고 결과를 비교합니다.",
            "source-stated two-condition comparative result table"
                => "원본에 명시된 두 조건 비교 결과표",
            "Ass'y VP+Fr by machine and ass'y by hand; separate lot and compare result."
                => "Ass'y VP+Fr를 기계 조립과 수작업 조립으로 나누고, 로트를 분리하여 결과를 비교합니다.",
            "Function-check results are reported for Ass'y by machine with Input 255 and Ass'y by hand with Input 320. The table preserves OK, NG BAKO categories, HEARING fields, Total NG, and their reported rates. A separate hidden 'After 1 day check again' condition has no usable denominator or numeric result set."
                => "기능 검사 결과는 기계 조립 Input 255개와 수작업 조립 Input 320개로 기록되어 있습니다. 표의 OK, NG BAKO 항목, HEARING 항목, 전체 NG 및 각 비율을 원본 그대로 보존했습니다. 별도의 숨김 '1일 후 재확인' 조건에는 사용할 수 있는 분모나 수치 결과가 없습니다.",
            "Visual-final results are reported for Ass'y by machine with Input 251 and Ass'y by hand with Input 311. The source reports OK, NG visual categories, Total NG, and percentage fields. Machine category-count cells are not supplied as nonempty captured cells, while their percentage cells and Total NG value are preserved separately."
                => "최종 외관 검사 결과는 기계 조립 Input 251개와 수작업 조립 Input 311개로 기록되어 있습니다. 원본에는 OK, 외관 NG 항목, 전체 NG 및 비율이 있습니다. 기계 조립의 항목별 건수 셀은 값이 없어 캡처되지 않았고, 해당 비율 셀과 전체 NG 값은 각각 보존했습니다.",
            "The assembly arms have different Input counts in both result tables."
                => "두 결과표 모두 조립 조건별 Input 수량이 서로 다릅니다.",
            "No randomization, specimen-level linkage, matching, acceptance criteria, or statistical analysis is supplied."
                => "무작위 배정, 시료 단위 연결, 매칭, 합격 기준 또는 통계 분석 정보가 제공되지 않았습니다.",
            "No randomization, matching, specimen-level linkage, replication protocol, or statistical analysis is supplied."
                => "무작위 배정, 매칭, 시료 단위 연결, 반복 시험 절차 또는 통계 분석 정보가 제공되지 않았습니다.",
            "The source does not label either assembly condition as Control, Reference, Baseline, Before, or Standard."
                => "원본은 어느 조립 조건도 대조군(Control), 참조군(Reference), 기준선(Baseline), 변경 전(Before) 또는 표준(Standard)으로 표시하지 않았습니다.",
            "The hidden 'After 1 day check again' row lacks an Input denominator and contains division-by-zero formula results; no numeric follow-up observations or comparison are drafted."
                => "숨김 처리된 '1일 후 재확인' 행에는 Input 분모가 없고 0으로 나누는 수식 결과가 있습니다. 따라서 후속 관측값이나 비교 결과를 수치로 작성하지 않았습니다.",
            "The hidden 'After 1 day check again' condition has blank source inputs and division-by-zero rate cells; cached formula zeros are not treated as measured observations."
                => "숨김 처리된 '1일 후 재확인' 조건은 원본 Input이 비어 있고 비율 셀에 0으로 나누는 수식이 있습니다. 캐시된 수식의 0값은 실제 측정값으로 처리하지 않았습니다.",
            "The visual-final Ass'y by machine category-count cells I40:L40 are not supplied as captured nonempty cells; their separate percentage cells are preserved without inferred numerators."
                => "최종 외관 검사의 기계 조립 항목별 건수 셀 I40:L40은 값이 없어 캡처되지 않았습니다. 별도의 비율 셀은 분자를 추정하지 않고 그대로 보존했습니다.",
            "For the visual-final Ass'y by hand row, Input 311 is not equal to OK 297 plus Total NG 7; the unreconciled residual is 7 and is not corrected or reclassified."
                => "최종 외관 검사의 수작업 조립 행에서 Input 311은 OK 297 + 전체 NG 7과 일치하지 않습니다. 미조정 차이 7은 수정하거나 재분류하지 않았습니다.",
            "The Decision section contains only the heading 'IV. Decision' and supplies no narrative decision or conclusion."
                => "Decision 영역에는 'IV. Decision' 제목만 있고 판단이나 결론 내용은 없습니다.",
            "The report date cell displays 22-Jul without a displayed year."
                => "보고서 날짜 셀은 연도 없이 22-Jul로 표시되어 있습니다.",
            "The filename identifier and literal report identifier differ."
                => "파일명 식별자와 보고서 본문 식별자가 서로 다릅니다.",
            _ => string.Empty,
        };
        if (!string.IsNullOrEmpty(exact))
            return exact;

        var match = Regex.Match(
            value,
            @"^The filename identifier '(?<file>.+)' differs from the literal report identifier '(?<report>.+)'; both identities are preserved without normalization\.$");
        if (match.Success)
            return $"파일명 식별자 '{match.Groups["file"].Value}'와 보고서 본문의 식별자 '{match.Groups["report"].Value}'가 다릅니다. 두 식별자를 정규화하지 않고 그대로 보존했습니다.";

        match = Regex.Match(
            value,
            @"^The compared assembly conditions have different Input counts: (?<first>[^ ]+) and (?<second>[^.]+)\.$");
        if (match.Success)
            return $"비교한 조립 조건의 검사 수량(Input)이 서로 다릅니다: {match.Groups["first"].Value}, {match.Groups["second"].Value}.";

        match = Regex.Match(
            value,
            @"^For (?<arm>.+), NG visual category-count cells (?<range>\S+) are not supplied as captured nonempty cells; zero numerators are not inferred from the percentage cells or Total NG\.$");
        if (match.Success)
            return $"{TranslateArm(match.Groups["arm"].Value)} 조건의 외관 NG 항목별 건수 셀 {match.Groups["range"].Value}가 비어 있어 캡처되지 않았습니다. 비율 셀이나 전체 NG에서 0건으로 추정하지 않았습니다.";

        match = Regex.Match(
            value,
            @"^For (?<arm>.+), Input (?<input>\S+) is not equal to OK (?<ok>\S+) plus Total NG (?<ng>\S+); the exact unreconciled residual is (?<residual>[^.]+)\.$");
        if (match.Success)
            return $"{TranslateArm(match.Groups["arm"].Value)} 조건에서 Input {match.Groups["input"].Value}은 OK {match.Groups["ok"].Value} 및 전체 NG {match.Groups["ng"].Value}의 합계와 일치하지 않습니다. 정확한 미조정 차이는 {match.Groups["residual"].Value}입니다.";

        match = Regex.Match(
            value,
            @"^The report date cell displays (?<report>.+) and the visual result date displays (?<visual>.+), both without a displayed year\.$");
        if (match.Success)
            return $"보고서 날짜는 {match.Groups["report"].Value}, 외관 결과 날짜는 {match.Groups["visual"].Value}로 표시되지만 두 날짜 모두 연도가 없습니다.";

        return value;
    }

    private static string TranslateArm(string value) =>
        value.Trim() switch
        {
            "Ass'y by machine" => "기계 조립",
            "Ass'y by hand" => "수작업 조립",
            _ => value,
        };

    private sealed record OutcomeColumn(
        string Key,
        string Label,
        List<JsonElement> Outcomes);

    private sealed record OutcomeCell(
        string Html,
        bool HasReviewIssue);

    private sealed record ReviewLimitation(
        string Original,
        string Display);

    private sealed record StudyReviewContext(
        List<string> AllIssues,
        Dictionary<string, List<string>> ArmIssues,
        Dictionary<string, List<string>> CellIssues);

    private static string H(string? value) =>
        WebUtility.HtmlEncode(value ?? string.Empty);

    private static string HA(string? value) =>
        H(value)
            .Replace("\r\n", "&#10;", StringComparison.Ordinal)
            .Replace("\n", "&#10;", StringComparison.Ordinal);

    private static string RenderMissingManifest(IngestWorkbookResult result) =>
        Page(
            "처리 내용 검증",
            "<h1>상세 manifest를 찾을 수 없습니다.</h1><p>"
            + H(result.ManifestPath)
            + "</p><p>처리 journal은 남아 있지만 Study·관측값 상세 파일이 없습니다.</p>");

    private static string Page(string title, string body) =>
        $$"""
        <!doctype html>
        <html lang="ko">
        <head>
          <meta charset="utf-8">
          <title>{{H(title)}}</title>
          <style>
            html { color-scheme: dark; }
            body { margin:0; padding:24px; background:#202020; color:#d6d6d6; font:14px "Segoe UI","Malgun Gothic",sans-serif; }
            header { display:flex; justify-content:space-between; gap:20px; align-items:center; border-bottom:1px solid #3a3a3a; padding-bottom:18px; }
            h1 { margin:4px 0 5px; color:#f0f0f0; font-size:22px; overflow-wrap:anywhere; }
            header p { margin:0; color:#8e8e8e; }
            .eyebrow { color:#a78bfa; font-size:10px; font-weight:800; }
            .status { display:inline-block; padding:5px 9px; border-radius:999px; font-size:11px; font-weight:800; white-space:nowrap; }
            .status.ok { color:#b8f3df; background:#203c34; border:1px solid #2c8c70; }
            .status.review { color:#ffe09b; background:#40351f; border:1px solid #8b6d2d; }
            .status.failed { color:#ffc0c0; background:#482626; border:1px solid #a74646; }
            .comparison-list { display:grid; gap:16px; margin-top:18px; }
            .comparison-card { overflow:hidden; background:#252525; border:1px solid #514168; border-radius:8px; }
            .comparison-heading { display:flex; justify-content:space-between; gap:16px; align-items:center; padding:14px 16px; border-bottom:1px solid #3a3a3a; }
            .comparison-heading span:first-child { color:#a78bfa; font-size:10px; font-weight:800; }
            .comparison-heading h2 { margin:4px 0 0; color:#eee; font-size:16px; }
            .comparison-table-wrap { max-width:100%; overflow:auto; }
            .comparison-table { width:max-content; min-width:100%; border-collapse:separate; border-spacing:0; background:#252525; font-size:13px; }
            .comparison-table th, .comparison-table td { min-width:92px; padding:10px; border-right:1px solid #3a3a3a; border-bottom:1px solid #3a3a3a; text-align:center; vertical-align:middle; white-space:nowrap; }
            .comparison-table th { position:sticky; top:0; z-index:3; background:#303030; color:#eee; text-align:left; }
            .comparison-table tbody tr:nth-child(even) td { background:#292929; }
            .cohort-column { position:sticky; left:0; z-index:2; width:72px; min-width:72px !important; background:#252525 !important; }
            th.cohort-column { z-index:5; background:#303030 !important; }
            .condition-column { position:sticky; left:92px; z-index:2; width:210px; min-width:210px !important; background:#252525 !important; text-align:left !important; }
            th.condition-column { z-index:5; background:#303030 !important; }
            .comparison-table tbody tr:nth-child(even) .condition-column,
            .comparison-table tbody tr:nth-child(even) .cohort-column { background:#292929 !important; }
            .cohort-badge { display:inline-block; padding:4px 7px; border-radius:4px; font-size:11px; font-weight:800; }
            .cohort-badge.control { color:#b8f3df; background:#203c34; border:1px solid #2c8c70; }
            .cohort-badge.comparison { color:#e2d6f5; background:#3c2a58; border:1px solid #68499a; }
            .metric-cell { color:#eee; font-weight:700; }
            .metric-cell a, .condition-column a { color:#c4b5fd; text-decoration:none; }
            .metric-cell a:hover, .condition-column a:hover { text-decoration:underline; }
            .rate-value a { color:#a9d8ff; }
            .value-separator { margin:0 5px; color:#666; font-weight:400; }
            .empty-value { color:#666; font-weight:400; }
            .metric-cell.review-issue { background:#3a321f !important; box-shadow:inset 0 0 0 1px #82672c; }
            .cohort-badge.review-issue { color:#ffe09b; background:#40351f; border-color:#8b6d2d; }
            .review-marker { display:inline-block; margin-left:6px; width:16px; height:16px; border-radius:50%; color:#241d0d; background:#e6c980; font-size:11px; font-weight:900; line-height:16px; text-align:center; cursor:help; }
            .warning { margin:14px; padding:12px; color:#e6c980; background:#302c22; border:1px solid #685a36; border-radius:6px; }
            footer { margin-top:18px; color:#8e8e8e; font-size:12px; }
          </style>
        </head>
        <body>{{body}}</body>
        </html>
        """;
}
