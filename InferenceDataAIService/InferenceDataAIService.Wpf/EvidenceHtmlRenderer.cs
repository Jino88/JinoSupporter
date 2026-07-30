using System.IO;
using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;

namespace InferenceDataAIService.Wpf;

internal static class EvidenceHtmlRenderer
{
    private const int MaxGridCells = 50_000;

    internal static string RenderAnswer(EvidenceAnswerSession session)
    {
        if (session.IsRelevance)
            return RenderRelevanceAnswer(session);
        if (session.IsContextual)
            return RenderContextualAnswer(session);

        var isHistoryAnswer = session.Citations.Any(citation =>
            citation.EvidenceId.StartsWith(
                "TF-EVD-",
                StringComparison.OrdinalIgnoreCase));
        var citations = session.Citations.Count == 0
            ? "<p>직접 열 수 있는 검증 근거 표가 없습니다.</p>"
            : "<ul>"
              + string.Join(
                  string.Empty,
                  session.Citations.Select(citation =>
                      $"<li><code>{H(citation.EvidenceId)}</code> — "
                      + $"{H(citation.FileName)} / "
                      + $"{H(citation.Sheet)}!{H(citation.Range)}</li>"))
              + "</ul>";
        var citationNote = isHistoryAnswer
            ? "TF-EVD는 전체 이력 인덱스에 저장된 원본 파일·시트·범위를 가리킵니다. 미리보기는 분석 요청에 보존된 셀이며 원본 Excel이 최종 근거입니다."
            : "왼쪽 근거 목록에서 EVD ID를 선택하면 Capture v2 표를 확인하거나 Excel의 정확한 시트·범위를 열 수 있습니다. 이미지 분석은 포함되지 않습니다.";
        return Page(
            isHistoryAnswer ? "전체 시험 이력 근거 답변" : "검토 DB 근거 답변",
            $"""
             <div class="summary">
               <strong>판정:</strong> {H(session.AnswerStatus)}
               &nbsp;·&nbsp; 관련 DATA {session.RelevantStudyCount:N0}건
               &nbsp;·&nbsp; 정량 사용 가능 효과 {session.EligibleEffectCount:N0}건
             </div>
             <pre>{H(session.Markdown)}</pre>
             <h2>클릭 가능한 원본 근거</h2>
             {citations}
             <p class="note">{H(citationNote)}</p>
             """);
    }

    private static string RenderRelevanceAnswer(EvidenceAnswerSession session)
    {
        using var document = JsonDocument.Parse(session.RawJson);
        var root = document.RootElement;
        var interpretation = root.GetProperty("queryInterpretation");
        var studies = root.GetProperty("studies");
        var reviewMatrix = RenderReviewMatrix(studies);
        return Page(
            "관련 시험 비교표",
            $"""
             <div class="comparison-description">
               <strong>비교군</strong>
               <div>{RenderArrayChips(interpretation, "conditions")}</div>
             </div>
             {reviewMatrix}
             """);
    }

    private static string RenderContextualAnswer(EvidenceAnswerSession session)
    {
        using var document = JsonDocument.Parse(session.RawJson);
        var root = document.RootElement;
        var intent = root.GetProperty("intent");
        var coverage = root.GetProperty("coverage");
        var answerStatus = JsonValue(root, "answerStatus");
        var confidence = JsonValue(root, "confidence");
        var statusLabel = answerStatus switch
        {
            "CONTEXTUAL_AI_ANSWERED" => "근거 기반 답변",
            "CONTEXTUAL_AI_PARTIAL" => "부분 답변",
            "CONTEXTUAL_AI_INSUFFICIENT_EVIDENCE" => "근거 부족",
            _ => answerStatus,
        };
        var confidenceLabel = confidence switch
        {
            "HIGH" => "높음",
            "MEDIUM" => "보통",
            "LOW" => "낮음",
            _ => confidence,
        };
        var candidateCount = coverage.TryGetProperty(
            "candidateStudyCount",
            out var candidateCountElement)
            ? candidateCountElement.GetInt32()
            : session.RelevantStudyCount;
        var excludedCount = coverage.TryGetProperty(
            "excludedCandidateCount",
            out var excludedCountElement)
            ? excludedCountElement.GetInt32()
            : Math.Max(0, candidateCount - session.RelevantStudyCount);
        var coreCitationCount = coverage.TryGetProperty(
            "citationCount",
            out var coreCitationCountElement)
            ? coreCitationCountElement.GetInt32()
            : session.Citations.Count;

        var conditions = RenderInlineValues(intent, "conditions");
        var metrics = RenderInlineValues(intent, "metrics");
        var findings = root.GetProperty("findings").GetArrayLength() == 0
            ? "<p class='empty'>질문에 직접 연결되는 핵심 판단을 만들 수 없었습니다.</p>"
            : "<div class='finding-list'>"
              + string.Join(
                  string.Empty,
                  root.GetProperty("findings")
                      .EnumerateArray()
                      .Select(item =>
                          "<article class='finding'>"
                          + $"<div>{H(JsonValue(item, "statement"))}</div>"
                          + (string.IsNullOrWhiteSpace(JsonValue(item, "significance"))
                              ? string.Empty
                              : $"<p>{H(JsonValue(item, "significance"))}</p>")
                          + $"<small>근거 {H(RenderIds(item, "evidenceIds"))}</small>"
                          + "</article>"))
              + "</div>";

        var trendRows = root.GetProperty("trendRows");
        var trend = trendRows.GetArrayLength() == 0
            ? "<p class='empty'>동일 지표로 비교 가능한 날짜별 관측값이 확인되지 않았습니다.</p>"
            : "<div class='table-wrap'><table class='answer-table'>"
              + "<thead><tr><th>날짜</th><th>조건</th><th>지표</th>"
              + "<th>값</th><th>해석</th></tr></thead><tbody>"
              + string.Join(
                  string.Empty,
                  trendRows.EnumerateArray().Select(item =>
                      "<tr>"
                      + $"<td>{H(JsonValue(item, "date"))}</td>"
                      + $"<td>{H(JsonValue(item, "condition"))}</td>"
                      + $"<td>{H(JsonValue(item, "metric"))}</td>"
                      + $"<td class='value'>{H(JsonValue(item, "value"))}</td>"
                      + $"<td>{H(JsonValue(item, "interpretation"))}</td>"
                      + "</tr>"))
              + "</tbody></table></div>";

        var limitationsElement = root.GetProperty("limitations");
        var limitations = limitationsElement.GetArrayLength() == 0
            ? "<p class='empty'>추가로 기록된 제한이 없습니다.</p>"
            : "<ul class='limitations'>"
              + string.Join(
                  string.Empty,
                  limitationsElement.EnumerateArray()
                      .Select(value => $"<li>{H(value.GetString())}</li>"))
              + "</ul>";

        var citations = session.Citations.Count == 0
            ? "<p class='empty'>직접 관련 Study의 원본 근거가 없습니다.</p>"
            : "<div class='table-wrap'><table class='evidence-table'>"
              + "<thead><tr><th>원본</th><th>시트·범위</th><th>근거 ID</th></tr></thead><tbody>"
              + string.Join(
                  string.Empty,
                  session.Citations.Select(citation =>
                      "<tr>"
                      + $"<td>{H(citation.FileName)}</td>"
                      + $"<td>{H(citation.Sheet)}!{H(citation.Range)}</td>"
                      + $"<td><code>{H(citation.EvidenceId)}</code></td>"
                      + "</tr>"))
              + "</tbody></table></div>";

        return Page(
            "문맥 AI 근거 답변",
            $"""
             <div class="summary status-grid">
               <div><span>판정</span><strong>{H(statusLabel)}</strong></div>
               <div><span>신뢰도</span><strong>{H(confidenceLabel)}</strong></div>
               <div><span>관련 Study</span><strong>{session.RelevantStudyCount:N0}건</strong></div>
               <div><span>후보 제외</span><strong>{excludedCount:N0}건</strong></div>
             </div>
             <section class="answer-card">
               <div class="eyebrow">질문에 대한 답</div>
               <p class="direct-answer">{H(JsonValue(root, "directAnswer"))}</p>
             </section>
             <section>
               <h2>AI가 이해한 질문</h2>
               <dl class="intent-grid">
                 <dt>대상</dt><dd>{H(JsonValue(intent, "subject"))}</dd>
                 <dt>조건</dt><dd>{conditions}</dd>
                 <dt>지표</dt><dd>{metrics}</dd>
                 <dt>비교</dt><dd>{H(JsonValue(intent, "comparison"))}</dd>
                 <dt>시간축</dt><dd>{H(JsonValue(intent, "timeScope"))}</dd>
               </dl>
               <p class="note">{H(JsonValue(root, "relevanceAssessment"))}</p>
             </section>
             <section><h2>핵심 판단</h2>{findings}</section>
             <section><h2>비교 가능한 관측값</h2>{trend}</section>
             <section><h2>해석 한계</h2>{limitations}</section>
             <section>
               <h2>관련 Study 원본 근거</h2>
               <p class="note">관련 원본 근거 {session.Citations.Count:N0}건 중 수치 판단에 직접 사용한 핵심 근거는 {coreCitationCount:N0}건입니다.</p>
               {citations}
             </section>
             <p class="warning">NEEDS_REVIEW 이력의 관측값을 설명한 결과입니다. 원본 검토 전에는 승인된 효과나 인과 결론으로 사용하면 안 됩니다.</p>
             """);
    }

    private static string RenderInlineValues(JsonElement element, string property)
    {
        if (!element.TryGetProperty(property, out var values)
            || values.ValueKind != JsonValueKind.Array
            || values.GetArrayLength() == 0)
            return "<span class='muted'>미확정</span>";
        return string.Join(
            " ",
            values.EnumerateArray()
                .Select(value => $"<span class='chip'>{H(value.GetString())}</span>"));
    }

    private static string RenderArrayChips(JsonElement element, string property)
    {
        if (!element.TryGetProperty(property, out var values)
            || values.ValueKind != JsonValueKind.Array
            || values.GetArrayLength() == 0)
            return "<span class='muted'>없음</span>";
        return string.Join(
            " ",
            values.EnumerateArray()
                .Select(value => $"<span class='chip'>{H(value.GetString())}</span>"));
    }

    private static string RenderNamedChips(
        JsonElement element,
        string property,
        string nameProperty,
        int maximum)
    {
        if (!element.TryGetProperty(property, out var values)
            || values.ValueKind != JsonValueKind.Array
            || values.GetArrayLength() == 0)
            return "<span class='muted'>없음</span>";
        var names = values.EnumerateArray()
            .Select(value => JsonValue(value, nameProperty))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var chips = string.Join(
            string.Empty,
            names.Take(maximum)
                .Select(value => $"<span class='chip'>{H(value)}</span>"));
        return chips
               + (names.Count > maximum
                   ? $"<span class='more-chip'>+{names.Count - maximum:N0}</span>"
                   : string.Empty);
    }

    private static string RenderNamedItems(
        JsonElement element,
        string property,
        string nameProperty,
        string detailProperty)
    {
        if (!element.TryGetProperty(property, out var values)
            || values.ValueKind != JsonValueKind.Array
            || values.GetArrayLength() == 0)
            return "<span class='muted'>명시 없음</span>";
        return "<div class='finding-list'>"
               + string.Join(
                   string.Empty,
                   values.EnumerateArray().Take(12).Select(value =>
                   {
                       var name = JsonValue(value, nameProperty);
                       var detail = JsonValue(value, detailProperty);
                       return "<div class='finding'>"
                              + $"<strong>{H(name)}</strong>"
                              + (string.IsNullOrWhiteSpace(detail)
                                  ? string.Empty
                                  : $"<small>{H(detail)}</small>")
                              + "</div>";
                   }))
               + "</div>";
    }

    private static string RenderReviewMatrix(JsonElement studies)
    {
        var reviewPoints = new List<ReviewPoint>();
        var studyIndex = 0;
        foreach (var study in studies.EnumerateArray())
        {
            var currentStudyIndex = studyIndex++;
            var rank = study.TryGetProperty(
                           "retrievalRank",
                           out var rankElement)
                       && rankElement.TryGetInt32(out var parsedRank)
                ? parsedRank
                : currentStudyIndex + 1;
            if (!study.TryGetProperty("rawDataPoints", out var points)
                || points.ValueKind != JsonValueKind.Array)
                continue;
            var rawPoints = points.EnumerateArray()
                .Select((point, index) =>
                    RawDataPointView.FromJson(point, index))
                .ToList();
            var fileName = JsonValue(study, "fileName");
            var sourcePath = JsonValue(study, "sourcePath");
            var studyDate = CompactRawCondition(JsonValue(study, "date"));
            var studyGroup = JsonValue(study, "studyGroup");
            var sourceTitles = ReviewSourceTitles(study, rawPoints);
            var separateRows = rawPoints
                .Where(point =>
                    point.Metric.Contains(
                        "separate",
                        StringComparison.OrdinalIgnoreCase)
                    || point.Metric.Contains(
                        "분리",
                        StringComparison.Ordinal))
                .Select(ReviewRowKey)
                .ToHashSet(StringComparer.Ordinal);
            foreach (var rawPoint in rawPoints)
            {
                var metricKey = ReviewMetricKey(
                    rawPoint,
                    separateRows.Contains(ReviewRowKey(rawPoint)));
                if (metricKey is null
                    || IsUnwantedReviewPercentage(rawPoint, metricKey))
                    continue;
                reviewPoints.Add(new ReviewPoint(
                    currentStudyIndex,
                    rank,
                    fileName,
                    sourcePath,
                    studyDate,
                    studyGroup,
                    sourceTitles.GetValueOrDefault(
                        ReviewSourceKey(rawPoint),
                        ReviewSourceLocation(rawPoint)),
                    metricKey,
                    rawPoint));
            }
            AppendDerivedFunctionNgPoints(
                reviewPoints,
                rawPoints,
                currentStudyIndex,
                rank,
                fileName,
                sourcePath,
                studyDate,
                studyGroup,
                sourceTitles);
        }

        if (reviewPoints.Count == 0)
            return "<div class='empty'>통합표로 표시할 원본 수량·불량 지표가 없습니다.</div>";

        var configuredColumns = new[]
        {
            new ReviewMetricColumn("input", "검사 수량"),
            new ReviewMetricColumn("ok", "OK"),
            new ReviewMetricColumn("total_ng", "전체 NG"),
            new ReviewMetricColumn("total_ng_rate", "전체 NG Rate"),
            new ReviewMetricColumn("sigma_ng", "Sigma NG"),
            new ReviewMetricColumn("sigma_ng_rate", "Sigma NG Rate"),
            new ReviewMetricColumn("hearing", "Hearing NG"),
            new ReviewMetricColumn("hearing_ng_rate", "Hearing NG Rate"),
            new ReviewMetricColumn("noise", "Noise"),
            new ReviewMetricColumn("touch", "Touch"),
            new ReviewMetricColumn("separate", "VP/CD 분리"),
            new ReviewMetricColumn("separate_rate", "VP/CD NG Rate")
        };
        var sourceReviews = reviewPoints
            .GroupBy(point => new
            {
                point.StudyIndex,
                point.RawPoint.EvidenceId,
                point.RawPoint.TableId,
                point.RawPoint.Sheet,
                point.RawPoint.Range
            })
            .OrderBy(group => group.First().Rank)
            .ThenBy(group => group.Key.StudyIndex)
            .ThenBy(group => group.Min(point => point.RawPoint.Index))
            .Select(group =>
            {
                var points = group.ToList();
                var family = ReviewFamilyFor(points);
                return new ReviewSource(
                    family.Key,
                    family.Title,
                    points);
            })
            .Where(source => HasReviewOutcome(source.Points))
            .ToList();
        if (sourceReviews.Count == 0)
            return "<div class='empty'>수량 외에 비교할 결과 지표가 있는 원본 검토표가 없습니다.</div>";
        var sources = sourceReviews
            .OrderByDescending(ReviewSourceDate)
            .ThenBy(source => source.First.Rank)
            .ThenBy(source => source.First.StudyIndex)
            .ThenBy(source => source.First.RawPoint.Index)
            .ToList();
        var content = new StringBuilder("<div class='review-set-list'>");
        foreach (var source in sources)
        {
            var sourceFirst = source.First;
            var rows = source.Points
                .GroupBy(point => point.RawPoint.Row)
                .Where(HasReviewOutcome)
                .OrderBy(group => group.Key)
                .ToList();
            var currentDate = sourceFirst.Date;
            var currentType = string.Empty;
            var preparedRows = new List<(
                IGrouping<int, ReviewPoint> Points,
                (string Date, string Type, string Content) Condition,
                bool IsControl)>();
            foreach (var row in rows)
            {
                var rawCondition = row.Select(point => point.RawPoint.Condition)
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
                    ?? string.Empty;
                var condition = SplitReviewCondition(
                    rawCondition,
                    currentDate,
                    currentType);
                if (!string.IsNullOrWhiteSpace(condition.Date)
                    && condition.Date != "-")
                    currentDate = condition.Date;
                if (!string.IsNullOrWhiteSpace(condition.Type)
                    && condition.Type != "-")
                    currentType = condition.Type;
                preparedRows.Add((
                    row,
                    condition,
                    IsControlReviewCondition(rawCondition)));
            }
            var visibleColumns = configuredColumns
                .Where(column => preparedRows.Any(preparedRow =>
                    preparedRow.Points.Any(point => string.Equals(
                        point.MetricKey,
                        column.Key,
                        StringComparison.Ordinal))))
                .ToList();

            var sourceHref = ReviewExcelHref(
                sourceFirst,
                sourceFirst.RawPoint.Range);
            content.Append("<section class='review-set-card'><div class='review-set-heading'><a href='")
                .Append(H(sourceHref))
                .Append("' title='클릭하여 원본 Excel 표 열기'><strong>")
                .Append(H(CompactReviewFamilyTitle(source.FamilyTitle)))
                .Append("</strong><span>")
                .Append(H(ReviewSetTitle(source)))
                .Append("</span><small>대조군 ")
                .Append(preparedRows.Count(row => row.IsControl))
                .Append(" · 비교군 ")
                .Append(preparedRows.Count(row => !row.IsControl))
                .Append("</small></a></div><div class='review-set-table-wrap'>")
                .Append("<table class='review-set-table'><thead><tr>")
                .Append("<th class='review-cohort'>구분</th>")
                .Append("<th class='review-date'>날짜</th>")
                .Append("<th class='review-comparison'>조건</th>");
            foreach (var column in visibleColumns)
            {
                content.Append("<th class='review-metric-header'>")
                    .Append(H(column.Label))
                    .Append("</th>");
            }
            content.Append("</tr></thead><tbody>");

            foreach (var preparedRow in preparedRows)
            {
                var row = preparedRow.Points;
                var condition = preparedRow.Condition;
                var comparisonHref = ReviewExcelHref(
                    sourceFirst,
                    sourceFirst.RawPoint.Range);
                content.Append("<tr class='")
                    .Append(preparedRow.IsControl
                        ? "review-control-row"
                        : "review-comparison-row")
                    .Append("'><td class='review-cohort'><span class='")
                    .Append(preparedRow.IsControl
                        ? "cohort-badge control"
                        : "cohort-badge comparison")
                    .Append("'>")
                    .Append(preparedRow.IsControl ? "대조군" : "비교군")
                    .Append("</span></td><td class='review-date'><strong>")
                    .Append(H(condition.Date))
                    .Append("</strong></td><td class='review-comparison'><a class='review-comparison-link' href='")
                    .Append(H(comparisonHref))
                    .Append("' title='클릭하여 원본 Excel 표 열기'><strong>")
                    .Append(H(condition.Type))
                    .Append("</strong>")
                    .Append(string.IsNullOrWhiteSpace(condition.Content)
                        ? string.Empty
                        : $"<small>{H(condition.Content)}</small>")
                    .Append("</a></td>");
                foreach (var column in visibleColumns)
                {
                    var values = row
                        .Where(point => string.Equals(
                            point.MetricKey,
                            column.Key,
                            StringComparison.Ordinal))
                        .GroupBy(point => new
                        {
                            point.RawPoint.DisplayValue,
                            point.RawPoint.Metric,
                            point.RawPoint.Coordinate
                        })
                        .Select(value => value.Key)
                        .OrderBy(value => value.Metric)
                        .ToList();
                    content.Append("<td class='review-metric-cell'>");
                    if (values.Count == 0)
                    {
                        content.Append("<span class='empty-value'>—</span>");
                    }
                    else
                    {
                        var detail = string.Join(
                            " | ",
                            values.Select(value =>
                                $"{value.Metric} = {value.DisplayValue} ({value.Coordinate})"));
                        var targetRange = values.Count == 1
                            ? values[0].Coordinate
                            : sourceFirst.RawPoint.Range;
                        var metricHref = ReviewExcelHref(
                            sourceFirst,
                            targetRange);
                        content.Append("<a class='review-value-link' href='")
                            .Append(H(metricHref))
                            .Append("' title='")
                            .Append(H($"클릭하여 Excel 열기 · {detail}"))
                            .Append("'><strong>")
                            .Append(H(string.Join(
                                " / ",
                                values.Select(value => value.DisplayValue))))
                            .Append("</strong></a>");
                    }
                    content.Append("</td>");
                }
                content.Append("</tr>");
            }
            content.Append("</tbody></table></div></section>");
        }
        content.Append("</div>");
        return content.ToString();
    }

    private static bool HasReviewOutcome(
        IEnumerable<ReviewPoint> points) =>
        points.Any(point => point.MetricKey is not "review_no" and not "input");

    private static string CompactReviewFamilyTitle(string title) =>
        title.Replace(" 불량 검토", string.Empty, StringComparison.Ordinal)
            .Replace(" 검토", string.Empty, StringComparison.Ordinal);

    private static bool IsControlReviewCondition(string condition) =>
        RawConditionValues(condition).Any(value =>
            value.Contains("normal", StringComparison.OrdinalIgnoreCase)
            || value.Contains("control", StringComparison.OrdinalIgnoreCase)
            || value.Contains("baseline", StringComparison.OrdinalIgnoreCase)
            || value.Contains("reference", StringComparison.OrdinalIgnoreCase)
            || value.Contains("standard", StringComparison.OrdinalIgnoreCase)
            || value.Contains("대조", StringComparison.Ordinal)
            || value.Contains("기준", StringComparison.Ordinal));

    private static string ReviewSetTitle(ReviewSource source)
    {
        var title = source.First.ReviewTitle;
        if (string.IsNullOrWhiteSpace(title)
            || title.Contains('!') && title.Contains(':'))
            title = source.First.StudyGroup;
        if (string.IsNullOrWhiteSpace(title))
            title = "원본 비교 세트";
        return title.Length > 90
            ? title[..87] + "..."
            : title;
    }

    private static string ReviewExcelHref(
        ReviewPoint source,
        string range) =>
        "inference-excel://open/?source="
        + Uri.EscapeDataString(source.SourcePath)
        + "&sheet="
        + Uri.EscapeDataString(source.RawPoint.Sheet)
        + "&range="
        + Uri.EscapeDataString(range);

    private static (string Key, string Title) ReviewFamilyFor(
        IReadOnlyList<ReviewPoint> points)
    {
        var metricKeys = points
            .Select(point => point.MetricKey)
            .ToHashSet(StringComparer.Ordinal);
        var sourceText = string.Join(
            " ",
            points.SelectMany(point => new[]
            {
                point.ReviewTitle,
                point.RawPoint.Metric
            })).ToLowerInvariant();
        if (metricKeys.Contains("separate")
            || metricKeys.Contains("separate_rate")
            || sourceText.Contains("separate", StringComparison.Ordinal)
            || sourceText.Contains("분리", StringComparison.Ordinal))
            return ("vp_cd_separation", "VP/CD 분리 검토");
        if (metricKeys.Contains("hearing")
            || metricKeys.Contains("hearing_ng_rate")
            || metricKeys.Contains("noise")
            || metricKeys.Contains("touch")
            || metricKeys.Contains("sigma_ng")
            || metricKeys.Contains("sigma_ng_rate")
            || sourceText.Contains("audiobus", StringComparison.Ordinal)
            || sourceText.Contains("function", StringComparison.Ordinal))
            return ("function_hearing", "기능·Hearing 불량 검토");
        if (sourceText.Contains("vision", StringComparison.Ordinal))
            return ("vision", "외관·Vision 불량 검토");
        if (sourceText.Contains("dyne", StringComparison.Ordinal)
            || sourceText.Contains("tension", StringComparison.Ordinal))
            return ("dyne_tension", "Dyne·Tension 검토");
        if (sourceText.Contains("air leak", StringComparison.Ordinal))
            return ("air_leak", "Air Leak 검토");
        if (metricKeys.Count == 1 && metricKeys.Contains("input"))
            return ("test_quantity", "시험 수량·조건 검토");
        return ("quantity_ng", "수량·불량 검토");
    }

    private static string ReviewSourceDate(ReviewSource source)
    {
        foreach (var condition in source.Points
                     .Select(point => point.RawPoint.Condition)
                     .Where(value => !string.IsNullOrWhiteSpace(value)))
        {
            var date = SplitReviewCondition(
                    condition,
                    source.First.Date)
                .Date;
            if (!string.IsNullOrWhiteSpace(date) && date != "-")
                return date;
        }
        return source.First.Date;
    }

    private static Dictionary<string, string> ReviewSourceTitles(
        JsonElement study,
        IReadOnlyList<RawDataPointView> points)
    {
        var groups = points
            .GroupBy(ReviewSourceKey)
            .OrderBy(group => group.Min(point => point.Index))
            .ToList();
        var titles = study.TryGetProperty("titles", out var titleValues)
                     && titleValues.ValueKind == JsonValueKind.Array
            ? titleValues.EnumerateArray()
                .Select(value => value.GetString() ?? string.Empty)
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .ToList()
            : [];
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var titleOffset = titles.Count >= groups.Count
            ? titles.Count - groups.Count
            : -1;
        for (var index = 0; index < groups.Count; index++)
        {
            var group = groups[index].ToList();
            var metricTitle = group
                .Select(point => ReviewTitleFromMetric(point.Metric))
                .FirstOrDefault(value =>
                    !string.IsNullOrWhiteSpace(value));
            var mappedTitle = titleOffset >= 0
                ? titles[titleOffset + index]
                : string.Empty;
            result[groups[index].Key] =
                !string.IsNullOrWhiteSpace(metricTitle)
                    ? metricTitle
                    : !string.IsNullOrWhiteSpace(mappedTitle)
                        ? mappedTitle
                        : ReviewSourceLocation(group[0]);
        }
        return result;
    }

    private static string ReviewTitleFromMetric(string metric)
    {
        var separator = metric.LastIndexOf(
            " / No",
            StringComparison.OrdinalIgnoreCase);
        return separator > 0
            ? metric[..separator].Trim()
            : string.Empty;
    }

    private static string ReviewSourceKey(RawDataPointView point) =>
        string.Join(
            '\u001f',
            point.EvidenceId,
            point.TableId,
            point.Sheet,
            point.Range);

    private static string ReviewSourceLocation(RawDataPointView point) =>
        string.IsNullOrWhiteSpace(point.Sheet)
            ? point.Range
            : $"{point.Sheet}!{point.Range}";

    private static string? ReviewMetricKey(
        RawDataPointView point,
        bool separateRow)
    {
        var metric = point.Metric.Trim().ToLowerInvariant();
        if (!string.IsNullOrWhiteSpace(
                ReviewTitleFromMetric(point.Metric)))
            return "review_no";
        if (metric.Contains("ng rate", StringComparison.Ordinal))
        {
            if (separateRow) return "separate_rate";
            if (metric.Contains("sigma", StringComparison.Ordinal))
                return "sigma_ng_rate";
            if (metric.Contains("hearing", StringComparison.Ordinal))
                return "hearing_ng_rate";
            return "total_ng_rate";
        }
        var hasNg = metric.Contains("ng", StringComparison.Ordinal);
        var hasQuantity = metric.Contains("q'ty", StringComparison.Ordinal)
                          || metric.Contains("qty", StringComparison.Ordinal)
                          || metric.Contains("quantity", StringComparison.Ordinal);
        if (hasQuantity && hasNg) return "total_ng";
        if (metric.Contains("input", StringComparison.Ordinal) || hasQuantity)
            return "input";
        if (metric.Contains("ok", StringComparison.Ordinal) && !hasNg)
            return "ok";
        if (metric.Contains("separate", StringComparison.Ordinal)
            || metric.Contains("분리", StringComparison.Ordinal))
            return "separate";
        if (metric.Contains("noise", StringComparison.Ordinal)) return "noise";
        if (metric.Contains("touch", StringComparison.Ordinal)) return "touch";
        if (metric.Contains("hearing", StringComparison.Ordinal))
            return "hearing";
        if (metric.Contains("sigma", StringComparison.Ordinal)
            && (metric.Contains("total ng", StringComparison.Ordinal)
                || metric.Contains("ng rate", StringComparison.Ordinal)
                || metric is "ng sigma" or "sigma ng"))
            return "sigma_ng";
        if (metric.Contains("total ng", StringComparison.Ordinal)
            || metric.Contains("ng rate", StringComparison.Ordinal)
            || metric is "ng" or "q'ty ng")
            return "total_ng";
        return null;
    }

    private static bool IsUnwantedReviewPercentage(
        RawDataPointView point,
        string metricKey)
    {
        if (metricKey.EndsWith("_rate", StringComparison.Ordinal))
            return false;
        return point.Metric.Contains(
                   "rate",
                   StringComparison.OrdinalIgnoreCase)
               || point.DisplayValue.Contains('%');
    }

    private static void AppendDerivedFunctionNgPoints(
        List<ReviewPoint> reviewPoints,
        IReadOnlyList<RawDataPointView> rawPoints,
        int studyIndex,
        int rank,
        string fileName,
        string sourcePath,
        string studyDate,
        string studyGroup,
        IReadOnlyDictionary<string, string> sourceTitles)
    {
        var existingRows = reviewPoints
            .Where(point => point.StudyIndex == studyIndex)
            .GroupBy(point => ReviewRowKey(point.RawPoint))
            .ToDictionary(
                group => group.Key,
                group => group.ToList(),
                StringComparer.Ordinal);
        foreach (var rawRow in rawPoints.GroupBy(ReviewRowKey))
        {
            if (!existingRows.TryGetValue(
                    rawRow.Key,
                    out var existingPoints))
                continue;
            var inputPoint = existingPoints.FirstOrDefault(point =>
                string.Equals(
                    point.MetricKey,
                    "input",
                    StringComparison.Ordinal)
                && TryReviewNumber(
                    point.RawPoint.DisplayValue,
                    out var input)
                && input > 0);
            if (inputPoint is null
                || !TryReviewNumber(
                    inputPoint.RawPoint.DisplayValue,
                    out var inputValue)
                || inputValue <= 0)
                continue;

            var existingTotalPoint = existingPoints.FirstOrDefault(point =>
                string.Equals(
                    point.MetricKey,
                    "total_ng",
                    StringComparison.Ordinal)
                && TryReviewNumber(
                    point.RawPoint.DisplayValue,
                    out _));
            double totalNg;
            if (existingTotalPoint is not null)
            {
                TryReviewNumber(
                    existingTotalPoint.RawPoint.DisplayValue,
                    out totalNg);
            }
            else
            {
                var components = rawRow
                    .Where(IsFunctionNgComponent)
                    .GroupBy(
                        point => point.Metric.Trim(),
                        StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.First())
                    .Select(point => TryReviewNumber(
                            point.DisplayValue,
                            out var value)
                        ? value
                        : (double?)null)
                    .ToList();
                if (components.Count < 2
                    || components.Any(value => !value.HasValue))
                    continue;
                totalNg = components.Sum(value => value!.Value);
            }

            var template = rawRow.First();
            var reviewTitle = sourceTitles.GetValueOrDefault(
                ReviewSourceKey(template),
                ReviewSourceLocation(template));
            var nextIndex = rawRow.Max(point => point.Index) + 1;
            ReviewPoint DerivedPoint(
                string metricKey,
                string metric,
                string displayValue) =>
                new(
                    studyIndex,
                    rank,
                    fileName,
                    sourcePath,
                    studyDate,
                    studyGroup,
                    reviewTitle,
                    metricKey,
                    new RawDataPointView(
                        nextIndex++,
                        template.EvidenceId,
                        template.TableId,
                        template.Sheet,
                        template.Range,
                        template.Row,
                        template.Condition,
                        metric,
                        string.Empty,
                        displayValue,
                        template.Range));

            if (existingTotalPoint is null)
            {
                reviewPoints.Add(DerivedPoint(
                    "total_ng",
                    "Total NG (원본 수식 복원)",
                    FormatReviewNumber(totalNg)));
            }
            if (!existingPoints.Any(point => string.Equals(
                    point.MetricKey,
                    "ok",
                    StringComparison.Ordinal)))
            {
                reviewPoints.Add(DerivedPoint(
                    "ok",
                    "OK (원본 수식 복원)",
                    FormatReviewNumber(inputValue - totalNg)));
            }
            if (!existingPoints.Any(point => string.Equals(
                    point.MetricKey,
                    "total_ng_rate",
                    StringComparison.Ordinal)))
            {
                reviewPoints.Add(DerivedPoint(
                    "total_ng_rate",
                    "Total NG Rate (원본 수식 복원)",
                    (totalNg / inputValue).ToString(
                        "0.00%",
                        CultureInfo.InvariantCulture)));
            }
        }
    }

    private static bool IsFunctionNgComponent(RawDataPointView point)
    {
        var metric = point.Metric
            .Trim()
            .Replace(" ", string.Empty, StringComparison.Ordinal)
            .ToLowerInvariant();
        return metric is "airleak"
            or "spl"
            or "thd"
            or "spl+thd"
            or "spl+thd+f0"
            or "noise"
            or "touch";
    }

    private static bool TryReviewNumber(
        string displayValue,
        out double value) =>
        double.TryParse(
            displayValue
                .Trim()
                .Replace(",", string.Empty, StringComparison.Ordinal),
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out value);

    private static string FormatReviewNumber(double value) =>
        value.ToString("0.##", CultureInfo.InvariantCulture);

    private static string ReviewRowKey(RawDataPointView point) =>
        $"{ReviewSourceKey(point)}\u001f{point.Row}";

    private static string RenderRawDataPoints(JsonElement study)
    {
        if (!study.TryGetProperty("rawDataPoints", out var points)
            || points.ValueKind != JsonValueKind.Array
            || points.GetArrayLength() == 0)
            return "<div class='empty'>캡처된 원본 수치 행이 없습니다. 왼쪽 근거 범위를 열어 원본 표를 확인하세요.</div>";
        var rawPoints = points.EnumerateArray()
            .Select((point, index) => RawDataPointView.FromJson(point, index))
            .Where(point => !string.IsNullOrWhiteSpace(point.Metric))
            .ToList();
        if (rawPoints.Count == 0)
            return "<div class='empty'>표시할 지표가 없습니다. 왼쪽 근거 범위를 열어 원본 표를 확인하세요.</div>";

        var sourceBlocks = new StringBuilder();
        foreach (var source in rawPoints
                     .GroupBy(point => new
                     {
                         point.EvidenceId,
                         point.TableId,
                         point.Sheet,
                         point.Range
                     })
                     .OrderBy(group => group.Min(point => point.Index)))
        {
            var sourcePoints = source.ToList();
            var metrics = sourcePoints
                .GroupBy(point => new
                {
                    point.Metric,
                    point.Unit,
                    Column = ExcelColumnLetters(point.Coordinate)
                })
                .Select(group => new MetricColumn(
                    group.Key.Metric,
                    group.Key.Unit,
                    group.Key.Column,
                    ExcelColumnNumber(group.Key.Column),
                    group.Min(point => point.Index)))
                .OrderBy(metric => metric.ColumnNumber)
                .ThenBy(metric => metric.FirstIndex)
                .ToList();

            var header = new StringBuilder(
                "<thead><tr><th class='raw-condition'>기준</th>"
                + "<th class='raw-condition'>시험 조건</th>");
            foreach (var metric in metrics)
            {
                header.Append("<th class='raw-metric'><strong>")
                    .Append(H(metric.Metric))
                    .Append("</strong>");
                if (!string.IsNullOrWhiteSpace(metric.Unit))
                    header.Append("<br><small>[")
                        .Append(H(metric.Unit))
                        .Append("]</small>");
                header.Append("</th>");
            }
            header.Append("</tr></thead>");

            var rows = new StringBuilder("<tbody>");
            foreach (var row in sourcePoints
                         .GroupBy(point => point.Row)
                         .OrderBy(group => group.Key))
            {
                var condition = SplitRawCondition(
                    row.Select(point => point.Condition)
                        .FirstOrDefault(value =>
                            !string.IsNullOrWhiteSpace(value))
                    ?? string.Empty);
                rows.Append("<tr><td class='raw-condition'><strong>")
                    .Append(H(condition.First))
                    .Append("</strong><small>원본 행 ")
                    .Append(row.Key > 0 ? H(row.Key) : "-")
                    .Append("</small></td><td class='raw-condition'>")
                    .Append(H(condition.Second))
                    .Append("</td>");

                foreach (var metric in metrics)
                {
                    var values = row
                        .Where(point =>
                            string.Equals(
                                point.Metric,
                                metric.Metric,
                                StringComparison.Ordinal)
                            && string.Equals(
                                point.Unit,
                                metric.Unit,
                                StringComparison.Ordinal)
                            && string.Equals(
                                ExcelColumnLetters(point.Coordinate),
                                metric.Column,
                                StringComparison.OrdinalIgnoreCase))
                        .GroupBy(point => new
                        {
                            point.DisplayValue,
                            point.Coordinate
                        })
                        .Select(group => group.Key)
                        .ToList();
                    var coordinates = string.Join(
                        ", ",
                        values.Select(value => value.Coordinate)
                            .Where(value => !string.IsNullOrWhiteSpace(value)));
                    rows.Append("<td class='value raw-value'")
                        .Append(string.IsNullOrWhiteSpace(coordinates)
                            ? string.Empty
                            : $" title='원본 셀 {H(coordinates)}'")
                        .Append(">");
                    if (values.Count > 0)
                    {
                        rows.Append(string.Join(
                            " / ",
                            values.Select(value =>
                                $"<strong>{H(value.DisplayValue)}</strong>")));
                    }
                    rows.Append("</td>");
                }
                rows.Append("</tr>");
            }
            rows.Append("</tbody>");

            var location = string.IsNullOrWhiteSpace(source.Key.Sheet)
                ? source.Key.Range
                : $"{source.Key.Sheet}!{source.Key.Range}";
            sourceBlocks.Append("<div class='raw-source'><div class='raw-source-heading'><h4>원본 범위 ")
                .Append(H(location))
                .Append("</h4><span>")
                .Append(sourcePoints.Select(point => point.Row).Distinct().Count())
                .Append("행 · ")
                .Append(metrics.Count)
                .Append("지표</span></div>");
            sourceBlocks.Append(
                    "<div class='table-wrap'><table class='answer-table raw-pivot-table'>")
                .Append(header)
                .Append(rows)
                .Append("</table></div></div>");
        }

        var truncated = study.TryGetProperty(
            "rawDataTruncated",
            out var truncatedElement)
            && truncatedElement.ValueKind == JsonValueKind.True;
        return "<div class='raw-data-block'>"
               + $"<h3>원본 수치 · {points.GetArrayLength():N0}개</h3>"
               + "<p class='note'>왼쪽은 원본 행의 기준·시험 조건, 오른쪽은 측정 지표입니다. 값에 마우스를 올리면 원본 셀 주소를 확인할 수 있습니다.</p>"
               + sourceBlocks
               + (truncated
                   ? "<p class='warning'>표시 상한 이후 값은 왼쪽 원본 근거 범위에서 확인하세요.</p>"
                   : string.Empty)
               + "</div>";
    }

    private static (string First, string Second) SplitRawCondition(
        string condition)
    {
        var parts = RawConditionValues(condition);
        return parts.Count switch
        {
            0 => ("-", string.Empty),
            1 => (parts[0], string.Empty),
            2 => (parts[0], parts[1]),
            _ => (parts[0], $"{parts[1]}: {string.Join(" · ", parts.Skip(2))}")
        };
    }

    private static (string Date, string Type, string Content)
        SplitReviewCondition(
            string condition,
            string fallbackDate,
            string fallbackType = "")
    {
        var parts = RawConditionValues(condition);
        var date = string.IsNullOrWhiteSpace(fallbackDate)
            ? "-"
            : fallbackDate;
        if (parts.Count == 0)
            return (
                date,
                string.IsNullOrWhiteSpace(fallbackType) ? "-" : fallbackType,
                string.Empty);
        if (LooksLikeDate(parts[0]))
            return (
                parts[0],
                parts.Count > 1 ? parts[1] : "-",
                parts.Count > 2
                    ? parts[2]
                    : string.Empty);
        if (!string.IsNullOrWhiteSpace(fallbackType)
            && fallbackType != "-")
            return (date, fallbackType, parts[0]);
        return (
            date,
            parts[0],
            parts.Count > 1
                ? parts[1]
                : string.Empty);
    }

    private static List<string> RawConditionValues(string condition) =>
        condition
            .Split(
                " | ",
                StringSplitOptions.RemoveEmptyEntries
                | StringSplitOptions.TrimEntries)
            .Select(part =>
            {
                var separator = part.IndexOf('=');
                var value = separator >= 0
                    ? part[(separator + 1)..].Trim()
                    : part.Trim();
                return CompactRawCondition(value);
            })
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .ToList();

    private static bool LooksLikeDate(string value) =>
        value.Length >= 8
        && value.Count(character => character is '-' or '/') >= 2
        && char.IsDigit(value[0]);

    private static string CompactRawCondition(string value)
    {
        if (value.Length >= 11
            && value[4] == '-'
            && value[7] == '-'
            && value[10] == 'T')
            return value[..10];
        return value;
    }

    private static string ExcelColumnLetters(string coordinate)
    {
        var result = new StringBuilder();
        foreach (var character in coordinate)
        {
            if (character == '$') continue;
            if (!char.IsLetter(character)) break;
            result.Append(char.ToUpperInvariant(character));
        }
        return result.ToString();
    }

    private static int ExcelColumnNumber(string column)
    {
        if (string.IsNullOrWhiteSpace(column)) return int.MaxValue;
        var result = 0;
        foreach (var character in column)
            result = checked(
                (result * 26) + (char.ToUpperInvariant(character) - 'A' + 1));
        return result;
    }

    private static string RenderIds(JsonElement element, string property)
    {
        if (!element.TryGetProperty(property, out var values)
            || values.ValueKind != JsonValueKind.Array)
            return string.Empty;
        return string.Join(", ", values.EnumerateArray().Select(value => value.GetString()));
    }

    internal static string RenderDetail(EvidenceDetailDocument detail)
    {
        using var document = JsonDocument.Parse(detail.RawJson);
        var root = document.RootElement;
        var preview = root.GetProperty("preview");
        var mergeCount = preview.GetProperty("mergedRanges").GetArrayLength();
        var capturedCount = preview
            .GetProperty("capturedCellCountInRange")
            .GetInt32();
        var grid = RenderEvidenceGrid(preview);
        return Page(
            $"근거 표 {detail.EvidenceId}",
            $"""
             <div class="summary">
               <strong>{H(detail.TrustStatus)}</strong>
               &nbsp;·&nbsp; {H(Path.GetFileName(detail.SourcePath))}
               &nbsp;·&nbsp; {H(detail.Sheet)}!{H(detail.Range)}
             </div>
             <p>캡처 셀 {capturedCount:N0}개 · 교차 병합 범위 {mergeCount:N0}개 · 이미지 분석 안 함</p>
             {grid}
             <p class="note">이 표는 EVD에 고정된 현재 Capture v2 revision만 읽었습니다. 다른 revision으로 대체 조회하지 않습니다.</p>
             """);
    }

    internal static string RenderIngest(
        IngestWorkbookResult result,
        RelatedStudiesDocument related)
    {
        var relatedRows = related.Studies.Count == 0
            ? "<p>현재 canonical DB에서 공통 factor/outcome/context/title 용어가 있는 다른 Study가 없습니다.</p>"
            : "<table><thead><tr><th>DATA ID</th><th>유사도</th><th>Study</th><th>원본</th></tr></thead><tbody>"
              + string.Join(
                  string.Empty,
                  related.Studies.Select(item =>
                      $"<tr><td><code>{H(item.PublicDataId)}</code></td>"
                      + $"<td>{item.SimilarityScore:F3}</td>"
                      + $"<td>{H(item.Title)}</td>"
                      + $"<td>{H(Path.GetFileName(item.SourcePath))}</td></tr>"))
              + "</tbody></table>";
        return Page(
            "신규 Excel 증분 적재 결과",
            $"""
             <div class="summary"><strong>{H(result.Status)}</strong> · 원본 상태 {H(result.WorkbookStatus)}</div>
             <dl>
               <dt>Revision</dt><dd><code>{H(result.RevisionUid)}</code></dd>
               <dt>Study 수</dt><dd>{result.StudyCount:N0}</dd>
               <dt>Journal</dt><dd><code>{H(result.JournalPath)}</code></dd>
             </dl>
             <h2>동일 원본 및 관련 Study</h2>
             <p>동일 SHA-256의 다른 원본: {related.ExactDuplicateSourceCount:N0}건. 아래 순위는 검색 보조용 어휘 유사도이며 관계·인과 근거가 아닙니다.</p>
             {relatedRows}
             <p class="note">AI가 생성한 Study는 자동 승인하지 않습니다. NEEDS_REVIEW 자료는 정량 답변에서 제외됩니다. 이미지 추출·분석은 수행하지 않았습니다.</p>
             """);
    }

    internal static string RenderReviewDetail(ReviewDetailDocument detail)
    {
        var evidenceRows = string.Join(
            string.Empty,
            detail.Evidence.Select(item =>
                $"<tr><td><code>{H(item.EvidenceId)}</code></td>"
                + $"<td>{H(item.Role)}</td><td>{H(item.Sheet)}</td>"
                + $"<td>{H(item.Range)}</td>"
                + $"<td>{H(item.ValueSummary)}</td></tr>"));
        var blockers = string.IsNullOrWhiteSpace(detail.BlockerSummary)
            ? "<p>현재 저장된 판정으로 결정론적 효과 계산 준비가 완료됐습니다.</p>"
            : "<pre>" + H(detail.BlockerSummary) + "</pre>";
        return Page(
            $"사람 검토 {detail.PublicComparisonId}",
            $"""
             <div class="summary">
               <strong>{H(detail.PublicComparisonId)}</strong>
               &nbsp;·&nbsp; {H(detail.PublicDataId)}
               &nbsp;·&nbsp; 현재 승인 준비 {detail.ApprovalReady}
             </div>
             <dl>
               <dt>Study</dt><dd>{H(detail.StudyTitle)}</dd>
               <dt>시험군</dt><dd>{H(detail.ComparedArmLabel)}</dd>
               <dt>대조군</dt><dd>{H(detail.ControlArmLabel)}</dd>
               <dt>원본</dt><dd>{H(detail.SourcePath)}</dd>
               <dt>Matching basis</dt><dd>{H(detail.MatchingBasis)}</dd>
             </dl>
             <h2>현재 차단 사유</h2>
             {blockers}
             <h2>직접 셀 근거와 관측값</h2>
             <table><thead><tr><th>EVD ID</th><th>구분</th><th>시트</th><th>범위</th><th>값</th></tr></thead>
             <tbody>{evidenceRows}</tbody></table>
             <p class="warning">승인은 인과관계 자동 판정이 아닙니다. 동일 대상·기간·측정법인지, 표본 차이와 다른 변경 요인이 없는지 원본 표에서 사람이 확인해야 합니다. 이미지 분석은 포함되지 않습니다.</p>
             """);
    }

    internal static string RenderReviewDecision(
        ReviewDecisionDocument decision) =>
        Page(
            $"사람 검토 결과 {decision.PublicComparisonId}",
            $"""
             <div class="summary">
               <strong>{H(decision.Decision)}</strong>
               &nbsp;·&nbsp; {H(decision.PublicComparisonId)}
               &nbsp;·&nbsp; 집계 가능 {decision.AggregationEligible}
             </div>
             <p>결정론적으로 생성·갱신된 효과 ID: {
                 (decision.EffectPublicIds.Count == 0
                     ? "없음"
                     : string.Join(
                         ", ",
                         decision.EffectPublicIds.Select(id =>
                             $"<code>{H(id)}</code>")))
             }</p>
             <pre>{H(decision.RawJson)}</pre>
             <p class="note">승인 효과는 현재 revision의 VERIFIED 비교·관측 근거에서만 계산됩니다. 이미지 분석은 수행하지 않았습니다.</p>
             """);

    private static string JsonValue(JsonElement element, string property)
    {
        if (!element.TryGetProperty(property, out var value)
            || value.ValueKind is JsonValueKind.Null
                or JsonValueKind.Undefined)
            return string.Empty;
        return value.ValueKind == JsonValueKind.String
            ? value.GetString() ?? string.Empty
            : value.GetRawText();
    }

    private static string RenderEvidenceGrid(JsonElement preview)
    {
        var range = preview.GetProperty("range");
        var start = range.GetProperty("start");
        var end = range.GetProperty("end");
        var startRow = start.GetProperty("row").GetInt32();
        var startColumn = start.GetProperty("column").GetInt32();
        var endRow = end.GetProperty("row").GetInt32();
        var endColumn = end.GetProperty("column").GetInt32();
        var rowCount = endRow - startRow + 1;
        var columnCount = endColumn - startColumn + 1;
        var area = (long)rowCount * columnCount;
        if (area > MaxGridCells)
            return RenderSparseCells(
                preview,
                $"근거 범위가 {rowCount:N0}행 × {columnCount:N0}열"
                + $"({area:N0}셀)로 커서 캡처된 셀만 표시합니다.");

        var cells = preview.GetProperty("cells")
            .EnumerateArray()
            .Select(CellView.FromJson)
            .ToDictionary(
                cell => (cell.Row, cell.Column),
                cell => cell);
        var hiddenRows = preview.GetProperty("rowDimensions")
            .EnumerateArray()
            .Where(item => item.GetProperty("hidden").GetBoolean())
            .Select(item => item.GetProperty("row").GetInt32())
            .ToHashSet();
        var hiddenColumns = new HashSet<int>();
        foreach (var dimension in preview
                     .GetProperty("columnDimensions")
                     .EnumerateArray())
        {
            if (!dimension.GetProperty("hidden").GetBoolean()) continue;
            for (var column = dimension.GetProperty("minColumn").GetInt32();
                 column <= dimension.GetProperty("maxColumn").GetInt32();
                 column++)
                hiddenColumns.Add(column);
        }

        var mergeStarts = new Dictionary<(int Row, int Column), MergeView>();
        var mergeCovered = new HashSet<(int Row, int Column)>();
        foreach (var merge in preview
                     .GetProperty("mergedRanges")
                     .EnumerateArray())
        {
            var minRow = Math.Max(
                startRow,
                merge.GetProperty("minRow").GetInt32());
            var minColumn = Math.Max(
                startColumn,
                merge.GetProperty("minColumn").GetInt32());
            var maxRow = Math.Min(
                endRow,
                merge.GetProperty("maxRow").GetInt32());
            var maxColumn = Math.Min(
                endColumn,
                merge.GetProperty("maxColumn").GetInt32());
            if (minRow > maxRow || minColumn > maxColumn) continue;
            var mergeView = new MergeView(
                minRow,
                minColumn,
                maxRow,
                maxColumn,
                merge.GetProperty("address").GetString()
                    ?? string.Empty,
                merge.TryGetProperty("anchorCell", out var anchorCell)
                && anchorCell.ValueKind == JsonValueKind.Object
                    ? CellView.FromJson(anchorCell)
                    : null);
            mergeStarts[(minRow, minColumn)] = mergeView;
            for (var row = minRow; row <= maxRow; row++)
            for (var column = minColumn; column <= maxColumn; column++)
                if (row != minRow || column != minColumn)
                    mergeCovered.Add((row, column));
        }

        var html = new StringBuilder(
            "<div class='grid-wrap'><table class='excel-grid'><thead><tr>"
            + "<th class='corner'></th>");
        for (var column = startColumn; column <= endColumn; column++)
        {
            html.Append("<th")
                .Append(hiddenColumns.Contains(column)
                    ? " class='hidden-dimension'"
                    : string.Empty)
                .Append('>')
                .Append(H(ColumnLabel(column)))
                .Append("</th>");
        }
        html.Append("</tr></thead><tbody>");
        for (var row = startRow; row <= endRow; row++)
        {
            html.Append("<tr><th")
                .Append(hiddenRows.Contains(row)
                    ? " class='hidden-dimension'"
                    : string.Empty)
                .Append('>')
                .Append(row)
                .Append("</th>");
            for (var column = startColumn;
                 column <= endColumn;
                 column++)
            {
                if (mergeCovered.Contains((row, column))) continue;
                cells.TryGetValue((row, column), out var cell);
                mergeStarts.TryGetValue((row, column), out var merge);
                var displayedCell = cell ?? merge?.ExternalAnchor;
                var css = new List<string>();
                if (hiddenRows.Contains(row)
                    || hiddenColumns.Contains(column))
                    css.Add("hidden-dimension");
                if (merge is not null) css.Add("merged-cell");
                if (!string.IsNullOrWhiteSpace(displayedCell?.Formula))
                    css.Add("formula-cell");
                html.Append("<td");
                if (css.Count > 0)
                    html.Append(" class='")
                        .Append(string.Join(' ', css))
                        .Append('\'');
                if (merge is not null)
                    html.Append(" rowspan='")
                        .Append(merge.MaxRow - merge.MinRow + 1)
                        .Append("' colspan='")
                        .Append(merge.MaxColumn - merge.MinColumn + 1)
                        .Append('\'');
                html.Append(" title='")
                    .Append(H(CellTitle(
                        displayedCell,
                        row,
                        column,
                        merge?.Address)))
                    .Append("'>")
                    .Append(H(displayedCell?.PreferredValue ?? string.Empty))
                    .Append("</td>");
            }
            html.Append("</tr>");
        }
        html.Append("</tbody></table></div>");
        return html.ToString();
    }

    private static string RenderSparseCells(
        JsonElement preview,
        string notice)
    {
        var rows = new StringBuilder();
        foreach (var cell in preview.GetProperty("cells").EnumerateArray())
        {
            rows.Append("<tr><td>")
                .Append(H(JsonValue(cell, "coordinate")))
                .Append("</td><td>")
                .Append(H(CellView.FromJson(cell).PreferredValue))
                .Append("</td><td>")
                .Append(H(JsonValue(cell, "rawValue")))
                .Append("</td><td>")
                .Append(H(JsonValue(cell, "formula")))
                .Append("</td><td>")
                .Append(H(JsonValue(cell, "cachedValue")))
                .Append("</td><td>")
                .Append(H(JsonValue(cell, "numberFormat")))
                .Append("</td><td>")
                .Append(H(JsonValue(cell, "mergeRange")))
                .Append("</td></tr>");
        }
        return $"""
                <p class="warning">{H(notice)}</p>
                <table>
                  <thead><tr><th>셀</th><th>표시값</th><th>원시값</th><th>수식</th><th>캐시값</th><th>서식</th><th>병합</th></tr></thead>
                  <tbody>{rows}</tbody>
                </table>
                """;
    }

    private static string CellTitle(
        CellView? cell,
        int row,
        int column,
        string? mergeAddress)
    {
        var builder = new StringBuilder()
            .Append(ColumnLabel(column))
            .Append(row);
        if (cell is not null)
        {
            if (!string.IsNullOrWhiteSpace(cell.RawValue))
                builder.Append("\n원시값: ").Append(cell.RawValue);
            if (!string.IsNullOrWhiteSpace(cell.Formula))
                builder.Append("\n수식: ").Append(cell.Formula);
            if (!string.IsNullOrWhiteSpace(cell.CachedValue))
                builder.Append("\n캐시값: ").Append(cell.CachedValue);
            if (!string.IsNullOrWhiteSpace(cell.NumberFormat))
                builder.Append("\n표시 형식: ").Append(cell.NumberFormat);
        }
        if (!string.IsNullOrWhiteSpace(mergeAddress))
            builder.Append("\n병합: ").Append(mergeAddress);
        return builder.ToString();
    }

    private static string ColumnLabel(int column)
    {
        var result = string.Empty;
        for (var value = column; value > 0; value = (value - 1) / 26)
            result = (char)('A' + ((value - 1) % 26)) + result;
        return result;
    }

    private static string Page(string title, string body) =>
        $$"""
         <!doctype html>
         <html lang="ko">
         <head>
           <meta charset="utf-8">
           <style>
             html { color-scheme: dark; }
             body { font-family: "Segoe UI", "Malgun Gothic", sans-serif; margin: 0; padding: 26px; color: #d6d6d6; background: #202020; font-size: 15px; }
             h1 { margin-top: 0; font-size: 26px; color: #eeeeee; } h2 { margin-top: 26px; font-size: 20px; color: #eeeeee; }
             a { color: #c4b5fd; }
             .summary { padding: 12px 14px; border: 1px solid #514168; background: #2d2736; border-radius: 6px; }
             pre { white-space: pre-wrap; overflow-wrap: anywhere; padding: 16px; color: #d6d6d6; background: #252525; border: 1px solid #3a3a3a; border-radius: 6px; line-height: 1.55; }
             table { border-collapse: collapse; width: 100%; color: #d6d6d6; background: #252525; font-size: 14px; }
             th, td { box-sizing: border-box; border: 1px solid #3a3a3a; padding: 10px; text-align: left; vertical-align: top; overflow-wrap: anywhere; }
             th { color: #eeeeee; background: #303030; position: sticky; top: 0; }
             tbody tr:nth-child(even) { background: #292929; }
             .grid-wrap { max-width: 100%; overflow: auto; max-height: 70vh; border: 1px solid #3a3a3a; background: #252525; }
             .excel-grid { width: max-content; min-width: 100%; table-layout: fixed; }
             .excel-grid th { min-width: 44px; text-align: center; z-index: 1; }
             .excel-grid td { min-width: 88px; height: 24px; white-space: pre-wrap; }
             .excel-grid .corner { min-width: 44px; left: 0; z-index: 2; }
             .merged-cell { background: #2d2736; text-align: center !important; vertical-align: middle !important; }
             .formula-cell { color: #c4b5fd; }
             .hidden-dimension { background-image: repeating-linear-gradient(135deg, transparent, transparent 4px, #2a2a2a 4px, #2a2a2a 8px); }
              .warning { padding: 10px 12px; color: #e6c980; border: 1px solid #685a36; background: #302c22; border-radius: 6px; }
              section { margin-top: 24px; }
              .status-grid { display: grid; grid-template-columns: repeat(4, minmax(110px, 1fr)); gap: 12px; }
              .status-grid div { display: flex; flex-direction: column; gap: 3px; }
              .status-grid span, .eyebrow { color: #8e8e8e; font-size: 12px; }
              .answer-card { padding: 20px; background: #252525; border: 1px solid #514168; border-left: 5px solid #8b5cf6; border-radius: 8px; }
              .direct-answer { margin: 7px 0 0; font-size: 18px; font-weight: 600; line-height: 1.65; }
              .intent-grid { display: grid; grid-template-columns: 80px 1fr; gap: 8px 12px; padding: 14px; background: #252525; border: 1px solid #3a3a3a; border-radius: 6px; }
              .intent-grid dt { margin: 0; color: #8e8e8e; }
              .intent-grid dd { margin: 0; }
              .chip { display: inline-block; margin: 0 4px 4px 0; padding: 3px 8px; border-radius: 999px; background: #3c2a58; color: #d8c7ff; }
              .more-chip { display: inline-block; margin-left: 3px; color: #a78bfa; font-weight: 700; }
              .finding-list { display: grid; gap: 10px; }
              .finding { padding: 14px 16px; background: #252525; border: 1px solid #3a3a3a; border-radius: 6px; line-height: 1.55; }
              .finding p { margin: 5px 0; color: #aaaaaa; }
              .finding small { color: #8e8e8e; overflow-wrap: anywhere; }
              .table-wrap { max-width: 100%; overflow-x: auto; }
              .answer-table { font-size: 13px; }
              .answer-table .value { white-space: nowrap; font-weight: 700; color: #c4b5fd; }
              .review-table-overview { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 12px; padding: 11px 13px; border: 1px solid #554268; border-radius: 7px; background: #2d2735; }
              .review-table-overview strong { color: #f0e5ff; }
              .review-table-overview span { color: #ae9cbd; font-size: 12px; }
              .review-table-block { margin-top: 16px; }
              .review-matrix-heading { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 8px 12px; padding: 11px 13px; border: 1px solid #4a4055; border-bottom: 0; border-radius: 7px 7px 0 0; background: #302a37; }
              .review-matrix-heading strong { color: #f0e9fa; font-size: 15px; }
              .review-matrix-heading span { color: #b5a3cd; font-size: 12px; text-align: right; }
              .review-matrix-wrap { max-width: 100%; max-height: 58vh; overflow: auto; border: 1px solid #4a4055; border-radius: 0 0 7px 7px; }
              .review-matrix-table { width: max-content; min-width: 100%; font-size: 13px; }
              .review-matrix-table th, .review-matrix-table td { padding: 9px; }
              .review-matrix-table .review-no { min-width: 70px; width: 70px; max-width: 70px; text-align: center; }
              .review-matrix-table .review-date { min-width: 100px; width: 100px; max-width: 100px; text-align: center; }
              .review-matrix-table .review-type { min-width: 150px; width: 150px; max-width: 150px; }
              .review-matrix-table .review-content { min-width: 200px; width: 200px; max-width: 200px; }
              .review-matrix-table .review-metric, .review-matrix-table .review-value { min-width: 110px; width: 110px; text-align: center; vertical-align: middle; }
              .review-matrix-table .review-no, .review-matrix-table .review-date, .review-matrix-table .review-type, .review-matrix-table .review-content { position: sticky; z-index: 2; background: #252525; }
              .review-matrix-table .review-no { left: 0; }
              .review-matrix-table .review-date { left: 70px; }
              .review-matrix-table .review-type { left: 170px; }
              .review-matrix-table .review-content { left: 320px; }
              .review-matrix-table thead .review-no, .review-matrix-table thead .review-date, .review-matrix-table thead .review-type, .review-matrix-table thead .review-content { z-index: 5; background: #343434; }
              .review-matrix-table tbody tr:nth-child(even) .review-no, .review-matrix-table tbody tr:nth-child(even) .review-date, .review-matrix-table tbody tr:nth-child(even) .review-type, .review-matrix-table tbody tr:nth-child(even) .review-content { background: #292929; }
              .review-no strong, .review-date strong, .review-type strong, .review-content strong { display: block; color: #eeeeee; line-height: 1.35; }
              .review-value strong { color: #d7c4f5; font-size: 15px; white-space: nowrap; }
              .value-separator { margin: 0 5px; color: #6f6f6f; }
              .normal-rate-item { display: inline-flex; flex-direction: column; align-items: center; gap: 2px; }
              .normal-rate-item small { color: #a99bb8; font-size: 10px; }
              .comparison-description { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 12px; padding: 11px 13px; border: 1px solid #514168; border-radius: 7px; background: #2d2735; }
              .comparison-description > strong { flex: 0 0 auto; color: #f0e5ff; }
              .comparison-description > div { min-width: 0; }
              .review-set-list { display: grid; gap: 16px; }
              .review-set-card { margin: 0; overflow: hidden; border: 1px solid #514168; border-radius: 8px; background: #252525; }
              .review-set-heading { border-bottom: 1px solid #5e4778; background: #292332; }
              .review-set-heading a { display: grid; grid-template-columns: auto minmax(0, 1fr) auto; align-items: center; gap: 12px; padding: 10px 12px; color: inherit; text-decoration: none; }
              .review-set-heading a:hover { background: #342942; }
              .review-set-heading strong { color: #dfcaff; font-size: 13px; white-space: nowrap; }
              .review-set-heading span { min-width: 0; overflow: hidden; color: #eeeeee; font-weight: 700; text-overflow: ellipsis; white-space: nowrap; }
              .review-set-heading small { color: #b9a8cc; font-size: 11px; white-space: nowrap; }
              .review-set-table-wrap { max-width: 100%; overflow-x: auto; }
              .review-set-table { width: max-content; min-width: 100%; font-size: 13px; }
              .review-set-table th, .review-set-table td { padding: 8px 9px; text-align: center; vertical-align: middle; }
              .review-set-table .review-cohort { position: sticky; left: 0; z-index: 2; min-width: 76px; width: 76px; max-width: 76px; }
              .review-set-table .review-date { position: sticky; left: 76px; z-index: 2; min-width: 92px; width: 92px; max-width: 92px; }
              .review-set-table .review-comparison { position: sticky; left: 168px; z-index: 2; min-width: 240px; width: 240px; max-width: 240px; text-align: left; }
              .review-set-table thead .review-cohort, .review-set-table thead .review-date, .review-set-table thead .review-comparison { z-index: 5; background: #343434; }
              .review-set-table .review-metric-header, .review-set-table .review-metric-cell { min-width: 98px; text-align: center; white-space: nowrap; }
              .review-set-table .review-date strong, .review-set-table .review-comparison strong { display: block; color: #eeeeee; line-height: 1.35; }
              .review-set-table .review-comparison small { display: block; margin-top: 3px; color: #b8b8b8; line-height: 1.35; }
              .cohort-badge { display: inline-flex; align-items: center; justify-content: center; min-width: 50px; padding: 4px 7px; border: 1px solid transparent; border-radius: 999px; font-size: 11px; font-weight: 800; white-space: nowrap; }
              .cohort-badge.control { border-color: #3d8b78; color: #b8f3df; background: #223c36; }
              .cohort-badge.comparison { border-color: #7251a0; color: #e0ccff; background: #38274f; }
              .review-control-row td { background: #232a28; }
              .review-control-row .review-cohort, .review-control-row .review-date, .review-control-row .review-comparison { background: #232a28; }
              .review-control-row:hover td { background: #293431; }
              .review-control-row:hover .review-cohort, .review-control-row:hover .review-date, .review-control-row:hover .review-comparison { background: #293431; }
              .review-comparison-row .review-cohort, .review-comparison-row .review-date, .review-comparison-row .review-comparison { background: #252525; }
              .review-comparison-row:hover td { background: #29252e; }
              .review-comparison-row:hover .review-cohort, .review-comparison-row:hover .review-date, .review-comparison-row:hover .review-comparison { background: #29252e; }
              .review-comparison-link { display: block; color: inherit; text-decoration: none; cursor: pointer; }
              .review-comparison-link:hover strong, .review-comparison-link:hover small { color: #d8c7ff; }
              .review-value-link { display: block; margin: -8px -9px; padding: 8px 9px; color: #f1e9ff; text-decoration: none; cursor: pointer; }
              .review-value-link:hover { color: #ffffff; background: #51367a; }
              .review-value-link strong { font-size: 14px; }
              .empty-value { color: #666666; }
              .study-list { display: grid; gap: 12px; }
              .study-card { overflow: hidden; border: 1px solid #454545; border-radius: 8px; background: #252525; }
              .study-card[open] { border-color: #66517f; }
              .study-card > summary { display: grid; grid-template-columns: 38px minmax(0, 1fr) auto auto; align-items: center; gap: 12px; padding: 14px 16px; cursor: pointer; list-style: none; background: #292929; }
              .study-card > summary::-webkit-details-marker { display: none; }
              .study-card > summary::after { content: "펼치기"; margin-left: 10px; color: #a78bfa; font-size: 12px; }
              .study-card[open] > summary::after { content: "접기"; }
              .study-rank { display: grid; place-items: center; width: 32px; height: 32px; border-radius: 50%; color: #efe8ff; background: #51367a; font-weight: 800; }
              .study-title { min-width: 0; }
              .study-title strong { display: block; color: #f2f2f2; font-size: 15px; overflow-wrap: anywhere; }
              .study-title small { display: block; margin-top: 4px; color: #9b9b9b; }
              .study-count { color: #c9b6e8; font-size: 12px; white-space: nowrap; }
              .study-body { padding: 18px; }
              .selection-reason { padding: 12px 14px; border-left: 4px solid #8b5cf6; background: #2c2832; }
              .selection-reason > span, .study-meta-grid > div > span { display: block; margin-bottom: 6px; color: #9b9b9b; font-size: 12px; font-weight: 700; }
              .selection-reason p { margin: 0 0 8px; line-height: 1.55; }
              .study-meta-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; margin-top: 12px; }
              .study-meta-grid > div { padding: 12px; border: 1px solid #3a3a3a; border-radius: 6px; background: #232323; }
              .raw-data-block { margin-top: 22px; }
              .raw-data-block h3 { margin-bottom: 4px; color: #eeeeee; font-size: 18px; }
              .raw-source { margin-top: 16px; overflow: hidden; border: 1px solid #414141; border-radius: 7px; background: #232323; }
              .raw-source-heading { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 10px 12px; background: #303030; }
              .raw-source-heading h4 { margin: 0; color: #eeeeee; font-size: 14px; }
              .raw-source-heading span { color: #a6a6a6; font-size: 12px; white-space: nowrap; }
              .raw-source .table-wrap { max-height: 520px; overflow: auto; }
              .raw-pivot-table { width: max-content; min-width: 100%; font-size: 14px; }
              .raw-pivot-table .raw-condition { min-width: 130px; width: 130px; max-width: 130px; }
              .raw-pivot-table .raw-condition:nth-child(2) { min-width: 210px; width: 210px; max-width: 210px; }
              .raw-pivot-table .raw-metric { min-width: 115px; width: 115px; text-align: center; white-space: normal; }
              .raw-pivot-table .raw-value { min-width: 115px; width: 115px; text-align: center; vertical-align: middle; font-size: 15px; }
              .raw-pivot-table small { display: block; margin-top: 3px; color: #8e8e8e; font-weight: 400; }
              .raw-pivot-table th:nth-child(1), .raw-pivot-table td:nth-child(1) { position: sticky; left: 0; z-index: 2; }
              .raw-pivot-table th:nth-child(2), .raw-pivot-table td:nth-child(2) { position: sticky; left: 130px; z-index: 2; }
              .raw-pivot-table th:nth-child(1), .raw-pivot-table th:nth-child(2) { z-index: 4; background: #343434; }
              .raw-pivot-table tbody td:nth-child(1), .raw-pivot-table tbody td:nth-child(2) { background: #252525; }
              .raw-pivot-table tbody tr:nth-child(even) td:nth-child(1), .raw-pivot-table tbody tr:nth-child(even) td:nth-child(2) { background: #292929; }
              .evidence-table { font-size: 12px; table-layout: fixed; }
              .evidence-table th:nth-child(1) { width: 48%; }
              .evidence-table th:nth-child(2) { width: 22%; }
              .limitations { margin: 0; padding: 14px 18px 14px 36px; background: #302c22; border: 1px solid #685a36; border-radius: 6px; line-height: 1.55; }
              .empty, .muted { color: #8e8e8e; }
             code { color: #c4b5fd; } .note { color: #aaaaaa; line-height: 1.5; }
             dt { font-weight: 600; margin-top: 10px; } dd { margin: 3px 0 0 0; }
             @media (max-width: 900px) {
               body { padding: 16px; }
               .status-grid { grid-template-columns: repeat(2, minmax(110px, 1fr)); }
               .study-card > summary { grid-template-columns: 34px minmax(0, 1fr); }
               .study-count { grid-column: 2; }
               .study-card > summary::after { grid-column: 2; margin-left: 0; }
               .study-meta-grid { grid-template-columns: 1fr; }
               .review-matrix-table .review-no { min-width: 60px; width: 60px; max-width: 60px; }
               .review-matrix-table .review-date { min-width: 90px; width: 90px; max-width: 90px; left: 60px; }
               .review-matrix-table .review-type { min-width: 130px; width: 130px; max-width: 130px; left: 150px; }
               .review-matrix-table .review-content { min-width: 170px; width: 170px; max-width: 170px; left: 280px; }
               .raw-pivot-table .raw-condition { min-width: 115px; width: 115px; max-width: 115px; }
               .raw-pivot-table .raw-condition:nth-child(2) { min-width: 170px; width: 170px; max-width: 170px; }
               .raw-pivot-table th:nth-child(2), .raw-pivot-table td:nth-child(2) { left: 115px; }
             }
           </style>
         </head>
         <body><h1>{{H(title)}}</h1>{{body}}</body>
         </html>
         """;

    private static string H(object? value) =>
        WebUtility.HtmlEncode(Convert.ToString(value) ?? string.Empty);

    private sealed record RawDataPointView(
        int Index,
        string EvidenceId,
        string TableId,
        string Sheet,
        string Range,
        int Row,
        string Condition,
        string Metric,
        string Unit,
        string DisplayValue,
        string Coordinate)
    {
        internal static RawDataPointView FromJson(
            JsonElement point,
            int index)
        {
            var coordinate = JsonValue(point, "coordinate");
            var row = point.TryGetProperty("row", out var rowElement)
                      && rowElement.ValueKind == JsonValueKind.Number
                      && rowElement.TryGetInt32(out var parsedRow)
                ? parsedRow
                : 0;
            return new RawDataPointView(
                index,
                JsonValue(point, "evidenceId"),
                JsonValue(point, "tableId"),
                JsonValue(point, "sheet"),
                JsonValue(point, "range"),
                row,
                JsonValue(point, "condition"),
                JsonValue(point, "metric"),
                JsonValue(point, "unit"),
                JsonValue(point, "displayValue"),
                coordinate);
        }
    }

    private sealed record MetricColumn(
        string Metric,
        string Unit,
        string Column,
        int ColumnNumber,
        int FirstIndex);

    private sealed record ReviewPoint(
        int StudyIndex,
        int Rank,
        string FileName,
        string SourcePath,
        string Date,
        string StudyGroup,
        string ReviewTitle,
        string MetricKey,
        RawDataPointView RawPoint);

    private sealed record ReviewSource(
        string FamilyKey,
        string FamilyTitle,
        IReadOnlyList<ReviewPoint> Points)
    {
        internal ReviewPoint First => Points[0];
    }

    private sealed record ReviewMetricColumn(
        string Key,
        string Label);

    private sealed record CellView(
        int Row,
        int Column,
        string DisplayValue,
        string CachedValue,
        string RawValue,
        string Formula,
        string NumberFormat)
    {
        internal string PreferredValue =>
            !string.IsNullOrWhiteSpace(DisplayValue)
                ? DisplayValue
                : !string.IsNullOrWhiteSpace(CachedValue)
                    ? CachedValue
                    : RawValue;

        internal static CellView FromJson(JsonElement cell) => new(
            cell.GetProperty("row").GetInt32(),
            cell.GetProperty("column").GetInt32(),
            JsonValue(cell, "displayValue"),
            JsonValue(cell, "cachedValue"),
            JsonValue(cell, "rawValue"),
            JsonValue(cell, "formula"),
            JsonValue(cell, "numberFormat"));
    }

    private sealed record MergeView(
        int MinRow,
        int MinColumn,
        int MaxRow,
        int MaxColumn,
        string Address,
        CellView? ExternalAnchor);
}
