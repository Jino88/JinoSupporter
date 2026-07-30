namespace InferenceDataAIService.Wpf;

// One legacy-browser-compatible dashboard shell. Profiles supply only facts that
// can be traced to captured cells; this layer never invents a business decision.
internal static class DashboardHtmlRenderer
{
    private const string Css = """
html,body{margin:0;padding:0;background:#f3f6fa;color:#172b4d;font-family:Segoe UI,Malgun Gothic,Arial,sans-serif;font-size:13px;line-height:1.45}body{overflow-x:hidden}header{padding:22px 28px 18px;background:#fff;border-bottom:1px solid #d5dfeb}h1{margin:0 0 6px;font-size:25px;word-wrap:break-word}.meta{color:#52606d;font-size:12px;word-wrap:break-word}main{width:94%;max-width:1480px;margin:0 auto;padding:18px 0 32px}.panel{margin:0 0 16px;padding:14px;background:#fff;border:1px solid #d5dfeb}h2{margin:0 0 10px;color:#123b68;font-size:18px}table{width:100%;border-collapse:collapse;background:#fff}th,td{padding:9px 10px;border:1px solid #d5dfeb;text-align:left;vertical-align:top;word-wrap:break-word}thead th{background:#eaf0f7;text-align:center;color:#344d6f}tbody th{background:#f7f9fc;min-width:135px}.table-wrap{width:100%;overflow-x:auto;border:1px solid #d5dfeb}.table-wrap table{min-width:760px;border:0}.table-wrap th,.table-wrap td{border-top:0;border-left:0}.num,.rate{text-align:right;font-weight:bold;white-space:nowrap}.evidence{font-family:Consolas,monospace;font-size:11px;color:#52606d}.component{margin:-3px 0 10px;color:#52606d;font-size:12px}.source-grid .grid-header{background:#eaf0f7;color:#344d6f;text-align:center;font-weight:700}.source-grid .grid-label{background:#f7f9fc;font-weight:600}.source-grid tr:has(.grid-header){background:#eef4fa}.badge{display:inline-block;padding:3px 8px;font-size:11px;font-weight:bold}.good{background:#e8f7ed;color:#067647}.bad{background:#fdecec;color:#b42318}.review{background:#fff4dd;color:#a15c00}.neutral{background:#eaf2ff;color:#175cd3}.empty{padding:12px;background:#fff8e7;color:#8c5d00}.requirements{padding-left:20px;margin:0}footer{margin-top:18px;padding:10px 12px;background:#eaf0f7;color:#52606d;font-size:11px;word-wrap:break-word}
""";

    internal static string Render(DashboardDocument document)
    {
        string Row(DashboardSummaryItem item) => $"<tr><th>{Html(item.Label)}</th><td>{Html(item.Content)}</td><td>{Badge(item.State)}</td><td class='evidence'>{Html(item.Evidence)}</td></tr>";
        var summary = string.Concat(document.Summary.Select(Row));
        var sections = string.Concat(document.Sections.Select(section => $"<section class='panel'><h2>{Html(section.Title)}</h2><p class='component'>{Html(section.Component)} · 근거: {Html(section.Evidence)}</p>{section.Html}</section>"));
        var findings = document.Findings.Count == 0
            ? "<p class='empty'>표시된 값만 보존했습니다. 비교 또는 결론을 만들 근거가 충분하지 않습니다.</p>"
            : "<table><thead><tr><th>항목</th><th>근거 기반 결과</th><th>상태</th><th>근거</th></tr></thead><tbody>" + string.Concat(document.Findings.Select(f => $"<tr><th>{Html(f.Label)}</th><td>{Html(f.Content)}</td><td>{Badge(f.State)}</td><td class='evidence'>{Html(f.Evidence)}</td></tr>")) + "</tbody></table>";
        var requirements = string.Concat(document.ReviewRequirements.Select(value => $"<li>{Html(value)}</li>"));
        return "<!doctype html><html lang='ko'><head><meta charset='utf-8'><meta http-equiv='X-UA-Compatible' content='IE=edge'><title>"
            + Html(document.Title) + "</title><style>" + Css + "</style></head><body><header><h1>" + Html(document.Title)
            + "</h1><div class='meta'>원본: " + Html(document.SourceTrace.RelativePath) + " · 배치: " + Html(document.SourceTrace.BatchId) + " · 렌더러: " + Html(document.SourceTrace.RendererVersion)
            + "</div></header><main><section class='panel'><h2>분석 요약</h2><div class='table-wrap'><table><thead><tr><th>항목</th><th>내용</th><th>상태</th><th>근거</th></tr></thead><tbody>" + summary
            + "</tbody></table></div></section>" + sections + "<section class='panel'><h2>핵심 확인 사항</h2><div class='table-wrap'>" + findings
            + "</div></section><section class='panel'><h2>최종 상태</h2><p>" + Badge(document.Assessment.State) + " " + Html(document.Assessment.Content)
            + "</p></section><section class='panel'><h2>검토 조건</h2><ul class='requirements'>" + requirements
            + "</ul></section><footer>추적 정보 · 레이아웃 서명: " + Html(document.SourceTrace.LayoutSignature) + " · 캡처: " + Html(document.SourceTrace.CaptureVersion) + " · 계획: " + Html(document.SourceTrace.PlanVersion)
            + "</footer></main></body></html>";
    }

    private static string Badge(string state)
    {
        var css = state switch
        {
            "SUPPORTED" or "OK" or "READY" or "CAPTURED_COMPONENTS" or "COMPONENTS_RENDERED" => "good",
            "NO_DECISION" or "NG" or "CONTRACT_MISMATCH" => "bad",
            "REVIEW_REQUIRED" or "PARTIAL" => "review",
            _ => "neutral"
        };
        return $"<span class='badge {css}'>{Html(state)}</span>";
    }
    internal static string Html(string? value) => System.Net.WebUtility.HtmlEncode(value ?? string.Empty);
}

internal sealed record DashboardDocument(string Title, DashboardSourceTrace SourceTrace, IReadOnlyList<DashboardSummaryItem> Summary, IReadOnlyList<DashboardSection> Sections, IReadOnlyList<DashboardFinding> Findings, DashboardAssessment Assessment, IReadOnlyList<string> ReviewRequirements);
internal sealed record DashboardSourceTrace(string RelativePath, string BatchId, string LayoutSignature, string CaptureVersion, string PlanVersion, string RendererVersion);
internal sealed record DashboardSummaryItem(string Label, string Content, string State, string Evidence);
internal sealed record DashboardSection(string Title, string Component, string Evidence, string Html);
internal sealed record DashboardFinding(string Label, string Content, string State, string Evidence);
internal sealed record DashboardAssessment(string State, string Content);
