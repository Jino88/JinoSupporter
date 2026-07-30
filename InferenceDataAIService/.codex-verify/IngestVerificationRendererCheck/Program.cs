using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

if (args.Length != 3)
    throw new ArgumentException(
        "Expected assembly, manifest, and journal paths.");

var assembly = Assembly.LoadFrom(Path.GetFullPath(args[0]));
const BindingFlags instanceFlags =
    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
var resultType = assembly.GetType(
    "InferenceDataAIService.Wpf.IngestWorkbookResult",
    throwOnError: true)!;
var sourcePath =
    @"D:\000. MyWorks\115. TIU L5S3-01 R Report compare ass'y VP+Fr by MC and By hand -  2026.07.22.xlsx";
var clientType = assembly.GetType(
    "InferenceDataAIService.Wpf.CanonicalEvidenceClient",
    throwOnError: true)!;
var client = clientType.GetConstructors(instanceFlags)[0].Invoke(
    [Directory.GetCurrentDirectory()]);
var loadLatest = clientType.GetMethod(
    "LoadLatestIngestResult",
    BindingFlags.Instance | BindingFlags.NonPublic)
    ?? throw new MissingMethodException(
        clientType.FullName,
        "LoadLatestIngestResult");
var result = loadLatest.Invoke(client, [sourcePath])
    ?? throw new InvalidOperationException(
        "The latest completed ingest journal was not restored.");

var rowType = assembly.GetType(
    "InferenceDataAIService.Wpf.RelatedStudyRow",
    throwOnError: true)!;
var listType = typeof(List<>).MakeGenericType(rowType);
var rows = Activator.CreateInstance(listType)!;
var relatedType = assembly.GetType(
    "InferenceDataAIService.Wpf.RelatedStudiesDocument",
    throwOnError: true)!;
var related = relatedType.GetConstructors(instanceFlags)[0].Invoke(
[
    "capture_revision_2c8c2be088c44b3b5aebb8dd",
    0,
    rows,
]);
var rendererType = assembly.GetType(
    "InferenceDataAIService.Wpf.IngestVerificationHtmlRenderer",
    throwOnError: true)!;
var render = rendererType.GetMethod(
    "Render",
    BindingFlags.Static | BindingFlags.NonPublic)
    ?? throw new MissingMethodException(rendererType.FullName, "Render");
var html = (string)render.Invoke(null, [result, related])!;

var checks = new Dictionary<string, bool>
{
    ["journalRestore"] = (string?)resultType
        .GetProperty("ManifestPath")!
        .GetValue(result) == Path.GetFullPath(args[1]),
    ["study1"] = html.Contains(
        "기능 검사 결과",
        StringComparison.Ordinal),
    ["study2"] = html.Contains(
        "최종 외관 검사 결과",
        StringComparison.Ordinal),
    ["value255"] = html.Contains(">255<", StringComparison.Ordinal),
    ["value320"] = html.Contains(">320<", StringComparison.Ordinal),
    ["excelLinks"] = html.Contains(
        "inference-excel://open/",
        StringComparison.Ordinal),
    ["needsReview"] = html.Contains(
        "NEEDS_REVIEW",
        StringComparison.Ordinal),
    ["comparisonTables"] =
        System.Text.RegularExpressions.Regex.Matches(
            html,
            "<table class='comparison-table'>").Count == 2,
    ["cohortLabels"] = html.Contains(
        ">기준군<",
        StringComparison.Ordinal)
        && html.Contains(
            ">비교군<",
            StringComparison.Ordinal),
    ["compactMetrics"] = html.Contains(
        ">검사 수량<",
        StringComparison.Ordinal)
        && html.Contains(
            ">전체 NG<",
            StringComparison.Ordinal)
        && html.Contains(
            "rate-value",
            StringComparison.Ordinal),
    ["emptyArmHidden"] = !html.Contains(
        "After 1 day check again",
        StringComparison.Ordinal),
    ["reviewCells"] = html.Contains(
        "metric-cell review-issue",
        StringComparison.Ordinal)
        && html.Contains(
            "review-marker",
            StringComparison.Ordinal),
    ["excelHover"] = html.Contains(
        "원본 Excel: Test!G19",
        StringComparison.Ordinal)
        && html.Contains(
            "원본 표시값: 255",
            StringComparison.Ordinal)
        && html.Contains(
            "클릭하여 원본 셀 열기",
            StringComparison.Ordinal),
    ["comparisonCautionsRemoved"] = !html.Contains(
        "비교한 조립 조건의 검사 수량",
        StringComparison.Ordinal)
        && !html.Contains(
            "무작위 배정",
            StringComparison.Ordinal)
        && !html.Contains(
            "어느 조립 조건도 대조군",
            StringComparison.Ordinal),
    ["functionInputClean"] =
        System.Text.RegularExpressions.Regex.IsMatch(
            html,
            "<td class='metric-cell'><a href='[^']*range=G19'[^>]*><strong>255</strong>",
            System.Text.RegularExpressions.RegexOptions.Singleline)
        && System.Text.RegularExpressions.Regex.IsMatch(
            html,
            "<td class='metric-cell'><a href='[^']*range=G21'[^>]*><strong>320</strong>",
            System.Text.RegularExpressions.RegexOptions.Singleline),
    ["reviewResidual"] = html.Contains(
            "Input 311은 OK 297 및 전체 NG 7의 합계와 일치하지 않습니다",
            StringComparison.Ordinal),
    ["reviewMissingCells"] = html.Contains(
            "외관 NG 항목별 건수 셀 I40:L40",
            StringComparison.Ordinal),
    ["clutterRemoved"] = new[]
    {
        "Workbook 분석 요약",
        "분석 ID",
        "Revision",
        "Manifest",
        "Journal",
        "중복·관련 Study",
        "검토 필요사항",
        "Study 직접 근거",
        "시험군·비교군·조건",
        "측정값·관측값",
        "비교 정의",
    }.All(value => !html.Contains(value, StringComparison.Ordinal)),
};
var linkCount = System.Text.RegularExpressions.Regex.Matches(
    html,
    "inference-excel://open/").Count;
Console.WriteLine(
    $"htmlLength={html.Length}; evidenceLinks={linkCount}; "
    + string.Join(
        "; ",
        checks.Select(pair => $"{pair.Key}={pair.Value}")));
if (!checks["reviewResidual"])
{
    var diagnosticIndex = html.IndexOf(
        "Input 311",
        StringComparison.Ordinal);
    Console.WriteLine(
        diagnosticIndex < 0
            ? "reviewResidualDiagnostic=Input 311 not found"
            : "reviewResidualDiagnostic="
              + html.Substring(
                  diagnosticIndex,
                  Math.Min(180, html.Length - diagnosticIndex)));
}
if (checks.Values.Any(value => !value) || linkCount == 0)
    Environment.ExitCode = 1;
