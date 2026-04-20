namespace JinoSupporter.Web.Services;

/// <summary>
/// Detects likely AI-extraction errors in NormalizedMeasurement rows by running
/// simple per-row-group sanity checks. Used as a soft warning — results are
/// advisory, not blocking.
/// </summary>
public static class MeasurementValidator
{
    public sealed record Issue(string Variable, string VariableDetail, string Kind, string Message);

    public static List<Issue> FindIssues(IReadOnlyList<NormalizedMeasurement> rows)
    {
        var issues = new List<Issue>();
        if (rows.Count == 0) return issues;

        // Pre-compute: does each TABLE (variableDetail) contain at least one row with
        // a non-empty defectType? If not, the source report had no per-defect columns
        // for that table — every row being `defectType=""` is legitimate, not a bug.
        //
        // This correctly distinguishes:
        //   (a) RA-line / OQC sub-tables with only Input/OK/Total NG    → skip all
        //   (b) Standard tables where AI dropped some rows' defects     → still flag
        // without relying on fragile `ngTotal <= 5` heuristics.
        var tableHasDefects = rows
            .GroupBy(r => r.VariableDetail ?? "")
            .ToDictionary(
                g => g.Key,
                g => g.Any(r => !string.IsNullOrWhiteSpace(r.DefectType)));

        var groups = rows
            .GroupBy(r => (r.Variable, r.VariableDetail, r.VariableGroup, r.Line, r.CheckType,
                           r.InputQty, r.OkQty, r.NgTotal));

        foreach (var g in groups)
        {
            var list = g.ToList();
            int ngTotal = g.Key.NgTotal;

            // Direct alignment-error marker from prompt
            if (list.Any(m => m.DefectType == "__ALIGN_ERROR__"))
            {
                issues.Add(new Issue(g.Key.Variable, g.Key.VariableDetail, "align",
                    "컬럼 정렬 실패 — 값 일부가 비어있거나 섹션 경계를 잘못 읽었을 가능성."));
                continue;
            }

            // Skip: entire table (variableDetail) is aggregate-only (no defect columns
            // in source). The AI correctly emitted single empty rows throughout — not a bug.
            if (!tableHasDefects.GetValueOrDefault(g.Key.VariableDetail ?? "", false))
                continue;

            // Skip: criterion-level evaluation (Wire Moving / Frame Deform / Hearing OQC
            // etc. — overall OK doesn't reduce when a criterion fails). Recognised when
            // any defectType contains clear criterion suffix words.
            if (list.Any(m => IsCriterionLabel(m.DefectType)))
                continue;

            // Skip: picture-sample catalog row with Input=OK=0 (no actual count data).
            if (g.Key.InputQty == 0 && g.Key.OkQty == 0)
                continue;

            // Sum only positive, named-defect rows (exclude synthetic "total" entries with empty defectType).
            int sumDefects = list
                .Where(m => !string.IsNullOrWhiteSpace(m.DefectType))
                .Sum(m => m.DefectCount);

            if (ngTotal > 0 && sumDefects == 0)
            {
                issues.Add(new Issue(g.Key.Variable, g.Key.VariableDetail, "missing",
                    $"NG Total={ngTotal} 인데 defect 행이 하나도 없음 — 항목 누락 의심."));
            }
            else if (ngTotal > 0 && sumDefects < ngTotal)
            {
                issues.Add(new Issue(g.Key.Variable, g.Key.VariableDetail, "undercount",
                    $"defect count 합={sumDefects} < NG Total={ngTotal} — 하나 이상의 결함 행이 누락됐을 수 있음."));
            }
            else if (ngTotal == 0 && sumDefects > 0)
            {
                issues.Add(new Issue(g.Key.Variable, g.Key.VariableDetail, "contradict",
                    $"NG Total=0 인데 defect count 합={sumDefects} — 모순."));
            }
        }

        return issues;
    }

    private static bool IsCriterionLabel(string defectType)
    {
        if (string.IsNullOrWhiteSpace(defectType)) return false;
        // Labels like "Wire Moving NG", "Frame Deform OK", "NG Hearing OQC" describe
        // per-criterion status where the unit may still be overall-OK — checksums
        // against overall ngTotal don't apply.
        return defectType.Contains("Moving",   StringComparison.OrdinalIgnoreCase) ||
               defectType.Contains("Deform",   StringComparison.OrdinalIgnoreCase) ||
               defectType.Contains(" OK",      StringComparison.OrdinalIgnoreCase) ||
               defectType.Contains("OQC",      StringComparison.OrdinalIgnoreCase);
    }
}
