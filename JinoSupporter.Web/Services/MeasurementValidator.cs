namespace JinoSupporter.Web.Services;

/// <summary>
/// Detects likely AI-extraction errors in NormalizedMeasurement rows by running
/// simple per-row-group sanity checks. Used as a soft warning — results are
/// advisory, not blocking.
/// </summary>
public static class MeasurementValidator
{
    public sealed record Issue(string Variable, string VariableDetail, string Kind, string Message);

    public static List<Issue> FindIssues(
        IReadOnlyList<NormalizedMeasurement> rows,
        string? excelPasteText = null)
    {
        var issues = new List<Issue>();
        if (rows.Count == 0) return issues;

        // Group by the logical unit-row (Variable, Group, Line, CheckType, Input, OK, NG).
        // VariableDetail is intentionally excluded: the AI sometimes splits one logical
        // row into multiple VariableDetail buckets (e.g. "Normal – aggregate" vs
        // "Normal – per defect") and each bucket holds only a partial defect list.
        // Summing defect counts across those buckets recovers the true per-row total.
        var groups = rows
            .GroupBy(r => (r.Variable, r.VariableGroup, r.Line, r.CheckType,
                           r.InputQty, r.OkQty, r.NgTotal));

        foreach (var g in groups)
        {
            var list = g.ToList();
            int ngTotal = g.Key.NgTotal;

            // Representative detail for Issue messages — distinct VariableDetails joined.
            string variableDetail = string.Join(
                " / ",
                list.Select(m => m.VariableDetail ?? "").Where(s => s.Length > 0).Distinct());

            // Direct alignment-error marker from prompt
            if (list.Any(m => m.DefectType == "__ALIGN_ERROR__"))
            {
                issues.Add(new Issue(g.Key.Variable, variableDetail, "align",
                    "Column alignment failed — some values may be empty or a section boundary was read incorrectly."));
                continue;
            }

            // Skip: this logical row has no per-defect columns in the source report
            // (aggregate-only). Every row being `defectType=""` is legitimate.
            if (!list.Any(m => !string.IsNullOrWhiteSpace(m.DefectType)))
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

            // Skip overcount silently: happens when two legitimately distinct sub-tables
            // happen to share (Variable, Group, Line, CheckType, Input, NG) and each
            // carries a full defect breakdown. Rare; not an undercount.
            if (ngTotal > 0 && sumDefects > ngTotal)
                continue;

            if (ngTotal > 0 && sumDefects == 0)
            {
                issues.Add(new Issue(g.Key.Variable, variableDetail, "missing",
                    $"NG Total={ngTotal} but no defect rows found — suspected missing entries."));
            }
            else if (ngTotal > 0 && sumDefects < ngTotal)
            {
                issues.Add(new Issue(g.Key.Variable, variableDetail, "undercount",
                    $"Sum of defect counts={sumDefects} < NG Total={ngTotal} — one or more defect rows may be missing."));
            }
            else if (ngTotal == 0 && sumDefects > 0)
            {
                issues.Add(new Issue(g.Key.Variable, variableDetail, "contradict",
                    $"NG Total=0 but sum of defect counts={sumDefects} — contradiction."));
            }
        }

        // Cross-check stored InputQty values against the human-pasted Excel ground
        // truth (when available). If the AI extracted an Input value that does not
        // appear verbatim in the source workbook, it likely misread the cell.
        if (!string.IsNullOrWhiteSpace(excelPasteText))
        {
            // Commas in thousand separators ("3,600") would break a digit-run match —
            // strip them before scanning.
            string haystack = excelPasteText.Replace(",", "");

            var seen = new HashSet<(string Variable, string Detail, int Input)>();
            foreach (var r in rows)
            {
                if (r.InputQty < 10) continue;                 // too-small ≈ noise
                if (r.DefectType == "__ALIGN_ERROR__") continue;

                var key = (r.Variable ?? "", r.VariableDetail ?? "", r.InputQty);
                if (!seen.Add(key)) continue;

                if (!ContainsWholeNumber(haystack, r.InputQty))
                {
                    issues.Add(new Issue(
                        r.Variable ?? "", r.VariableDetail ?? "",
                        "excel-mismatch",
                        $"InputQty={r.InputQty} not found in Excel paste — value may be mis-read."));
                }
            }
        }

        return issues;
    }

    /// <summary>
    /// True when the decimal representation of <paramref name="value"/> appears in
    /// <paramref name="text"/> as a standalone number (no adjacent digit on either side).
    /// Example: <c>ContainsWholeNumber("total 100", 100) = true</c>,
    /// <c>ContainsWholeNumber("item 1000", 100) = false</c>.
    /// </summary>
    private static bool ContainsWholeNumber(string text, int value)
    {
        string needle = value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        int idx = 0;
        while ((idx = text.IndexOf(needle, idx, StringComparison.Ordinal)) >= 0)
        {
            bool leftOk  = idx == 0                    || !char.IsDigit(text[idx - 1]);
            bool rightOk = idx + needle.Length == text.Length
                           || !char.IsDigit(text[idx + needle.Length]);
            if (leftOk && rightOk) return true;
            idx++;
        }
        return false;
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
