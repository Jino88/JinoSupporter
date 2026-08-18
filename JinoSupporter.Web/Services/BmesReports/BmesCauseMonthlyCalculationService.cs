using System.Globalization;
using JinoSupporter.Web.Services.BmesReports.Contracts;

namespace JinoSupporter.Web.Services.BmesReports;

public sealed class BmesCauseMonthlyCalculationService
{
    public CauseMonthlyTabDto Calculate(BmesReportRequest request, BmesDailyCalculationSnapshot daily)
    {
        NgRateReportService.NgRateReport report = daily.Hierarchy.ByMid
            ?? throw new InvalidOperationException("Daily model hierarchy is required for cause-monthly.");
        IReadOnlyList<ReportPeriodDto> periods = BmesReportProjection.FromNgRate(report)
            .Where(period => period.Kind == "month")
            .ToArray();
        var rows = new List<CauseRowDto>(Catalog.Count);

        for (int index = 0; index < Catalog.Count; index++)
        {
            CauseCatalogRow seed = Catalog[index];
            Dictionary<string, double?> ppm = periods.ToDictionary(
                period => period.Key,
                period => BmesReportProjection.Finite(seed.IsSubtotal
                    ? SubtotalMonthPpm(index, period.Key, report, request.Groups)
                    : DetailMonthPpm(seed, period.Key, report, request.Groups)),
                StringComparer.Ordinal);

            double? share = ParseShare(seed.Share);
            int parentIndex = seed.IsSubtotal ? index : ParentSubtotalIndex(index);
            Dictionary<string, double?> weighted = periods.ToDictionary(
                period => period.Key,
                period => share is > 0 && parentIndex >= 0
                    ? BmesReportProjection.Finite(
                        SubtotalMonthPpm(parentIndex, period.Key, report, request.Groups) * share.Value)
                    : 0d,
                StringComparer.Ordinal);

            rows.Add(new CauseRowDto
            {
                RowId = $"cause::{index:D3}",
                Model = seed.Model,
                Type = NullIfEmpty(seed.Type),
                Process = NullIfEmpty(seed.Process),
                NgName = NullIfEmpty(seed.NgName),
                Number = int.TryParse(seed.Number, NumberStyles.Integer, CultureInfo.InvariantCulture, out int number)
                    ? number
                    : null,
                Cause = NullIfEmpty(seed.Cause),
                ShareRatio = share,
                IsSubtotal = seed.IsSubtotal,
                PpmByPeriod = ppm,
                WeightedPpmByPeriod = weighted,
            });
        }

        IReadOnlyList<CauseModelMonthlyRowDto> modelRows = Catalog
            .Select(row => row.Model)
            .Where(model => !string.IsNullOrWhiteSpace(model))
            .Distinct(StringComparer.Ordinal)
            .Select(model => new CauseModelMonthlyRowDto
            {
                Model = model,
                PpmByPeriod = periods.ToDictionary(
                    period => period.Key,
                    period => BmesReportProjection.Finite(ModelMonthPpm(model, period.Key, report, request.Groups)),
                    StringComparer.Ordinal),
            })
            .ToArray();

        return new CauseMonthlyTabDto
        {
            Periods = periods,
            Rows = rows,
            ModelMonthlyRows = modelRows,
        };
    }

    private static double SubtotalMonthPpm(
        int subtotalIndex,
        string periodKey,
        NgRateReportService.NgRateReport report,
        IReadOnlyList<ModelGroupRecord> groups)
    {
        if (subtotalIndex < 0 || subtotalIndex >= Catalog.Count)
            return 0;
        string model = Catalog[subtotalIndex].Model;
        double total = 0;
        for (int index = subtotalIndex + 1; index < Catalog.Count; index++)
        {
            CauseCatalogRow child = Catalog[index];
            if (child.IsSubtotal || !string.Equals(child.Model, model, StringComparison.Ordinal))
                break;
            total += DetailMonthPpm(child, periodKey, report, groups);
        }
        return total;
    }

    private static double DetailMonthPpm(
        CauseCatalogRow row,
        string periodKey,
        NgRateReportService.NgRateReport report,
        IReadOnlyList<ModelGroupRecord> groups)
    {
        if (string.IsNullOrWhiteSpace(row.Process) || string.IsNullOrWhiteSpace(row.NgName))
            return 0;
        HashSet<string> keys = FindModelKeys(row.Model, report, groups).ToHashSet(StringComparer.Ordinal);
        if (keys.Count == 0)
            return 0;

        double rawSum = 0;
        int rawCount = 0;
        foreach (var item in report.GroupRawIn)
        {
            var key = item.Key;
            if (!keys.Contains(key.Group) ||
                !string.Equals(key.PeriodKey, periodKey, StringComparison.Ordinal) ||
                !LooseTextMatches(key.PN, row.Process) ||
                !LooseTextMatches(key.NG, row.NgName) ||
                (!string.IsNullOrWhiteSpace(row.Type) && !LooseTextMatches(key.PT, row.Type)))
                continue;

            double ppm = item.Value.I > 0 ? Math.Round(item.Value.N / item.Value.I * 1_000_000d, 0) : 0;
            if (ppm <= 0)
                continue;
            rawSum += ppm;
            rawCount++;
        }
        if (rawCount > 0)
            return rawSum / rawCount;

        double sum = 0;
        int count = 0;
        foreach (NgRateReportService.ReasonRow reason in report.ReasonRows.Where(reason => !reason.IsTotal))
        {
            if (!LooseTextMatches(reason.ProcessName, row.Process) ||
                !LooseTextMatches(reason.NgName, row.NgName) ||
                (!string.IsNullOrWhiteSpace(row.Type) && !LooseTextMatches(reason.ProcessType, row.Type)))
                continue;
            foreach (NgRateReportService.GroupPivotRow group in reason.Groups)
            {
                if (!keys.Contains(group.GroupName))
                    continue;
                double value = group.Ppm.GetValueOrDefault(periodKey);
                if (value <= 0)
                    continue;
                sum += value;
                count++;
            }
        }
        return count == 0 ? 0 : sum / count;
    }

    private static double ModelMonthPpm(
        string model,
        string periodKey,
        NgRateReportService.NgRateReport report,
        IReadOnlyList<ModelGroupRecord> groups)
    {
        double sum = 0;
        int count = 0;
        foreach (string key in FindModelKeys(model, report, groups))
        {
            double value = report.GroupSummary.GetValueOrDefault(key)?
                .FirstOrDefault(row => row.IsTotal)?
                .Ppm.GetValueOrDefault(periodKey) ?? 0;
            if (value <= 0)
                continue;
            sum += value;
            count++;
        }
        return count == 0 ? 0 : sum / count;
    }

    private static IEnumerable<string> FindModelKeys(
        string model,
        NgRateReportService.NgRateReport report,
        IReadOnlyList<ModelGroupRecord> groups)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (ModelGroupRecord group in groups)
        foreach (MidGroupRecord mid in group.MidGroups)
        {
            if (!ModelMatches(mid.Material, model))
                continue;
            string key = NgRateModeSupport.ModelKey(group.Name, mid.Material);
            if (seen.Add(key))
                yield return key;
        }

        foreach (string key in report.GroupSummary.Keys)
        {
            int separator = key.IndexOf("::", StringComparison.Ordinal);
            string material = separator >= 0 && separator + 2 < key.Length ? key[(separator + 2)..] : key;
            if (ModelMatches(material, model) && seen.Add(key))
                yield return key;
        }
    }

    private static bool ModelMatches(string material, string requestedModel)
    {
        string materialKey = Normalize(material);
        if (materialKey.Length == 0)
            return false;
        foreach (string alias in ModelAliases(requestedModel))
        {
            string aliasKey = Normalize(alias);
            if (aliasKey.Length > 0 &&
                (string.Equals(materialKey, aliasKey, StringComparison.OrdinalIgnoreCase) ||
                 materialKey.Contains(aliasKey, StringComparison.OrdinalIgnoreCase) ||
                 aliasKey.Contains(materialKey, StringComparison.OrdinalIgnoreCase)))
                return true;
        }
        return false;
    }

    private static IEnumerable<string> ModelAliases(string model)
    {
        yield return model;
        string[] parts = model.Split('/', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length <= 1)
            yield break;
        int dash = parts[0].LastIndexOf('-');
        string prefix = dash >= 0 ? parts[0][..(dash + 1)] : string.Empty;
        foreach (string part in parts)
        {
            yield return part;
            if (prefix.Length > 0 && !part.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                yield return prefix + part;
        }
    }

    private static double? ParseShare(string share)
    {
        string text = (share ?? string.Empty).Trim().TrimEnd('%').Trim();
        if (!double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double value))
            return null;
        double ratio = value > 1 ? value / 100d : value;
        return ratio is >= 0 and <= 1 && double.IsFinite(ratio) ? ratio : null;
    }

    private static int ParentSubtotalIndex(int rowIndex)
    {
        if (rowIndex < 0 || rowIndex >= Catalog.Count)
            return -1;
        string model = Catalog[rowIndex].Model;
        for (int index = rowIndex - 1; index >= 0; index--)
        {
            CauseCatalogRow row = Catalog[index];
            if (!string.Equals(row.Model, model, StringComparison.Ordinal))
                break;
            if (row.IsSubtotal)
                return index;
        }
        return -1;
    }

    private static bool LooseTextMatches(string left, string right)
    {
        string leftKey = Normalize(left);
        string rightKey = Normalize(right);
        return leftKey.Length > 0 && rightKey.Length > 0 &&
               (string.Equals(leftKey, rightKey, StringComparison.OrdinalIgnoreCase) ||
                leftKey.Contains(rightKey, StringComparison.OrdinalIgnoreCase) ||
                rightKey.Contains(leftKey, StringComparison.OrdinalIgnoreCase));
    }

    private static string Normalize(string value) =>
        new((value ?? string.Empty).ToUpperInvariant().Where(char.IsLetterOrDigit).ToArray());

    private static string? NullIfEmpty(string value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;

    private sealed record CauseCatalogRow(
        string Model,
        string Type,
        string Process,
        string NgName,
        string Number,
        string Cause,
        string Share,
        bool IsSubtotal);

    private static CauseCatalogRow Subtotal(string model) => new(model, "", "", "", "", "", "", true);
    private static CauseCatalogRow Row(string model, string type, string process, string ngName, string number = "", string cause = "", string share = "") =>
        new(model, type, process, ngName, number, cause, share, false);

    private static readonly IReadOnlyList<CauseCatalogRow> Catalog =
    [
        Subtotal("BRS-161016S08ZZ"),
        Row("BRS-161016S08ZZ", "FUNCTION", "Hearing Test", "Noise Defect", "1", "AWF", "40%"),
        Row("BRS-161016S08ZZ", "FUNCTION", "Hearing Test", "Touch Defect", "2", "조립 설비 트러블", "20%"),
        Row("BRS-161016S08ZZ", "FUNCTION", "SIGMA TEST", "THD", "3", "원자재", "40%"),
        Row("BRS-161016S08ZZ", "FUNCTION", "SIGMA TEST", "SPL THD F0"),
        Row("BRS-161016S08ZZ", "FUNCTION", "SIGMA TEST", "SPL R&B"),
        Row("BRS-161016S08ZZ", "FUNCTION", "SIGMA TEST", "SPL THD"),
        Row("BRS-161016S08ZZ", "FUNCTION", "SIGMA TEST", "SPL"),
        Row("BRS-161016S08ZZ", "FUNCTION", "SIGMA TEST", "SPL R&B F0"),
        Row("BRS-161016S08ZZ", "FUNCTION", "SIGMA TEST", "R&B"),
        Subtotal("BRS-161016S08ZZ"),
        Row("BRS-161016S08ZZ", "SUB", "VP CD VISION INSPECTION", "VP CD separate", "1", "조립 설비 트러블", "50%"),
        Row("BRS-161016S08ZZ", "SUB", "VP CD VISION INSPECTION", "Offset", "2", "본드 선정 부적합", "50%"),
        Subtotal("BRS-161016S08ZZ"),
        Row("BRS-161016S08ZZ", "SUB", "YOKE VISION INSPECTION", "Separate Short SMG", "1", "조립 설비 트러블", "50%"),
        Row("BRS-161016S08ZZ", "SUB", "YOKE VISION INSPECTION 2", "Separated SMG Short", "2", "본드 선정 부적합", "50%"),
        Subtotal("BRS-161016S08ZZ"),
        Row("BRS-161016S08ZZ", "MAIN", "LONG EDGE V/P VISUAL INSPECTION", "Separate VP", "1", "공법 부적합", "70%"),
        Row("BRS-161016S08ZZ", "MAIN", "SHORT EDGE V/P VISUAL INSPECTION", "Separate VP", "2", "본딩 설비 트러블", "30%"),
        Row("BRS-161016S08ZZ", "MAIN", "SPK SEMI ARRAY", "VP saparate"),
        Subtotal("BRS-161016S08ZZ"),
        Row("BRS-161016S08ZZ", "MAIN", "S/P COIL ASS Y", "Other", "1", "조립 설비 트러블", "100%"),
        Subtotal("MSU-L20S15-07"),
        Row("MSU-L20S15-07", "FUNCTION", "Hearing Test", "Noise Defect", "1", "AWF", "65%"),
        Row("MSU-L20S15-07", "FUNCTION", "Hearing Test", "Touch Defect", "2", "조립 설비 트러블", "30%"),
        Row("MSU-L20S15-07", "FUNCTION", "SIGMA TEST", "SPL"),
        Row("MSU-L20S15-07", "FUNCTION", "SIGMA TEST", "SPL THD"),
        Row("MSU-L20S15-07", "FUNCTION", "SIGMA TEST", "THD"),
        Subtotal("MSU-L20S15-07"),
        Row("MSU-L20S15-07", "MAIN", "AWF", "NG Coil", "1", "AWF", "100%"),
        Row("MSU-L20S15-07", "MAIN", "Wire Array & Vision Inspection", "Over turn"),
        Row("MSU-L20S15-07", "MAIN", "Wire Array & Vision Inspection", "Don t cut wire"),
        Row("MSU-L20S15-07", "MAIN", "Wire Array & Vision Inspection", "Drop coil"),
        Subtotal("MSM-X526/X626B"),
        Row("MSM-X526/X626B", "FUNCTION", "BAKO TEST", "R&B NG", "1", "개발 ISSUE 미 종결", "80%"),
        Row("MSM-X526/X626B", "FUNCTION", "Hearing Test", "Noise Defect", "2", "AWF", "20%"),
        Subtotal("MSM-X526/X626B"),
        Row("MSM-X526/X626B", "FUNCTION", "AIR LEAK TEST", "High", "1", "JIG 마모", "50%"),
        Row("MSM-X526/X626B", "", "", "", "2", "조립 설비 트러블", "50%"),
        Subtotal("MSM-X526/X626B"),
        Row("MSM-X526/X626B", "VISUAL", "Visual Inspection", "Upper damage"),
        Row("MSM-X526/X626B", "VISUAL", "Visual Inspection", "Lower damage"),
        Row("MSM-X526/X626B", "VISUAL", "Visual Inspection", "Dome Damage"),
        Subtotal("MSM-S931B"),
        Row("MSM-S931B", "FUNCTION", "BAKO TEST", "R&B NG", "1", "원자재", "80%"),
        Row("MSM-S931B", "FUNCTION", "BAKO TEST", "Other", "2", "본딩 설비 트러블", "20%"),
        Row("MSM-S931B", "FUNCTION", "Hearing Test", "Noise Defect"),
        Subtotal("MSM-S931B"),
        Row("MSM-S931B", "FUNCTION", "AIR LEAK TEST 2", "High", "1", "원자재", "100%"),
        Row("MSM-S931B", "FUNCTION", "AIR LEAK TEST 2", "Other"),
        Row("MSM-S931B", "FUNCTION", "AIR LEAK TEST 2", "Lower"),
        Row("MSM-S931B", "FUNCTION", "AIR LEAK TEST 2", "NG airleak test"),
        Subtotal("TIU-C11-20"),
        Row("TIU-C11-20", "MAIN", "Fame Coil Inspection AI", "Froming NG", "1", "AWF", "95%"),
        Row("TIU-C11-20", "MAIN", "Fame Coil Inspection AI", "wire separated", "2", "원자재", "5%"),
        Row("TIU-C11-20", "MAIN", "Fame Coil Inspection AI", "Offset"),
        Row("TIU-C11-20", "MAIN", "Fame Coil Inspection AI", "Coil damage"),
        Row("TIU-C11-20", "MAIN", "Spot Welding Inspection AI", "Wire offset"),
        Row("TIU-C11-20", "MAIN", "Grill Ass y Inspection AI", "Coil damage"),
        Subtotal("TIU-C11-20"),
        Row("TIU-C11-20", "MAIN", "Laser Cutting VP Inspection AI", "No Cuting", "1", "조립 설비 트러블", "70%"),
        Row("TIU-C11-20", "MAIN", "Laser Cutting VP Inspection AI", "Offset", "2", "공법 부적합", "30%"),
        Subtotal("TIU-C11-20"),
        Row("TIU-C11-20", "MAIN", "Frame Gluing Inspection AI", "Offset bonding", "1", "본딩 설비 트러블", "50%"),
        Row("TIU-C11-20", "MAIN", "Frame Gluing Inspection AI", "Over glue FS", "2", "공법 부적합", "50%"),
        Subtotal("TIU-C11-20"),
        Row("TIU-C11-20", "SUB", "VP Dome Vision Inspection", "NG over bond", "1", "본딩 설비 트러블", "40%"),
        Row("TIU-C11-20", "SUB", "VP Dome Vision Inspection", "NG short bond,", "2", "조립 설비 트러블", "40%"),
        Row("TIU-C11-20", "SUB", "Vision Inspection Ass y YOKE", "offset", "3", "원자재", "10%"),
        Row("TIU-C11-20", "SUB", "V/P Dome Inspection AI", "over glue", "4", "공법 부적합", "10%"),
        Row("TIU-C11-20", "SUB", "V/P Dome Inspection AI", "Offset"),
        Row("TIU-C11-20", "SUB", "V/P Gluing Inspection AI", "Over glue"),
        Row("TIU-C11-20", "SUB", "V/P Gluing Inspection AI", "VP damage"),
        Row("TIU-C11-20", "SUB", "V/P Gluing Inspection AI", "Leak glue"),
        Row("TIU-C11-20", "SUB", "V/P Gluing Inspection AI", "Offset"),
        Subtotal("TIU-C11-20"),
        Row("TIU-C11-20", "FUNCTION", "Hearing Test", "Noise", "1", "이물", "70%"),
        Row("TIU-C11-20", "FUNCTION", "Hearing Test", "Touch", "2", "AWF", "25%"),
        Row("TIU-C11-20", "FUNCTION", "Hearing Test", "Low sound", "3", "작업자 터치", "5%"),
        Row("TIU-C11-20", "FUNCTION", "Audio Bus Test", "SPL"),
        Subtotal("TIU-L5S3-01"),
        Row("TIU-L5S3-01", "FUNCTION", "FR(FREQUENCY RESPONSE) TEST", "High frequency", "1", "공법 부적합", "30%"),
        Row("TIU-L5S3-01", "FUNCTION", "FR(FREQUENCY RESPONSE) TEST", "Low frequency", "2", "조립 설비 트러블", "30%"),
        Row("TIU-L5S3-01", "FUNCTION", "Hearing Test", "Low sound", "3", "본딩 설비 트러블", "30%"),
        Row("TIU-L5S3-01", "", "", "", "4", "JIG 마모", "10%"),
        Subtotal("TIU-L5S3-01"),
        Row("TIU-L5S3-01", "SUB", "Semi Frame Vision Inspection", "Separated F-PCB", "1", "조립 설비 트러블", "100%"),
        Row("TIU-L5S3-01", "SUB", "Semi Frame Vision Inspection", "Offset F-PCB"),
        Row("TIU-L5S3-01", "SUB", "Frame F-PCB Inspection AI", "Separated F-PCB"),
        Row("TIU-L5S3-01", "SUB", "Frame F-PCB Inspection AI", "Offset F-PCB"),
        Subtotal("TIU-L5S3-01"),
        Row("TIU-L5S3-01", "SUB", "Semi Frame Vision Inspection", "Damage", "1", "조립 설비 트러블", "100%"),
        Row("TIU-L5S3-01", "SUB", "Frame F-PCB Inspection AI", "Damage"),
        Row("TIU-L5S3-01", "MAIN", "V/P Gluing Inspection AI", "Frame damages"),
        Subtotal("TIU-L5S3-01"),
        Row("TIU-L5S3-01", "SUB", "Semi Yoke Vision Inspection", "offset yoke", "1", "조립 설비 트러블", "100%"),
        Row("TIU-L5S3-01", "SUB", "Semi Yoke Vision Inspection", "Damages"),
        Row("TIU-L5S3-01", "SUB", "AI Yoke Inspection", "Damage"),
        Row("TIU-L5S3-01", "SUB", "AI Yoke Inspection", "Offset Yoke"),
        Subtotal("TIU-L5S3-01"),
        Row("TIU-L5S3-01", "SUB", "Frame Gluing Inspection AI", "Offset bonding", "1", "본딩 설비 트러블", "100%"),
        Row("TIU-L5S3-01", "SUB", "Frame Gluing Inspection AI", "Over glue"),
        Row("TIU-L5S3-01", "SUB", "Frame Gluing Inspection AI", "Leak glue"),
        Row("TIU-L5S3-01", "SUB", "Frame Gluing Inspection AI 2", "Over glue"),
        Row("TIU-L5S3-01", "SUB", "Frame Gluing Inspection AI 2", "Offset glue"),
        Row("TIU-L5S3-01", "SUB", "Frame Gluing Inspection AI 2", "Leak glue"),
        Row("TIU-L5S3-01", "VISUAL", "Visual Inspection", "Leak glue"),
        Row("TIU-L5S3-01", "VISUAL", "Visual Inspection", "Over glue"),
        Subtotal("TIU-L5S3-01"),
        Row("TIU-L5S3-01", "MAIN", "Wire Pattern Inspection AI", "Offset", "1", "AWF", "100%"),
        Row("TIU-L5S3-01", "MAIN", "Wire Pattern Inspection AI", "Coil Separated"),
        Row("TIU-L5S3-01", "MAIN", "Wire Pattern Inspection AI", "Coil damage"),
        Row("TIU-L5S3-01", "MAIN", "Wire Pattern Inspection AI", "Coil shortage"),
        Row("TIU-L5S3-01", "MAIN", "Spot Welding Inspection AI", "Wire offset"),
        Subtotal("TIU-L5S3-01"),
        Row("TIU-L5S3-01", "MAIN", "V/P Gluing Inspection AI", "Leak glue", "1", "본딩 설비 트러블", "100%"),
        Row("TIU-L5S3-01", "MAIN", "V/P Gluing Inspection AI", "Offset"),
        Row("TIU-L5S3-01", "MAIN", "V/P Gluing Inspection AI", "Over glue"),
        Subtotal("ASSY 338 RA1"),
        Row("ASSY 338 RA1", "VISUAL", "VISUAL INSPECTION 3", "Scratch Rear", "1", "원자재", "40%"),
        Row("ASSY 338 RA1", "VISUAL", "VISUAL INSPECTION 3", "Damage Rear", "2", "공법 부적합", "60%"),
        Row("ASSY 338 RA1", "VISUAL", "VISUAL INSPECTION 3", "Lumpy face Rear"),
        Row("ASSY 338 RA1", "VISUAL", "EDGE Surface Vision Check", "Damage Rear"),
        Row("ASSY 338 RA1", "MAIN", "MIC TWITER Ass y Vision Check", "Damage Rear"),
        Subtotal("ASSY 338 RA1"),
        Row("ASSY 338 RA1", "SUB4", "Hook Base BRK Inner Spring Ass y Vis", "Damage Bracket", "1", "공법 부적합", "50%"),
        Row("ASSY 338 RA1", "SUB4", "Hook Base BRK Inner Spring Ass y Vis", "Offset pin", "2", "JIG 마모", "50%"),
        Row("ASSY 338 RA1", "SUB4", "VISUAL INSPECTION", "Offset assembly Spring"),
        Row("ASSY 338 RA1", "SUB4", "VISUAL INSPECTION", "Damage Bracket"),
        Row("ASSY 338 RA1", "MAIN", "BRACKET INNER Ass y Vision Check", "Damage Bracket"),
        Row("ASSY 338 RA1", "VISUAL", "VISUAL INSPECTION 3", "Damage BRK"),
        Subtotal("ASSY 338 RA1"),
        Row("ASSY 338 RA1", "VISUAL", "VISUAL INSPECTION 3", "Deform Main port", "1", "조립 설비 트러블", "100%"),
        Row("ASSY 338 RA1", "VISUAL", "VISUAL INSPECTION 3", "Damage Main port"),
        Subtotal("ASSY 338 RA1"),
        Row("ASSY 338 RA1", "FUNCTION", "SOUND QUALITY TEST", "Noise", "1"),
    ];
}
