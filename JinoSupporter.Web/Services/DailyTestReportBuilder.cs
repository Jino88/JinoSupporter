using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace JinoSupporter.Web.Services;

public static class DailyTestReportBuilder
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    public static bool HasBuilderSpec(string parametersJson)
    {
        if (!TryParseObject(parametersJson, out JsonObject? root) || root is null) return false;
        return root.ContainsKey("reportBuilderSpec")
               || root.ContainsKey("analysisContract") && root.ContainsKey("analysisData");
    }

    public static bool TryMergeParameters(string existingParametersJson, string addedParametersJson, out string mergedParametersJson)
    {
        mergedParametersJson = "";
        if (!TryParseObject(existingParametersJson, out JsonObject? existing)
            || existing is null
            || !TryParseObject(addedParametersJson, out JsonObject? added)
            || added is null)
        {
            return false;
        }

        JsonObject merged = CloneObject(existing);
        merged["scope"] = "cumulative";
        merged["version"] = 1;

        if (existing["dataSummary"] is JsonObject || added["dataSummary"] is JsonObject)
        {
            merged["dataSummary"] = MergeDataSummary(
                existing["dataSummary"] as JsonObject,
                added["dataSummary"] as JsonObject);
        }

        JsonObject mergedData = MergeObjects(
            existing["analysisData"] as JsonObject,
            added["analysisData"] as JsonObject,
            "analysisData");
        if (mergedData.Count > 0)
            merged["analysisData"] = mergedData;

        if (!merged.ContainsKey("reportBuilderSpec"))
            merged["reportBuilderSpec"] = BuildFallbackBuilderSpec(merged);

        mergedParametersJson = merged.ToJsonString(JsonOptions);
        return HasRenderableAnalysisData(merged);
    }

    public static bool TryRenderHtml(string parametersJson, out string html)
    {
        html = "";
        if (!TryParseObject(parametersJson, out JsonObject? root) || root is null || !HasRenderableAnalysisData(root))
            return false;

        var sb = new StringBuilder();
        string projectName = FirstText(root["projectName"], root["name"], "Daily Test Data");
        JsonObject? dataSummary = root["dataSummary"] as JsonObject;
        JsonObject? analysisData = root["analysisData"] as JsonObject;
        JsonObject? builderSpec = root["reportBuilderSpec"] as JsonObject;

        AppendHeading(sb, projectName);
        AppendActionBoard(sb, analysisData, builderSpec);
        AppendSummaryCards(sb, dataSummary, analysisData?["totals"] as JsonObject);
        AppendMatrices(sb, analysisData?["matrices"] as JsonArray, builderSpec);
        AppendTopRisks(sb, analysisData?["topRisks"] as JsonArray);
        AppendAggregates(sb, analysisData?["aggregates"] as JsonObject);
        AppendNotes(sb, root, builderSpec);

        html = sb.ToString().Trim();
        return !string.IsNullOrWhiteSpace(html);
    }

    private static bool TryParseObject(string json, out JsonObject? obj)
    {
        obj = null;
        if (string.IsNullOrWhiteSpace(json)) return false;
        try
        {
            obj = JsonNode.Parse(json) as JsonObject;
            return obj is not null;
        }
        catch
        {
            return false;
        }
    }

    private static JsonObject CloneObject(JsonObject source)
        => JsonNode.Parse(source.ToJsonString()) as JsonObject ?? new JsonObject();

    private static JsonNode? CloneNode(JsonNode? node)
        => node is null ? null : JsonNode.Parse(node.ToJsonString());

    private static JsonObject MergeDataSummary(JsonObject? existing, JsonObject? added)
    {
        JsonObject merged = existing is null ? new JsonObject() : CloneObject(existing);
        if (added is null) return merged;

        foreach ((string key, JsonNode? value) in added)
        {
            if (value is null) continue;

            if (IsRowCountKey(key))
            {
                double sum = Number(merged[key]) + Number(value);
                if (sum > 0) merged[key] = sum % 1 == 0 ? (int)sum : sum;
                continue;
            }

            if (value is JsonArray addedArray)
            {
                merged[key] = MergeArrays(merged[key] as JsonArray, addedArray, key);
                continue;
            }

            string addedText = Text(value);
            if (string.IsNullOrWhiteSpace(addedText)) continue;

            if (key.Contains("start", StringComparison.OrdinalIgnoreCase))
            {
                string current = Text(merged[key]);
                merged[key] = EarliestNonEmpty(current, addedText);
            }
            else if (key.Contains("end", StringComparison.OrdinalIgnoreCase))
            {
                string current = Text(merged[key]);
                merged[key] = LatestNonEmpty(current, addedText);
            }
            else if (string.IsNullOrWhiteSpace(Text(merged[key])))
            {
                merged[key] = CloneNode(value);
            }
        }

        merged["coverage"] = "Updated by saved Daily Test report builder.";
        return merged;
    }

    private static JsonObject MergeObjects(JsonObject? existing, JsonObject? added, string objectName)
    {
        JsonObject merged = existing is null ? new JsonObject() : CloneObject(existing);
        if (added is null) return merged;

        foreach ((string key, JsonNode? value) in added)
        {
            if (value is null) continue;
            JsonNode? current = merged[key];

            if (current is JsonObject currentObj && value is JsonObject addedObj)
            {
                merged[key] = MergeObjects(currentObj, addedObj, key);
            }
            else if (current is JsonArray currentArray && value is JsonArray addedArray)
            {
                merged[key] = MergeArrays(currentArray, addedArray, key);
            }
            else if (IsNumber(current) && IsNumber(value))
            {
                if (IsRateLikeKey(key))
                    merged[key] = CloneNode(current);
                else
                    merged[key] = NormalizeNumber(Number(current) + Number(value));
            }
            else if (current is null)
            {
                merged[key] = CloneNode(value);
            }
            else if (string.IsNullOrWhiteSpace(Text(current)) && !string.IsNullOrWhiteSpace(Text(value)))
            {
                merged[key] = CloneNode(value);
            }
        }

        RecomputeRates(merged);
        return merged;
    }

    private static JsonArray MergeArrays(JsonArray? existing, JsonArray added, string key)
    {
        if (key.Equals("matrices", StringComparison.OrdinalIgnoreCase))
            return MergeMatrices(existing, added);

        if (key.Equals("topRisks", StringComparison.OrdinalIgnoreCase))
            return MergeObjectArray(existing, added, itemLimit: 50, sortByRisk: true);

        if (key.Contains("file", StringComparison.OrdinalIgnoreCase))
            return MergePrimitiveArray(existing, added, itemLimit: 200);

        return MergeObjectArray(existing, added, itemLimit: 500, sortByRisk: false);
    }

    private static JsonArray MergePrimitiveArray(JsonArray? existing, JsonArray added, int itemLimit)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var output = new JsonArray();

        void Add(JsonNode? node)
        {
            string text = Text(node);
            if (string.IsNullOrWhiteSpace(text) || !seen.Add(text) || output.Count >= itemLimit) return;
            output.Add(text);
        }

        if (existing is not null)
            foreach (JsonNode? node in existing) Add(node);
        foreach (JsonNode? node in added) Add(node);

        return output;
    }

    private static JsonArray MergeObjectArray(JsonArray? existing, JsonArray added, int itemLimit, bool sortByRisk)
    {
        var byKey = new Dictionary<string, JsonObject>(StringComparer.OrdinalIgnoreCase);
        var order = new List<string>();

        void Add(JsonNode? node)
        {
            if (node is not JsonObject obj)
            {
                string primitiveKey = Text(node);
                if (string.IsNullOrWhiteSpace(primitiveKey) || byKey.ContainsKey(primitiveKey)) return;
                byKey[primitiveKey] = new JsonObject { ["value"] = primitiveKey };
                order.Add(primitiveKey);
                return;
            }

            string itemKey = ObjectKey(obj);
            if (string.IsNullOrWhiteSpace(itemKey))
                itemKey = obj.ToJsonString();

            if (byKey.TryGetValue(itemKey, out JsonObject? current))
            {
                byKey[itemKey] = MergeObjects(current, obj, itemKey);
            }
            else
            {
                byKey[itemKey] = CloneObject(obj);
                order.Add(itemKey);
            }
        }

        if (existing is not null)
            foreach (JsonNode? node in existing) Add(node);
        foreach (JsonNode? node in added) Add(node);

        IEnumerable<JsonObject> items = order.Select(key => byKey[key]);
        if (sortByRisk)
            items = items.OrderByDescending(RiskValue);

        var output = new JsonArray();
        foreach (JsonObject item in items.Take(itemLimit))
            output.Add(CloneNode(item));
        return output;
    }

    private static JsonArray MergeMatrices(JsonArray? existing, JsonArray added)
    {
        var byKey = new Dictionary<string, JsonObject>(StringComparer.OrdinalIgnoreCase);
        var order = new List<string>();

        void Add(JsonNode? node)
        {
            if (node is not JsonObject matrix) return;
            string key = FirstText(matrix["id"], matrix["title"], MatrixDimensionKey(matrix), Guid.NewGuid().ToString("N"));
            if (!byKey.TryGetValue(key, out JsonObject? current))
            {
                byKey[key] = CloneObject(matrix);
                order.Add(key);
                return;
            }

            JsonObject merged = MergeObjects(current, matrix, key);
            merged["cells"] = MergeMatrixCells(current["cells"] as JsonArray, matrix["cells"] as JsonArray);
            byKey[key] = merged;
        }

        if (existing is not null)
            foreach (JsonNode? node in existing) Add(node);
        foreach (JsonNode? node in added) Add(node);

        var output = new JsonArray();
        foreach (string key in order)
            output.Add(CloneNode(byKey[key]));
        return output;
    }

    private static JsonArray MergeMatrixCells(JsonArray? existing, JsonArray? added)
    {
        var byKey = new Dictionary<string, JsonObject>(StringComparer.OrdinalIgnoreCase);
        var order = new List<string>();

        void Add(JsonNode? node)
        {
            if (node is not JsonObject cell) return;
            string key = CellKey(cell);
            if (string.IsNullOrWhiteSpace(key)) return;
            if (byKey.TryGetValue(key, out JsonObject? current))
            {
                byKey[key] = MergeObjects(current, cell, key);
            }
            else
            {
                byKey[key] = CloneObject(cell);
                order.Add(key);
            }
        }

        if (existing is not null)
            foreach (JsonNode? node in existing) Add(node);
        if (added is not null)
            foreach (JsonNode? node in added) Add(node);

        var output = new JsonArray();
        foreach (string key in order)
            output.Add(CloneNode(byKey[key]));
        return output;
    }

    private static JsonObject BuildFallbackBuilderSpec(JsonObject parameters)
        => new()
        {
            ["version"] = 1,
            ["mode"] = "programmatic-html-maker",
            ["layout"] = new JsonArray("actionBoard", "kpiSummary", "matrices", "topRisks", "aggregates", "notes"),
            ["renderRules"] = new JsonObject
            {
                ["rateFormat"] = "ordinary rate first, NG/Total second; LOT-vs-NORMAL cells show LOT NG ppm and NORMAL NG ppm",
                ["emptyCellText"] = "-",
                ["source"] = "fallback from analysisContract and analysisData"
            }
        };

    private static bool HasRenderableAnalysisData(JsonObject root)
    {
        if (root["analysisData"] is not JsonObject data) return false;
        return data.Count > 0
               && (data["totals"] is JsonObject
                   || data["aggregates"] is JsonObject
                   || data["matrices"] is JsonArray
                   || data["topRisks"] is JsonArray);
    }

    private static void AppendHeading(StringBuilder sb, string projectName)
    {
        sb.Append("<h2>").Append(Enc(projectName)).Append("</h2>");
        sb.Append("<p><strong>Program-built report</strong> generated from the saved Daily Test HTML Maker spec and cumulative parameters.</p>");
    }

    private static void AppendActionBoard(StringBuilder sb, JsonObject? analysisData, JsonObject? builderSpec)
    {
        JsonArray? risks = analysisData?["topRisks"] as JsonArray;
        sb.Append("<h3>Action Board</h3>");
        sb.Append("<div style=\"display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:8px;margin:8px 0 12px 0;\">");

        if (risks is not null && risks.Count > 0)
        {
            foreach (JsonObject risk in risks.OfType<JsonObject>().Take(4))
            {
                string label = FirstText(risk["label"], risk["name"], risk["dimension"], risk["row"], risk["item"], "Risk");
                string value = FirstText(risk["rate"], risk["value"], risk["ngRate"], risk["score"], "");
                sb.Append("<div style=\"border:1px solid #e2e8f0;border-left:4px solid #dc2626;padding:8px;border-radius:6px;background:#fff7f7;\">")
                    .Append("<strong>").Append(Enc(label)).Append("</strong>");
                if (!string.IsNullOrWhiteSpace(value))
                    sb.Append("<br><span>").Append(Enc(value)).Append("</span>");
                sb.Append("</div>");
            }
        }
        else
        {
            sb.Append("<div style=\"border:1px solid #e2e8f0;border-left:4px solid #2563eb;padding:8px;border-radius:6px;background:#f8fbff;\">")
                .Append("<strong>Updated</strong><br><span>Cumulative data merged with saved report builder.</span></div>");
        }

        sb.Append("</div>");
    }

    private static void AppendSummaryCards(StringBuilder sb, JsonObject? dataSummary, JsonObject? totals)
    {
        var cards = new List<(string Label, string Value)>();
        AddSummaryCards(cards, dataSummary, "Data");
        AddSummaryCards(cards, totals, "Total");
        if (cards.Count == 0) return;

        sb.Append("<h3>KPI Summary</h3>");
        sb.Append("<div style=\"display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin:8px 0 12px 0;\">");
        foreach ((string label, string value) in cards.Take(12))
        {
            sb.Append("<div style=\"border:1px solid #dbe4ee;border-radius:6px;padding:8px;background:#ffffff;\">")
                .Append("<span style=\"display:block;color:#64748b;font-size:12px;\">").Append(Enc(label)).Append("</span>")
                .Append("<strong>").Append(Enc(value)).Append("</strong>")
                .Append("</div>");
        }
        sb.Append("</div>");
    }

    private static void AddSummaryCards(List<(string Label, string Value)> cards, JsonObject? obj, string prefix)
    {
        if (obj is null) return;
        foreach ((string key, JsonNode? value) in obj)
        {
            if (value is JsonObject or JsonArray) continue;
            string text = Text(value);
            if (string.IsNullOrWhiteSpace(text)) continue;
            cards.Add((Humanize(key), text));
        }
    }

    private static void AppendMatrices(StringBuilder sb, JsonArray? matrices, JsonObject? builderSpec)
    {
        if (matrices is null || matrices.Count == 0) return;
        foreach (JsonObject matrix in matrices.OfType<JsonObject>())
        {
            JsonArray? cells = matrix["cells"] as JsonArray;
            if (cells is null || cells.Count == 0) continue;

            string title = FirstText(matrix["title"], matrix["name"], "Matrix");
            string rowDimension = FirstText(matrix["rowDimension"], matrix["rowField"], "Row");
            string columnDimension = FirstText(matrix["columnDimension"], matrix["columnField"], "Column");
            string valueMetric = FirstText(matrix["valueMetric"], matrix["metric"], "Value");
            string countBasis = FirstText(matrix["countBasis"], matrix["basis"], "");

            var rows = cells.OfType<JsonObject>().Select(RowKey).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.Ordinal).ToList();
            var cols = cells.OfType<JsonObject>().Select(ColumnKey).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.Ordinal).ToList();
            var byCell = cells.OfType<JsonObject>().ToDictionary(c => RowKey(c) + "\u001f" + ColumnKey(c), c => c, StringComparer.Ordinal);
            if (rows.Count == 0 || cols.Count == 0) continue;

            sb.Append("<h3>").Append(Enc(title)).Append("</h3>");
            sb.Append("<table style=\"border-collapse:collapse;width:100%;margin:8px 0 12px 0;font-size:12px;\">");
            sb.Append("<thead><tr><th style=\"border:1px solid #cbd5e1;padding:6px;background:#f1f5f9;text-align:left;\">")
                .Append(Enc(rowDimension)).Append(" \\ ").Append(Enc(columnDimension)).Append("</th>");
            foreach (string col in cols)
                sb.Append("<th style=\"border:1px solid #cbd5e1;padding:6px;background:#f1f5f9;text-align:left;\">").Append(Enc(col)).Append("</th>");
            sb.Append("</tr></thead><tbody>");

            foreach (string row in rows)
            {
                sb.Append("<tr><th style=\"border:1px solid #cbd5e1;padding:6px;background:#f8fafc;text-align:left;\">").Append(Enc(row)).Append("</th>");
                foreach (string col in cols)
                {
                    if (!byCell.TryGetValue(row + "\u001f" + col, out JsonObject? cell))
                    {
                        sb.Append("<td style=\"border:1px solid #cbd5e1;padding:6px;color:#94a3b8;\">-</td>");
                        continue;
                    }

                    string value = CellValueText(cell, valueMetric, countBasis);
                    sb.Append("<td style=\"border:1px solid #cbd5e1;padding:6px;")
                        .Append(HeatStyle(RiskValue(cell))).Append("\">")
                        .Append(value)
                        .Append("</td>");
                }
                sb.Append("</tr>");
            }

            sb.Append("</tbody></table>");
        }
    }

    private static void AppendTopRisks(StringBuilder sb, JsonArray? risks)
    {
        if (risks is null || risks.Count == 0) return;
        List<JsonObject> items = risks.OfType<JsonObject>().Take(20).ToList();
        if (items.Count == 0) return;
        List<string> keys = items.SelectMany(o => o.Select(p => p.Key)).Distinct(StringComparer.OrdinalIgnoreCase).Take(8).ToList();
        AppendObjectTable(sb, "Top Risks", items, keys);
    }

    private static void AppendAggregates(StringBuilder sb, JsonObject? aggregates)
    {
        if (aggregates is null || aggregates.Count == 0) return;
        foreach ((string key, JsonNode? value) in aggregates.Take(4))
        {
            if (value is JsonArray arr)
            {
                List<JsonObject> items = arr.OfType<JsonObject>().Take(20).ToList();
                if (items.Count == 0) continue;
                List<string> keys = items.SelectMany(o => o.Select(p => p.Key)).Distinct(StringComparer.OrdinalIgnoreCase).Take(8).ToList();
                AppendObjectTable(sb, Humanize(key), items, keys);
            }
            else if (value is JsonObject obj)
            {
                AppendKeyValueTable(sb, Humanize(key), obj);
            }
        }
    }

    private static void AppendObjectTable(StringBuilder sb, string title, IReadOnlyList<JsonObject> items, IReadOnlyList<string> keys)
    {
        if (items.Count == 0 || keys.Count == 0) return;
        sb.Append("<h3>").Append(Enc(title)).Append("</h3>");
        sb.Append("<table style=\"border-collapse:collapse;width:100%;margin:8px 0 12px 0;font-size:12px;\"><thead><tr>");
        foreach (string key in keys)
            sb.Append("<th style=\"border:1px solid #cbd5e1;padding:6px;background:#f1f5f9;text-align:left;\">").Append(Enc(Humanize(key))).Append("</th>");
        sb.Append("</tr></thead><tbody>");
        foreach (JsonObject item in items)
        {
            sb.Append("<tr>");
            foreach (string key in keys)
                sb.Append("<td style=\"border:1px solid #cbd5e1;padding:6px;\">").Append(Enc(Text(item[key]))).Append("</td>");
            sb.Append("</tr>");
        }
        sb.Append("</tbody></table>");
    }

    private static void AppendKeyValueTable(StringBuilder sb, string title, JsonObject obj)
    {
        sb.Append("<h3>").Append(Enc(title)).Append("</h3>");
        sb.Append("<table style=\"border-collapse:collapse;width:100%;margin:8px 0 12px 0;font-size:12px;\"><tbody>");
        foreach ((string key, JsonNode? value) in obj)
        {
            if (value is JsonObject or JsonArray) continue;
            sb.Append("<tr><th style=\"border:1px solid #cbd5e1;padding:6px;background:#f8fafc;text-align:left;\">")
                .Append(Enc(Humanize(key))).Append("</th><td style=\"border:1px solid #cbd5e1;padding:6px;\">")
                .Append(Enc(Text(value))).Append("</td></tr>");
        }
        sb.Append("</tbody></table>");
    }

    private static void AppendNotes(StringBuilder sb, JsonObject root, JsonObject? builderSpec)
    {
        sb.Append("<h3>Notes</h3><ul>");
        sb.Append("<li>HTML was regenerated by the application from saved parameters, not re-authored from raw cumulative data.</li>");
        if (builderSpec is not null)
            sb.Append("<li>Saved reportBuilderSpec was used as the HTML Maker contract.</li>");
        sb.Append("</ul>");
    }

    private static string CellValueText(JsonObject cell, string valueMetric, string countBasis)
    {
        double? lotRate = FirstNullableNumber(cell,
            "lotNgRate", "lotRate", "targetNgRate", "targetRate", "rate", "ngRate", "value", "metricValue", valueMetric);
        double? normalRate = FirstNullableNumber(cell,
            "normalNgRate", "normalRate", "sameDateNormalNgRate", "sameDateNormalRate", "baselineNgRate", "baselineRate", "referenceNgRate", "referenceRate");
        if (lotRate.HasValue && normalRate.HasValue)
        {
            double diff = lotRate.Value - normalRate.Value;
            string lotNg = FirstText(cell["lotNg"], cell["targetNg"], cell["ng"], cell["NG"], cell["defect"], cell["defects"], "");
            string lotTotal = FirstText(cell["lotTotal"], cell["targetTotal"], cell["total"], cell["Total"], cell["count"], cell["sample"], "");
            string normalNg = FirstText(cell["normalNg"], cell["baselineNg"], cell["referenceNg"], "");
            string normalTotal = FirstText(cell["normalTotal"], cell["baselineTotal"], cell["referenceTotal"], "");

            var compare = new StringBuilder();
            compare.Append("<strong>DIFF : ").Append(Enc(FormatPpm(diff, signed: true))).Append("</strong>");
            compare.Append("<br><span style=\"color:#475569;\">LOT NG : ").Append(Enc(FormatPpm(lotRate.Value, signed: false))).Append("</span>");
            if (!string.IsNullOrWhiteSpace(lotNg) && !string.IsNullOrWhiteSpace(lotTotal))
                compare.Append("<br><span style=\"color:#64748b;\">LOT COUNT : ").Append(Enc(lotNg)).Append("/").Append(Enc(lotTotal)).Append("</span>");
            compare.Append("<br><span style=\"color:#475569;\">NORMAL NG : ").Append(Enc(FormatPpm(normalRate.Value, signed: false))).Append("</span>");
            if (!string.IsNullOrWhiteSpace(normalNg) && !string.IsNullOrWhiteSpace(normalTotal))
                compare.Append("<br><span style=\"color:#64748b;\">NORMAL COUNT : ").Append(Enc(normalNg)).Append("/").Append(Enc(normalTotal)).Append("</span>");
            return compare.ToString();
        }

        string value = FirstText(cell["rate"], cell["ngRate"], cell["value"], cell["metricValue"], cell[valueMetric], "");
        string ng = FirstText(cell["ng"], cell["NG"], cell["defect"], cell["defects"], "");
        string total = FirstText(cell["total"], cell["Total"], cell["count"], cell["sample"], "");
        var sb = new StringBuilder();
        sb.Append("<strong>").Append(Enc(string.IsNullOrWhiteSpace(value) ? "metric unavailable" : FormatMetric(value))).Append("</strong>");
        if (!string.IsNullOrWhiteSpace(ng) && !string.IsNullOrWhiteSpace(total))
            sb.Append("<br><span style=\"color:#475569;\">").Append(Enc(ng)).Append("/").Append(Enc(total)).Append("</span>");
        else if (!string.IsNullOrWhiteSpace(countBasis))
            sb.Append("<br><span style=\"color:#475569;\">").Append(Enc(countBasis)).Append("</span>");
        return sb.ToString();
    }

    private static string FormatMetric(string value)
    {
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double n))
            return n is >= 0 and <= 100 ? n.ToString("0.##", CultureInfo.InvariantCulture) + "%" : n.ToString("0.##", CultureInfo.InvariantCulture);
        return value;
    }

    private static string FormatPpm(double rate, bool signed)
    {
        double fraction = Math.Abs(rate) <= 1.0 ? rate : rate / 100.0;
        long ppm = (long)Math.Round(fraction * 1_000_000.0);
        string prefix = signed && ppm > 0 ? "+" : "";
        return prefix + ppm.ToString("N0", CultureInfo.InvariantCulture) + " ppm";
    }

    private static string HeatStyle(double value)
    {
        if (value >= 15) return "background:#f8d7da;color:#991b1b;";
        if (value >= 8) return "background:#fff3cd;color:#854d0e;";
        if (value > 0) return "background:#d1e7dd;color:#065f46;";
        return "background:#ffffff;";
    }

    private static void RecomputeRates(JsonObject obj)
    {
        foreach ((string _, JsonNode? node) in obj.ToList())
        {
            if (node is JsonObject child) RecomputeRates(child);
            if (node is JsonArray arr)
                foreach (JsonObject arrObj in arr.OfType<JsonObject>()) RecomputeRates(arrObj);
        }

        double ng = FirstNumber(obj, "ng", "NG", "defect", "defects", "fail", "fails");
        double total = FirstNumber(obj, "total", "Total", "sample", "samples", "count");
        if (ng >= 0 && total > 0)
        {
            double rate = ng / total * 100.0;
            string? rateKey = obj.Select(p => p.Key).FirstOrDefault(IsRateLikeKey);
            obj[rateKey ?? "rate"] = Math.Round(rate, 4);
        }
    }

    private static string ObjectKey(JsonObject obj)
        => FirstText(obj["id"], obj["key"], obj["name"], obj["label"], obj["dimension"], obj["group"], obj["row"], obj["column"], "");

    private static string MatrixDimensionKey(JsonObject matrix)
        => FirstText(matrix["rowDimension"], matrix["rowField"], "") + "|" + FirstText(matrix["columnDimension"], matrix["columnField"], "");

    private static string CellKey(JsonObject cell)
        => RowKey(cell) + "\u001f" + ColumnKey(cell);

    private static string RowKey(JsonObject cell)
        => FirstText(cell["row"], cell["rowValue"], cell["rowKey"], cell["y"], cell["dimension1"], "");

    private static string ColumnKey(JsonObject cell)
        => FirstText(cell["column"], cell["columnValue"], cell["columnKey"], cell["x"], cell["dimension2"], "");

    private static double RiskValue(JsonObject obj)
        => FirstNumber(obj, "rate", "ngRate", "value", "score", "percent");

    private static double FirstNumber(JsonObject obj, params string[] names)
    {
        foreach (string name in names)
        {
            if (obj.TryGetPropertyValue(name, out JsonNode? node) && TryNumber(node, out double value))
                return value;
        }
        return -1;
    }

    private static double? FirstNullableNumber(JsonObject obj, params string[] names)
    {
        foreach (string name in names)
        {
            if (string.IsNullOrWhiteSpace(name)) continue;
            if (obj.TryGetPropertyValue(name, out JsonNode? node) && TryNumber(node, out double value))
                return value;
        }
        return null;
    }

    private static string FirstText(params object?[] values)
    {
        foreach (object? value in values)
        {
            string text = value switch
            {
                JsonNode node => Text(node),
                string s => s,
                null => "",
                _ => value.ToString() ?? ""
            };
            if (!string.IsNullOrWhiteSpace(text)) return text;
        }
        return "";
    }

    private static bool IsNumber(JsonNode? node)
        => TryNumber(node, out _);

    private static bool TryNumber(JsonNode? node, out double value)
    {
        value = 0;
        if (node is null) return false;
        if (node is JsonValue jsonValue)
        {
            if (jsonValue.TryGetValue(out int i)) { value = i; return true; }
            if (jsonValue.TryGetValue(out long l)) { value = l; return true; }
            if (jsonValue.TryGetValue(out double d)) { value = d; return true; }
            if (jsonValue.TryGetValue(out decimal dec)) { value = (double)dec; return true; }
            if (jsonValue.TryGetValue(out string? s))
                return double.TryParse((s ?? "").Trim().TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }
        return double.TryParse(Text(node).Trim().TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
    }

    private static double Number(JsonNode? node)
        => TryNumber(node, out double value) ? value : 0;

    private static JsonNode NormalizeNumber(double value)
        => Math.Abs(value % 1) < 0.0000001 ? JsonValue.Create((long)value)! : JsonValue.Create(Math.Round(value, 6))!;

    private static string Text(JsonNode? node)
    {
        if (node is null) return "";
        if (node is JsonValue value)
        {
            if (value.TryGetValue(out string? s)) return s ?? "";
            if (value.TryGetValue(out double d)) return d.ToString("0.####", CultureInfo.InvariantCulture);
            if (value.TryGetValue(out int i)) return i.ToString(CultureInfo.InvariantCulture);
            if (value.TryGetValue(out long l)) return l.ToString(CultureInfo.InvariantCulture);
            if (value.TryGetValue(out bool b)) return b ? "true" : "false";
        }
        return node.ToJsonString();
    }

    private static string Humanize(string key)
    {
        if (string.IsNullOrWhiteSpace(key)) return "";
        var sb = new StringBuilder();
        foreach (char ch in key)
        {
            if (sb.Length > 0 && char.IsUpper(ch) && !char.IsWhiteSpace(sb[^1])) sb.Append(' ');
            sb.Append(ch is '_' or '-' ? ' ' : ch);
        }
        return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(sb.ToString().Trim());
    }

    private static bool IsRateLikeKey(string key)
        => key.Contains("rate", StringComparison.OrdinalIgnoreCase)
           || key.Contains("percent", StringComparison.OrdinalIgnoreCase)
           || key.Contains("ratio", StringComparison.OrdinalIgnoreCase)
           || key.Contains("avg", StringComparison.OrdinalIgnoreCase)
           || key.Contains("average", StringComparison.OrdinalIgnoreCase);

    private static bool IsRowCountKey(string key)
        => key.Equals("rowCount", StringComparison.OrdinalIgnoreCase)
           || key.Equals("recordCount", StringComparison.OrdinalIgnoreCase)
           || key.Equals("sampleCount", StringComparison.OrdinalIgnoreCase);

    private static string EarliestNonEmpty(string a, string b)
    {
        if (string.IsNullOrWhiteSpace(a)) return b;
        if (string.IsNullOrWhiteSpace(b)) return a;
        return string.CompareOrdinal(a, b) <= 0 ? a : b;
    }

    private static string LatestNonEmpty(string a, string b)
    {
        if (string.IsNullOrWhiteSpace(a)) return b;
        if (string.IsNullOrWhiteSpace(b)) return a;
        return string.CompareOrdinal(a, b) >= 0 ? a : b;
    }

    private static string Enc(string value)
        => WebUtility.HtmlEncode(value ?? "");
}
