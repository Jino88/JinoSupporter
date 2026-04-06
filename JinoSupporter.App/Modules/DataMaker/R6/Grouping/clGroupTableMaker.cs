using DataMaker.Logger;
using DataMaker.R6.SQLProcess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DataMaker.R6.CONSTANT;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DataMaker.R6.Grouping
{
    public class clGroupDebugSourceRow
    {
        public string ModelName { get; set; } = "";
        public string ProcessName { get; set; } = "";
        public string NGName { get; set; } = "";
        public string NGCode { get; set; } = "";
        public string ProductDate { get; set; } = "";
        public double InputQty { get; set; }
        public double NGQty { get; set; }
    }

    public class clGroupDebugRateResult
    {
        public string GroupName { get; set; } = "";
        public string GroupTableName { get; set; } = "";
        public string ProcessName { get; set; } = "";
        public string NGName { get; set; } = "";
        public string ProductDate { get; set; } = "";
        public int MatchedRowCount { get; set; }
        public double InputQty { get; set; }
        public double NGQty { get; set; }
        public double NGRate { get; set; } // NG / Input
        public double NGRatePercent => NGRate * 100.0;

        public List<clGroupDebugSourceRow> SourceRows { get; set; } = new List<clGroupDebugSourceRow>();
        public string SourceRowsMessage { get; set; } = "";
        public double SourceInputQty { get; set; }
        public double SourceNGQty { get; set; }
        public double SourceNGRate => SourceInputQty > 0 ? SourceNGQty / SourceInputQty : 0;
        public int SourceRowCount => SourceRows.Count;
    }

    public class clGroupTableMaker
    {
        private const string GroupModelTableSuffix = "__MODELS";
        private clSQLFileIO sql;

        public clGroupTableMaker(string dbPath)
        {
            sql = new clSQLFileIO(dbPath);
        }

        public void CreateGroupTables(List<clModelGroupData> groupedModels)
        {
            if (groupedModels == null || groupedModels.Count == 0)
            {
                return;
            }

            DataTable procData = sql.LoadTable(OPTION_TABLE_NAME.PROC) ?? new DataTable();
            ILookup<string, DataRow>? procRowsByModel = null;

            if (procData != null &&
                procData.Columns.Contains(CONSTANT.ModelName_WithLineShift.NEW))
            {
                procRowsByModel = procData.AsEnumerable()
                    .Where(row =>
                    {
                        string model = row[CONSTANT.ModelName_WithLineShift.NEW]?.ToString();
                        return !string.IsNullOrWhiteSpace(model);
                    })
                    .ToLookup(
                        row => row[CONSTANT.ModelName_WithLineShift.NEW]?.ToString() ?? string.Empty,
                        StringComparer.Ordinal);
            }

            foreach (var group in groupedModels)
            {
                // CONSTANT의 공통 메서드를 사용하여 테이블 이름 생성
                string tableName = CONSTANT.GetGroupTableName(group);
                CreateGroupTable(tableName, group.ModelList, procData, procRowsByModel);
            }
        }

        private void CreateGroupTable(
            string tableName,
            List<string> models,
            DataTable procData,
            ILookup<string, DataRow>? procRowsByModel)
        {
            clLogger.Log($"Creating {tableName} with {models.Count} models");

            // 1. 기존 테이블 삭제
            if (sql.IsTableExist(tableName))
            {
                sql.DropTable(tableName);
            }

            // 2. ProcTable에서 해당 모델들 데이터 로드
            var procDataForGroup = LoadProcTableData(models, procData, procRowsByModel);

            // 3. 그룹핑 및 집계
            var grouped = GroupAndAggregate(procDataForGroup, models);

            // 4. 테이블 생성 및 저장
            SaveGroupTable(tableName, grouped);
            SaveGroupModelTable(tableName, models);

            clLogger.Log($"{tableName} created with {grouped.Rows.Count} rows");
        }

        private DataTable LoadProcTableData(
            List<string> models,
            DataTable sourceProcData,
            ILookup<string, DataRow>? procRowsByModel = null)
        {
            if (sourceProcData == null || sourceProcData.Columns.Count == 0)
            {
                return new DataTable();
            }

            if (!sourceProcData.Columns.Contains(CONSTANT.ModelName_WithLineShift.NEW))
            {
                clLogger.LogWarning($"PROC table is missing column: {CONSTANT.ModelName_WithLineShift.NEW}");
                return sourceProcData.Clone();
            }

            var modelSet = new HashSet<string>(
                (models ?? new List<string>())
                    .Where(m => !string.IsNullOrWhiteSpace(m))
            );

            if (modelSet.Count == 0)
            {
                return sourceProcData.Clone();
            }

            List<DataRow> filteredRows;
            if (procRowsByModel != null)
            {
                filteredRows = modelSet
                    .SelectMany(model => procRowsByModel[model])
                    .ToList();
            }
            else
            {
                filteredRows = sourceProcData.AsEnumerable()
                    .Where(row =>
                    {
                        string model = row[CONSTANT.ModelName_WithLineShift.NEW]?.ToString();
                        return !string.IsNullOrWhiteSpace(model) && modelSet.Contains(model);
                    })
                    .ToList();
            }

            if (filteredRows.Count == 0)
            {
                return sourceProcData.Clone();
            }

            var filtered = filteredRows.CopyToDataTable();

            return filtered;
        }

        private DataTable GroupAndAggregate(DataTable data, List<string> models)
        {
            var result = new DataTable();
            result.Columns.Add(CONSTANT.ModelName_WithLineShift.NEW, typeof(string));
            result.Columns.Add(CONSTANT.PROCESSNAME.NEW, typeof(string));
            result.Columns.Add(CONSTANT.PROCESSTYPE.NEW, typeof(string));
            result.Columns.Add(CONSTANT.NGNAME.NEW, typeof(string));
            result.Columns.Add(CONSTANT.NGCODE.NEW, typeof(string));
            result.Columns.Add(CONSTANT.REASON.NEW, typeof(string));
            result.Columns.Add(CONSTANT.PRODUCT_DATE.NEW, typeof(string));
            result.Columns.Add(CONSTANT.MONTH.NEW, typeof(long));
            result.Columns.Add(CONSTANT.WEEK.NEW, typeof(long));
            result.Columns.Add(CONSTANT.QTYINPUT.NEW, typeof(double));
            result.Columns.Add(CONSTANT.QTYNG.NEW, typeof(double));

            // 모델명 합치기 (정렬 후)
            var sortedModels = models.OrderBy(m => m).ToList();
            string combinedModelName = string.Join("_", sortedModels);

            var grouped = data.AsEnumerable()
                .GroupBy(row => new
                {
                    ProcessName = row.Field<string>(CONSTANT.PROCESSNAME.NEW),
                    NGName = row.Field<string>(CONSTANT.NGNAME.NEW),
                    NGCode = row.Field<string>(CONSTANT.NGCODE.NEW),
                    ProductDate = row.Field<string>(CONSTANT.PRODUCT_DATE.NEW)
                })
                .Select(g => new
                {
                    g.Key.ProcessName,
                    ProcessType = g.First().Field<string>(CONSTANT.PROCESSTYPE.NEW),
                    g.Key.NGName,
                    g.Key.NGCode,
                    Reason = g.First().Field<string>(CONSTANT.REASON.NEW),
                    g.Key.ProductDate,
                    Month = g.First().Field<long>(CONSTANT.MONTH.NEW),
                    Week = g.First().Field<long>(CONSTANT.WEEK.NEW),
                    TotalInput = g.Sum(r => Convert.ToDouble(r[CONSTANT.QTYINPUT.NEW])),
                    TotalNG = g.Sum(r => Convert.ToDouble(r[CONSTANT.QTYNG.NEW]))
                });

            foreach (var item in grouped)
            {
                result.Rows.Add(
                    combinedModelName,
                    item.ProcessName,
                    item.ProcessType,
                    item.NGName,
                    item.NGCode,
                    item.Reason,
                    item.ProductDate,
                    item.Month, // 추가
                    item.Week, // 추가
                    item.TotalInput,
                    item.TotalNG
                );
            }

            return result;
        }

        private void SaveGroupTable(string tableName, DataTable data)
        {
            sql.WriteTable(tableName, data);
        }

        private void SaveGroupModelTable(string groupTableName, List<string> models)
        {
            string modelTableName = GetGroupModelTableName(groupTableName);

            if (sql.IsTableExist(modelTableName))
            {
                sql.DropTable(modelTableName);
            }

            var modelTable = new DataTable();
            modelTable.Columns.Add(CONSTANT.ModelName_WithLineShift.NEW, typeof(string));

            var normalizedModels = (models ?? new List<string>())
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Select(m => m.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(m => m)
                .ToList();

            foreach (var model in normalizedModels)
            {
                modelTable.Rows.Add(model);
            }

            sql.WriteTable(modelTableName, modelTable);
        }

        private static string GetGroupModelTableName(string groupTableName)
        {
            return $"{groupTableName}{GroupModelTableSuffix}";
        }

        public clGroupDebugRateResult GetDebugRate(string groupName, string processName, string ngName, string productDate)
        {
            if (string.IsNullOrWhiteSpace(groupName))
            {
                throw new ArgumentException("groupName is required.", nameof(groupName));
            }

            var groupInfo = new clModelGroupData
            {
                GroupName = groupName,
                ModelList = new List<string>()
            };

            string tableName = CONSTANT.GetGroupTableName(groupInfo);
            var result = new clGroupDebugRateResult
            {
                GroupName = groupName?.Trim() ?? "",
                GroupTableName = tableName,
                ProcessName = processName?.Trim() ?? "",
                NGName = ngName?.Trim() ?? "",
                ProductDate = productDate?.Trim() ?? ""
            };

            if (!sql.IsTableExist(tableName))
            {
                clLogger.LogWarning($"Debug rate lookup failed: table not found ({tableName})");
                return result;
            }

            DataTable data = sql.LoadTable(tableName);
            if (data == null || data.Rows.Count == 0)
            {
                clLogger.LogWarning($"Debug rate lookup: table has no rows ({tableName})");
                return result;
            }

            string normalizedProcess = NormalizeForCompare(processName);
            string normalizedNg = NormalizeForCompare(ngName);
            string targetDate = (productDate ?? "").Trim();

            var matchedRows = data.AsEnumerable()
                .Where(row =>
                    NormalizeForCompare(row.Field<string>(CONSTANT.PROCESSNAME.NEW)) == normalizedProcess &&
                    NormalizeForCompare(row.Field<string>(CONSTANT.NGNAME.NEW)) == normalizedNg &&
                    IsDateMatched(row.Field<string>(CONSTANT.PRODUCT_DATE.NEW), targetDate))
                .ToList();

            result.MatchedRowCount = matchedRows.Count;
            result.InputQty = matchedRows.Sum(r => ToDoubleSafe(r[CONSTANT.QTYINPUT.NEW]));
            result.NGQty = matchedRows.Sum(r => ToDoubleSafe(r[CONSTANT.QTYNG.NEW]));
            result.NGRate = result.InputQty > 0 ? result.NGQty / result.InputQty : 0;

            var matchedNgCodes = matchedRows
                .Select(r => r[CONSTANT.NGCODE.NEW]?.ToString()?.Trim() ?? "")
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            result.SourceRows = LoadSourceRowsFromProc(
                tableName,
                result.GroupName,
                matchedRows,
                normalizedProcess,
                normalizedNg,
                targetDate,
                matchedNgCodes,
                out string sourceMessage);
            result.SourceRowsMessage = sourceMessage;
            result.SourceInputQty = result.SourceRows.Sum(r => r.InputQty);
            result.SourceNGQty = result.SourceRows.Sum(r => r.NGQty);

            clLogger.Log(
                $"DebugRate [{tableName}] " +
                $"Process='{result.ProcessName}', NG='{result.NGName}', Date='{result.ProductDate}', " +
                $"Rows={result.MatchedRowCount}, Input={result.InputQty}, NG={result.NGQty}, NGRATE={result.NGRate}, " +
                $"SourceRows={result.SourceRowCount}");

            return result;
        }

        private List<clGroupDebugSourceRow> LoadSourceRowsFromProc(
            string groupTableName,
            string groupName,
            List<DataRow> matchedGroupRows,
            string normalizedProcess,
            string normalizedNg,
            string targetDate,
            HashSet<string> matchedNgCodes,
            out string message)
        {
            string modelTableName = GetGroupModelTableName(groupTableName);
            message = "";

            DataTable procData = sql.LoadTable(OPTION_TABLE_NAME.PROC);
            if (procData == null || procData.Rows.Count == 0)
            {
                message = $"Source row lookup skipped: {OPTION_TABLE_NAME.PROC} is empty.";
                return new List<clGroupDebugSourceRow>();
            }

            string[] requiredColumns =
            {
                CONSTANT.ModelName_WithLineShift.NEW,
                CONSTANT.PROCESSNAME.NEW,
                CONSTANT.NGNAME.NEW,
                CONSTANT.PRODUCT_DATE.NEW,
                CONSTANT.QTYINPUT.NEW,
                CONSTANT.QTYNG.NEW
            };

            foreach (string column in requiredColumns)
            {
                if (!procData.Columns.Contains(column))
                {
                    message = $"Source row lookup skipped: {OPTION_TABLE_NAME.PROC} missing column ({column}).";
                    return new List<clGroupDebugSourceRow>();
                }
            }

            bool hasNgCodeColumn = procData.Columns.Contains(CONSTANT.NGCODE.NEW);
            string modelSetResolveMessage = "";
            bool useModelFilter = false;
            var modelSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (sql.IsTableExist(modelTableName))
            {
                DataTable modelTable = sql.LoadTable(modelTableName);
                if (modelTable != null && modelTable.Columns.Contains(CONSTANT.ModelName_WithLineShift.NEW))
                {
                    modelSet = modelTable.AsEnumerable()
                        .Select(r => r[CONSTANT.ModelName_WithLineShift.NEW]?.ToString()?.Trim() ?? "")
                        .Where(m => !string.IsNullOrWhiteSpace(m))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    if (modelSet.Count > 0)
                    {
                        useModelFilter = true;
                        modelSetResolveMessage = $"Source rows resolved from {OPTION_TABLE_NAME.PROC} using model table {modelTableName}.";
                    }
                    else
                    {
                        modelSetResolveMessage = $"Model table exists but empty ({modelTableName}).";
                    }
                }
                else
                {
                    modelSetResolveMessage = $"Model table exists but invalid ({modelTableName}).";
                }
            }
            else
            {
                modelSetResolveMessage = $"Model table not found ({modelTableName}).";
            }

            if (!useModelFilter)
            {
                var procModelUniverse = procData.AsEnumerable()
                    .Select(r => r[CONSTANT.ModelName_WithLineShift.NEW]?.ToString()?.Trim() ?? "")
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                var fallbackSet = BuildFallbackModelSet(procModelUniverse, matchedGroupRows, groupName);
                if (fallbackSet.Count > 0)
                {
                    modelSet = fallbackSet;
                    useModelFilter = true;
                    modelSetResolveMessage += $" Fallback model inference applied ({modelSet.Count} models).";
                }
                else
                {
                    modelSetResolveMessage += " Fallback model inference failed; source rows are not filtered by model.";
                }
            }

            var sourceRows = procData.AsEnumerable()
                .Where(row =>
                {
                    string modelName = row[CONSTANT.ModelName_WithLineShift.NEW]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrWhiteSpace(modelName))
                    {
                        return false;
                    }

                    if (useModelFilter && !modelSet.Contains(modelName))
                    {
                        return false;
                    }

                    if (NormalizeForCompare(row.Field<string>(CONSTANT.PROCESSNAME.NEW)) != normalizedProcess)
                    {
                        return false;
                    }

                    if (NormalizeForCompare(row.Field<string>(CONSTANT.NGNAME.NEW)) != normalizedNg)
                    {
                        return false;
                    }

                    if (!IsDateMatched(row.Field<string>(CONSTANT.PRODUCT_DATE.NEW), targetDate))
                    {
                        return false;
                    }

                    if (hasNgCodeColumn && matchedNgCodes.Count > 0)
                    {
                        string ngCode = row[CONSTANT.NGCODE.NEW]?.ToString()?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(ngCode) && !matchedNgCodes.Contains(ngCode))
                        {
                            return false;
                        }
                    }

                    return true;
                })
                .Select(row => new clGroupDebugSourceRow
                {
                    ModelName = row[CONSTANT.ModelName_WithLineShift.NEW]?.ToString()?.Trim() ?? "",
                    ProcessName = row[CONSTANT.PROCESSNAME.NEW]?.ToString()?.Trim() ?? "",
                    NGName = row[CONSTANT.NGNAME.NEW]?.ToString()?.Trim() ?? "",
                    NGCode = hasNgCodeColumn ? (row[CONSTANT.NGCODE.NEW]?.ToString()?.Trim() ?? "") : "",
                    ProductDate = row[CONSTANT.PRODUCT_DATE.NEW]?.ToString()?.Trim() ?? "",
                    InputQty = ToDoubleSafe(row[CONSTANT.QTYINPUT.NEW]),
                    NGQty = ToDoubleSafe(row[CONSTANT.QTYNG.NEW])
                })
                .OrderBy(r => r.ModelName)
                .ThenBy(r => r.NGCode)
                .ThenBy(r => r.ProductDate)
                .ToList();

            message = modelSetResolveMessage;
            return sourceRows;
        }

        private static HashSet<string> BuildFallbackModelSet(
            HashSet<string> procModelUniverse,
            List<DataRow> matchedGroupRows,
            string groupName)
        {
            var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (procModelUniverse == null || procModelUniverse.Count == 0)
            {
                return result;
            }

            var candidates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (matchedGroupRows != null)
            {
                foreach (var row in matchedGroupRows)
                {
                    string value = row[CONSTANT.ModelName_WithLineShift.NEW]?.ToString()?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        candidates.Add(value);
                    }
                }
            }

            if (!string.IsNullOrWhiteSpace(groupName))
            {
                candidates.Add(groupName.Trim());
            }

            foreach (string candidate in candidates)
            {
                if (procModelUniverse.Contains(candidate))
                {
                    result.Add(candidate);
                }
            }

            if (result.Count > 0)
            {
                return result;
            }

            foreach (string candidate in candidates)
            {
                var tokens = candidate.Split('_')
                    .Select(t => t.Trim())
                    .Where(t => !string.IsNullOrWhiteSpace(t));

                foreach (string token in tokens)
                {
                    if (procModelUniverse.Contains(token))
                    {
                        result.Add(token);
                    }
                }
            }

            return result;
        }

        private static string NormalizeForCompare(string? value)
        {
            return CONSTANT.Normalize(value ?? "").Trim();
        }

        private static bool IsDateMatched(string? rowDate, string? targetDate)
        {
            string left = (rowDate ?? "").Trim();
            string right = (targetDate ?? "").Trim();

            if (left.Equals(right, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (DateTime.TryParse(left, out DateTime leftDate) && DateTime.TryParse(right, out DateTime rightDate))
            {
                return leftDate.Date == rightDate.Date;
            }

            return false;
        }

        private static double ToDoubleSafe(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return 0;
            }

            if (value is double d)
            {
                return d;
            }

            if (double.TryParse(value.ToString(), out double parsed))
            {
                return parsed;
            }

            return 0;
        }

        public void Dispose()
        {
            sql?.Dispose();
        }
    }
}
