using DataMaker.Logger;
using DataMaker.R6.Grouping;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using static DataMaker.R6.CONSTANT;

namespace DataMaker.R6.Report
{
    /// <summary>
    /// Worst Reason Report 생성 클래스
    /// 대표 그룹의 상위 Reason별로 표시 (Reason 내에서 ProcessName+NGName 조합)
    /// </summary>
    public class clWorstReasonReportMakerVer1 : clBaseReportMaker
    {
        private readonly int referenceGroupIndex;
        private readonly int topReasonCount;
        private readonly int topDefectsPerReason;

        private sealed class ReferenceAggregates
        {
            public required List<AggregatedReasonData> Rows { get; init; }
            public required Dictionary<string, List<AggregatedReasonData>> RowsByReason { get; init; }
            public required Dictionary<string, HashSet<(string ProcessType, string ProcessName, string NGName)>> CurrentDefectsByReason { get; init; }
            public required Dictionary<string, Dictionary<(string ProcessType, string ProcessName, string NGName), DefectPeriodCache>> DefectCacheByReason { get; init; }
        }

        public clWorstReasonReportMakerVer1(string dbPath, List<clModelGroupData> groupedModels, int referenceGroupIndex, int topReasonCount = 20, int? topDefectsPerReason = null)
            : base(dbPath, groupedModels)
        {
            this.referenceGroupIndex = referenceGroupIndex;
            this.topReasonCount = topReasonCount;
            this.topDefectsPerReason = topDefectsPerReason ?? CONSTANT.OPTION.TopDefectsPerReason;
        }

        public clWorstReasonReportMakerVer1(string dbPath, List<clModelGroupData> groupedModels, List<(int GroupIndex, DataTable Data)> preloadedGroupDataList, int referenceGroupIndex, int topReasonCount = 20, int? topDefectsPerReason = null)
            : base(dbPath, groupedModels, preloadedGroupDataList)
        {
            this.referenceGroupIndex = referenceGroupIndex;
            this.topReasonCount = topReasonCount;
            this.topDefectsPerReason = topDefectsPerReason ?? CONSTANT.OPTION.TopDefectsPerReason;
        }

        public override Task<DataTable> CreateReport()
        {
            return Task.Run(() =>
            {
                clLogger.Log("=== Start Creating Worst Reason Report ===");
                clLogger.Log($"Reference Group Index: {referenceGroupIndex}");
                clLogger.Log($"Top Reason Count: {topReasonCount}");
                clLogger.Log($"Top Defects Per Reason: {topDefectsPerReason}");

                ReportProgress(0, 0, "Loading group data...");
                var groupDataList = LoadOrReuseGroupData();

                ReportProgress(100, 80, "Creating pivot table...");
                var result = CreatePivotTable(groupDataList);

                return result;
            });
        }

        private DataTable CreatePivotTable(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var result = new DataTable();

            ReportProgress(0, 80, "Analyzing period coverage...");

            cachedMonths = GetUniqueMonths(groupDataList);
            cachedWeeks = GetUniqueWeeks(groupDataList);
            cachedDates = GetUniqueDates(groupDataList);

            clLogger.Log($"Period coverage - Months: {string.Join(", ", cachedMonths)}, Weeks: {string.Join(", ", cachedWeeks)}, Dates: {cachedDates.Count} days");

            ReportProgress(10, 82, "Creating table structure...");

            result.Columns.Add("Reason", typeof(string));
            result.Columns.Add("Number", typeof(object));
            result.Columns.Add("ProcessType", typeof(string));
            result.Columns.Add("ProcessName", typeof(string));
            result.Columns.Add("NGName", typeof(string));
            result.Columns.Add("Unit", typeof(string));

            foreach (var date in cachedDates)
            {
                result.Columns.Add(date, typeof(object));
            }

            if (cachedDates.Count > 0 && cachedWeeks.Count > 0)
            {
                result.Columns.Add(SEPARATOR_1, typeof(object));
            }

            foreach (var week in cachedWeeks)
            {
                result.Columns.Add($"W{week}", typeof(object));
            }

            if (cachedWeeks.Count > 0 && cachedMonths.Count > 0)
            {
                result.Columns.Add(SEPARATOR_2, typeof(object));
            }

            foreach (var month in cachedMonths)
            {
                result.Columns.Add($"M{month}", typeof(object));
            }

            clLogger.Log($"Created columns - Months: {cachedMonths.Count}, Weeks: {cachedWeeks.Count}, Dates: {cachedDates.Count}");

            ReportProgress(30, 85, "Finding worst reasons...");
            FillWorstReasonDataRows(result, groupDataList);

            ReportProgress(100, 100, "Worst reason report completed!");

            if (result.Columns.Contains("Unit"))
            {
                result.Columns.Remove("Unit");
            }

            return result;
        }

        private void FillWorstReasonDataRows(DataTable result, List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var referenceGroupData = groupDataList.FirstOrDefault(g => g.GroupIndex == referenceGroupIndex);
            if (referenceGroupData.Data == null)
            {
                clLogger.LogWarning($"Reference group {referenceGroupIndex} not found!");
                return;
            }

            clLogger.Log($"Reference group {referenceGroupIndex} has {referenceGroupData.Data.Rows.Count} rows");
            clLogger.Log("Pre-aggregating reference group data for performance...");

            var referenceAggregates = BuildReferenceAggregates(referenceGroupData.Data);
            clLogger.Log($"Pre-aggregated to {referenceAggregates.Rows.Count} combinations");

            List<AggregatedReasonData> referencePeriodRows;
            string referencePeriodLabel;
            string currentPeriodLabel;

            if (CONSTANT.OPTION.RankingPeriodType == "Week")
            {
                int weekOffset = CONSTANT.OPTION.WorstRankingWeekOffset;
                if (weekOffset >= cachedWeeks.Count)
                {
                    clLogger.LogWarning($"Week offset {weekOffset} exceeds available weeks {cachedWeeks.Count}. Using oldest week.");
                    weekOffset = cachedWeeks.Count - 1;
                }

                long referenceWeek = cachedWeeks.Skip(weekOffset).First();
                long currentWeek = cachedWeeks.First();
                referencePeriodRows = referenceAggregates.Rows.Where(r => r.Week == referenceWeek).ToList();
                referencePeriodLabel = $"W{referenceWeek}";
                currentPeriodLabel = $"W{currentWeek}";
                clLogger.Log($"Filtering reference week (W{referenceWeek}, offset={weekOffset}): {referencePeriodRows.Count} combinations");
            }
            else
            {
                int monthOffset = CONSTANT.OPTION.WorstRankingMonthOffset;
                if (monthOffset >= cachedMonths.Count)
                {
                    clLogger.LogWarning($"Month offset {monthOffset} exceeds available months {cachedMonths.Count}. Using oldest month.");
                    monthOffset = cachedMonths.Count - 1;
                }

                long referenceMonth = cachedMonths.Skip(monthOffset).First();
                long currentMonth = cachedMonths.First();
                referencePeriodRows = referenceAggregates.Rows.Where(r => r.Month == referenceMonth).ToList();
                referencePeriodLabel = $"M{referenceMonth}";
                currentPeriodLabel = $"M{currentMonth}";
                clLogger.Log($"Filtering reference month (M{referenceMonth}, offset={monthOffset}): {referencePeriodRows.Count} combinations");
            }

            var topReasons = CalculateTopReasons(referencePeriodRows, referenceAggregates, currentPeriodLabel);

            clLogger.Log($"Found {topReasons.Count} top reasons");

            foreach (var reason in topReasons.Take(10))
            {
                clLogger.Log($"  {reason.Reason} | PPM={reason.TotalPPM}");
            }

            if (topReasons.Count > 10)
            {
                clLogger.Log($"  ... and {topReasons.Count - 10} more");
            }

            foreach (var reasonInfo in topReasons)
            {
                clLogger.Log($"  Reason '{reasonInfo.Reason}': {reasonInfo.Defects.Count} defect combinations (top {topDefectsPerReason}, from reference period {referencePeriodLabel})");
                AddReasonRows(result, reasonInfo, referenceAggregates);
            }
        }

        private ReferenceAggregates BuildReferenceAggregates(DataTable referenceGroupData)
        {
            var rows = referenceGroupData.AsEnumerable()
                .GroupBy(r => new
                {
                    Reason = r.Field<string>(CONSTANT.REASON.NEW),
                    ProcessType = r.Field<string>(CONSTANT.PROCESSTYPE.NEW),
                    ProcessName = r.Field<string>(CONSTANT.PROCESSNAME.NEW),
                    NGName = r.Field<string>(CONSTANT.NGNAME.NEW),
                    Date = r.Field<string>(CONSTANT.PRODUCT_DATE.NEW),
                    Week = r.Field<long>(CONSTANT.WEEK.NEW),
                    Month = r.Field<long>(CONSTANT.MONTH.NEW)
                })
                .Select(g => new AggregatedReasonData
                {
                    Reason = g.Key.Reason ?? string.Empty,
                    ProcessType = g.Key.ProcessType ?? string.Empty,
                    ProcessName = g.Key.ProcessName ?? string.Empty,
                    NGName = g.Key.NGName ?? string.Empty,
                    Date = g.Key.Date ?? string.Empty,
                    Week = g.Key.Week,
                    Month = g.Key.Month,
                    InputQty = g.Sum(r => r.Field<double>(CONSTANT.QTYINPUT.NEW)),
                    NGQty = g.Sum(r => r.Field<double>(CONSTANT.QTYNG.NEW))
                })
                .ToList();

            var rowsByReason = rows
                .GroupBy(r => r.Reason)
                .ToDictionary(g => g.Key, g => g.ToList());

            var currentDefectsByReason = BuildCurrentDefectsByReason(rowsByReason);
            var defectCacheByReason = BuildDefectCacheByReason(rowsByReason);

            return new ReferenceAggregates
            {
                Rows = rows,
                RowsByReason = rowsByReason,
                CurrentDefectsByReason = currentDefectsByReason,
                DefectCacheByReason = defectCacheByReason
            };
        }

        private Dictionary<string, HashSet<(string ProcessType, string ProcessName, string NGName)>> BuildCurrentDefectsByReason(
            Dictionary<string, List<AggregatedReasonData>> rowsByReason)
        {
            var result = new Dictionary<string, HashSet<(string ProcessType, string ProcessName, string NGName)>>();

            if (CONSTANT.OPTION.RankingPeriodType == "Week")
            {
                long currentWeek = cachedWeeks.First();
                foreach (var (reason, rows) in rowsByReason)
                {
                    result[reason] = rows
                        .Where(r => r.Week == currentWeek && r.NGQty > 0)
                        .Select(r => (r.ProcessType, r.ProcessName, r.NGName))
                        .ToHashSet();
                }
            }
            else
            {
                long currentMonth = cachedMonths.First();
                foreach (var (reason, rows) in rowsByReason)
                {
                    result[reason] = rows
                        .Where(r => r.Month == currentMonth && r.NGQty > 0)
                        .Select(r => (r.ProcessType, r.ProcessName, r.NGName))
                        .ToHashSet();
                }
            }

            return result;
        }

        private Dictionary<string, Dictionary<(string ProcessType, string ProcessName, string NGName), DefectPeriodCache>> BuildDefectCacheByReason(
            Dictionary<string, List<AggregatedReasonData>> rowsByReason)
        {
            var result = new Dictionary<string, Dictionary<(string ProcessType, string ProcessName, string NGName), DefectPeriodCache>>();

            foreach (var (reason, rows) in rowsByReason)
            {
                var defectCache = new Dictionary<(string ProcessType, string ProcessName, string NGName), DefectPeriodCache>();

                foreach (var row in rows)
                {
                    var defectKey = (row.ProcessType, row.ProcessName, row.NGName);
                    if (!defectCache.TryGetValue(defectKey, out var cache))
                    {
                        cache = new DefectPeriodCache();
                        defectCache[defectKey] = cache;
                    }

                    cache.ByDate[row.Date] = Sum(cache.ByDate, row.Date, row.InputQty, row.NGQty);
                    cache.ByWeek[row.Week] = Sum(cache.ByWeek, row.Week, row.InputQty, row.NGQty);
                    cache.ByMonth[row.Month] = Sum(cache.ByMonth, row.Month, row.InputQty, row.NGQty);
                }

                result[reason] = defectCache;
            }

            return result;
        }

        private static (double Input, double NG) Sum(Dictionary<string, (double Input, double NG)> source, string key, double inputQty, double ngQty)
        {
            source.TryGetValue(key, out var existing);
            return (existing.Input + inputQty, existing.NG + ngQty);
        }

        private static (double Input, double NG) Sum(Dictionary<long, (double Input, double NG)> source, long key, double inputQty, double ngQty)
        {
            source.TryGetValue(key, out var existing);
            return (existing.Input + inputQty, existing.NG + ngQty);
        }

        private List<ReasonInfo> CalculateTopReasons(
            List<AggregatedReasonData> referencePeriodRows,
            ReferenceAggregates referenceAggregates,
            string currentPeriodLabel)
        {
            var topReasons = new List<ReasonInfo>();

            foreach (var reasonGroup in referencePeriodRows.GroupBy(r => r.Reason))
            {
                string reason = reasonGroup.Key;

                referenceAggregates.CurrentDefectsByReason.TryGetValue(reason, out var currentDefects);
                currentDefects ??= new HashSet<(string ProcessType, string ProcessName, string NGName)>();

                clLogger.Log($"  Reason '{reason}': {currentDefects.Count} defects with NG in current period {currentPeriodLabel}");

                var defects = reasonGroup
                    .GroupBy(r => (r.ProcessType, r.ProcessName, r.NGName))
                    .Where(g => currentDefects.Contains(g.Key))
                    .Select(g =>
                    {
                        double totalInput = g.Sum(r => r.InputQty);
                        double totalNG = g.Sum(r => r.NGQty);
                        return new DefectInfo
                        {
                            ProcessType = g.Key.ProcessType,
                            ProcessName = g.Key.ProcessName,
                            NGName = g.Key.NGName,
                            PPM = totalInput > 0 ? Math.Round((totalNG / totalInput) * 1000000, 0) : 0
                        };
                    })
                    .OrderByDescending(d => d.PPM)
                    .Take(topDefectsPerReason)
                    .ToList();

                topReasons.Add(new ReasonInfo
                {
                    Reason = reason,
                    TotalPPM = defects.Sum(d => d.PPM),
                    Defects = defects
                });
            }

            return topReasons
                .Where(r => !string.IsNullOrWhiteSpace(r.Reason) && r.Reason != "Unknown")
                .OrderByDescending(r => r.TotalPPM)
                .Take(topReasonCount)
                .ToList();
        }

        private void AddReasonRows(DataTable result, ReasonInfo reasonInfo, ReferenceAggregates referenceAggregates)
        {
            var totalRow = result.NewRow();
            totalRow["Reason"] = reasonInfo.Reason;
            totalRow["Number"] = string.Empty;
            totalRow["ProcessType"] = string.Empty;
            totalRow["ProcessName"] = "Total";
            totalRow["NGName"] = string.Empty;
            totalRow["Unit"] = "PPM";

            FillTotalPeriods(totalRow, result, reasonInfo, referenceAggregates);
            result.Rows.Add(totalRow);

            int defectNumber = 1;
            foreach (var defect in reasonInfo.Defects)
            {
                var row = result.NewRow();
                row["Reason"] = reasonInfo.Reason;
                row["Number"] = defectNumber;
                row["ProcessType"] = defect.ProcessType;
                row["ProcessName"] = defect.ProcessName;
                row["NGName"] = defect.NGName;
                row["Unit"] = "PPM";

                FillDefectPeriods(row, result, reasonInfo.Reason, defect, referenceAggregates);
                result.Rows.Add(row);
                defectNumber++;
            }

            while (defectNumber <= topDefectsPerReason)
            {
                var emptyRow = result.NewRow();
                emptyRow["Reason"] = reasonInfo.Reason;
                emptyRow["Number"] = defectNumber;
                emptyRow["ProcessType"] = string.Empty;
                emptyRow["ProcessName"] = string.Empty;
                emptyRow["NGName"] = string.Empty;
                emptyRow["Unit"] = "PPM";

                foreach (var date in cachedDates)
                {
                    emptyRow[date] = DBNull.Value;
                }

                SetSeparatorValue(emptyRow, result);

                foreach (var week in cachedWeeks)
                {
                    emptyRow[$"W{week}"] = DBNull.Value;
                }

                foreach (var month in cachedMonths)
                {
                    emptyRow[$"M{month}"] = DBNull.Value;
                }

                result.Rows.Add(emptyRow);
                defectNumber++;
            }
        }

        private void FillTotalPeriods(DataRow row, DataTable result, ReasonInfo reasonInfo, ReferenceAggregates referenceAggregates)
        {
            foreach (var date in cachedDates)
            {
                row[date] = reasonInfo.Defects.Sum(defect => CalculatePeriodPpm(referenceAggregates, reasonInfo.Reason, defect, date));
            }

            SetSeparatorValue(row, result);

            foreach (var week in cachedWeeks)
            {
                row[$"W{week}"] = reasonInfo.Defects.Sum(defect => CalculatePeriodPpm(referenceAggregates, reasonInfo.Reason, defect, week));
            }

            foreach (var month in cachedMonths)
            {
                row[$"M{month}"] = reasonInfo.Defects.Sum(defect => CalculatePeriodPpm(referenceAggregates, reasonInfo.Reason, defect, month));
            }
        }

        private void FillDefectPeriods(DataRow row, DataTable result, string reason, DefectInfo defect, ReferenceAggregates referenceAggregates)
        {
            foreach (var date in cachedDates)
            {
                row[date] = CalculatePeriodPpm(referenceAggregates, reason, defect, date);
            }

            SetSeparatorValue(row, result);

            foreach (var week in cachedWeeks)
            {
                row[$"W{week}"] = CalculatePeriodPpm(referenceAggregates, reason, defect, week);
            }

            foreach (var month in cachedMonths)
            {
                row[$"M{month}"] = CalculatePeriodPpm(referenceAggregates, reason, defect, month);
            }
        }

        private double CalculatePeriodPpm(ReferenceAggregates referenceAggregates, string reason, DefectInfo defect, string date)
        {
            if (!TryGetDefectPeriodCache(referenceAggregates, reason, defect, out var cache))
            {
                return 0;
            }

            return cache.ByDate.TryGetValue(date, out var data) && data.Input > 0
                ? Math.Round((data.NG / data.Input) * 1000000, 0)
                : 0;
        }

        private double CalculatePeriodPpm(ReferenceAggregates referenceAggregates, string reason, DefectInfo defect, long week)
        {
            if (!TryGetDefectPeriodCache(referenceAggregates, reason, defect, out var cache))
            {
                return 0;
            }

            return cache.ByWeek.TryGetValue(week, out var data) && data.Input > 0
                ? Math.Round((data.NG / data.Input) * 1000000, 0)
                : 0;
        }

        private double CalculatePeriodPpm(ReferenceAggregates referenceAggregates, string reason, DefectInfo defect, int month)
        {
            if (!TryGetDefectPeriodCache(referenceAggregates, reason, defect, out var cache))
            {
                return 0;
            }

            return cache.ByMonth.TryGetValue(month, out var data) && data.Input > 0
                ? Math.Round((data.NG / data.Input) * 1000000, 0)
                : 0;
        }

        private bool TryGetDefectPeriodCache(ReferenceAggregates referenceAggregates, string reason, DefectInfo defect, out DefectPeriodCache cache)
        {
            cache = null;

            if (!referenceAggregates.DefectCacheByReason.TryGetValue(reason, out var reasonCache))
            {
                return false;
            }

            return reasonCache.TryGetValue((defect.ProcessType, defect.ProcessName, defect.NGName), out cache);
        }

        private sealed class AggregatedReasonData
        {
            public required string Reason { get; init; }
            public required string ProcessType { get; init; }
            public required string ProcessName { get; init; }
            public required string NGName { get; init; }
            public required string Date { get; init; }
            public long Week { get; init; }
            public long Month { get; init; }
            public double InputQty { get; init; }
            public double NGQty { get; init; }
        }

        private sealed class DefectPeriodCache
        {
            public Dictionary<string, (double Input, double NG)> ByDate { get; } = new Dictionary<string, (double, double)>();
            public Dictionary<long, (double Input, double NG)> ByWeek { get; } = new Dictionary<long, (double, double)>();
            public Dictionary<long, (double Input, double NG)> ByMonth { get; } = new Dictionary<long, (double, double)>();
        }

        private sealed class DefectInfo
        {
            public required string ProcessType { get; init; }
            public required string ProcessName { get; init; }
            public required string NGName { get; init; }
            public double PPM { get; init; }
        }

        private sealed class ReasonInfo
        {
            public required string Reason { get; init; }
            public double TotalPPM { get; init; }
            public required List<DefectInfo> Defects { get; init; }
        }
    }
}
