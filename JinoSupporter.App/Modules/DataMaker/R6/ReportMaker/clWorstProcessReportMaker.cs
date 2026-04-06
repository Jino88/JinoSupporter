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
    /// Worst Process Report 생성 클래스
    /// ProcessName 기준으로 지정된 그룹에서 불량률 상위 추출
    /// 각 ProcessName마다 대표 그룹 먼저 표시 후 다른 그룹들 표시
    /// </summary>
    public class clWorstProcessReportMaker : clBaseReportMaker
    {
        private readonly int referenceGroupIndex;
        private readonly int topProcessCount;

        private const string SEPARATOR_AFTER_PROCESSNAME = "BLANK_AFTER_PROCESS";
        private const string SEPARATOR_1 = "BLANK_1";
        private const string SEPARATOR_2 = "BLANK_2";

        private sealed class GroupAggregates
        {
            public required List<AggregatedData> Rows { get; init; }
            public required Dictionary<(string ProcessType, string ProcessName), ProcessPeriodCache> ByProcess { get; init; }
        }

        public clWorstProcessReportMaker(string dbPath, List<clModelGroupData> groupedModels,
            int referenceGroupIndex, int topProcessCount = 10)
            : base(dbPath, groupedModels)
        {
            this.referenceGroupIndex = referenceGroupIndex;
            this.topProcessCount = topProcessCount;
        }

        public clWorstProcessReportMaker(string dbPath, List<clModelGroupData> groupedModels,
            List<(int GroupIndex, DataTable Data)> preloadedGroupDataList,
            int referenceGroupIndex, int topProcessCount = 10)
            : base(dbPath, groupedModels, preloadedGroupDataList)
        {
            this.referenceGroupIndex = referenceGroupIndex;
            this.topProcessCount = topProcessCount;
        }

        public override Task<DataTable> CreateReport()
        {
            return Task.Run(() =>
            {
                clLogger.Log("=== Start Creating Worst Process Report (ProcessName based) ===");
                clLogger.Log($"Reference Group Index: {referenceGroupIndex}");
                clLogger.Log($"Top Process Count: {topProcessCount}");

                ReportProgress(0, 0, "Loading group data...");
                var groupDataList = LoadOrReuseGroupData();

                ReportProgress(100, 80, "Creating pivot table...");
                var result = CreatePivotTable(groupDataList);

                ReportProgress(100, 100, "Worst Process report completed!");
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
            CreateTableStructure(result);

            ReportProgress(30, 85, "Finding worst processes...");
            FillWorstProcessDataRows(result, groupDataList);

            return result;
        }

        private void CreateTableStructure(DataTable result)
        {
            result.Columns.Add("No", typeof(int));
            result.Columns.Add("ProcessType", typeof(string));
            result.Columns.Add("ProcessName", typeof(string));
            result.Columns.Add(SEPARATOR_AFTER_PROCESSNAME, typeof(object));
            result.Columns.Add("GroupName", typeof(string));

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
        }

        private void FillWorstProcessDataRows(DataTable result, List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            clLogger.Log("Pre-aggregating data for performance...");
            var aggregatesByGroup = BuildAggregates(groupDataList);

            if (!aggregatesByGroup.TryGetValue(referenceGroupIndex, out var referenceAggregates))
            {
                clLogger.LogWarning($"Reference group {referenceGroupIndex} not found!");
                return;
            }

            clLogger.Log($"Reference group {referenceGroupIndex} has {referenceAggregates.Rows.Count} aggregated combinations");

            List<AggregatedData> referenceData;
            if (CONSTANT.OPTION.RankingPeriodType == "Week")
            {
                int weekOffset = CONSTANT.OPTION.WorstRankingWeekOffset;
                if (weekOffset >= cachedWeeks.Count)
                {
                    clLogger.LogWarning($"Week offset {weekOffset} exceeds available weeks {cachedWeeks.Count}. Using oldest week.");
                    weekOffset = cachedWeeks.Count - 1;
                }

                long referenceWeek = cachedWeeks.Skip(weekOffset).First();
                referenceData = referenceAggregates.Rows.Where(r => r.Week == referenceWeek).ToList();
                clLogger.Log($"Filtering reference week (W{referenceWeek}, offset={weekOffset}): {referenceData.Count} combinations");
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
                referenceData = referenceAggregates.Rows.Where(r => r.Month == referenceMonth).ToList();
                clLogger.Log($"Filtering reference month (M{referenceMonth}, offset={monthOffset}): {referenceData.Count} combinations");
            }

            var topProcesses = CalculateProcessPPMs(referenceData)
                .OrderByDescending(p => p.TotalPPM)
                .Take(topProcessCount)
                .ToList();

            clLogger.Log($"Found {topProcesses.Count} top ProcessNames");

            int processNo = 1;
            foreach (var processInfo in topProcesses)
            {
                AddProcessRow(result, aggregatesByGroup, processNo, processInfo.ProcessType, processInfo.ProcessName, referenceGroupIndex);

                foreach (var group in groupedModels.Where(g => g.Index != referenceGroupIndex).OrderBy(g => g.Index))
                {
                    AddProcessRow(result, aggregatesByGroup, processNo, processInfo.ProcessType, processInfo.ProcessName, group.Index);
                }

                processNo++;
            }
        }

        private Dictionary<int, GroupAggregates> BuildAggregates(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var result = new Dictionary<int, GroupAggregates>();

            foreach (var (groupIndex, data) in groupDataList)
            {
                var rows = data.AsEnumerable()
                    .GroupBy(r => new
                    {
                        ProcessType = r.Field<string>(CONSTANT.PROCESSTYPE.NEW),
                        ProcessName = r.Field<string>(CONSTANT.PROCESSNAME.NEW),
                        NGName = r.Field<string>(CONSTANT.NGNAME.NEW),
                        Date = r.Field<string>(CONSTANT.PRODUCT_DATE.NEW),
                        Week = r.Field<long>(CONSTANT.WEEK.NEW),
                        Month = r.Field<long>(CONSTANT.MONTH.NEW)
                    })
                    .Select(g => new AggregatedData
                    {
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

                result[groupIndex] = new GroupAggregates
                {
                    Rows = rows,
                    ByProcess = BuildProcessPeriodCache(rows)
                };

                clLogger.Log($"  Group {groupIndex}: Pre-aggregated to {rows.Count} combinations");
            }

            return result;
        }

        private Dictionary<(string ProcessType, string ProcessName), ProcessPeriodCache> BuildProcessPeriodCache(List<AggregatedData> rows)
        {
            var result = new Dictionary<(string ProcessType, string ProcessName), ProcessPeriodCache>();

            foreach (var row in rows)
            {
                var processKey = (row.ProcessType, row.ProcessName);
                if (!result.TryGetValue(processKey, out var cache))
                {
                    cache = new ProcessPeriodCache();
                    result[processKey] = cache;
                }

                cache.ByDate[row.Date] = Sum(cache.ByDate, row.Date, row.NGName, row.InputQty, row.NGQty);
                cache.ByWeek[row.Week] = Sum(cache.ByWeek, row.Week, row.NGName, row.InputQty, row.NGQty);
                cache.ByMonth[row.Month] = Sum(cache.ByMonth, row.Month, row.NGName, row.InputQty, row.NGQty);
            }

            return result;
        }

        private static Dictionary<string, (double InputQty, double NGQty)> Sum(
            Dictionary<string, Dictionary<string, (double InputQty, double NGQty)>> source,
            string periodKey,
            string ngName,
            double inputQty,
            double ngQty)
        {
            if (!source.TryGetValue(periodKey, out var byNg))
            {
                byNg = new Dictionary<string, (double InputQty, double NGQty)>();
                source[periodKey] = byNg;
            }

            byNg.TryGetValue(ngName, out var existing);
            byNg[ngName] = (existing.InputQty + inputQty, existing.NGQty + ngQty);
            return byNg;
        }

        private static Dictionary<string, (double InputQty, double NGQty)> Sum(
            Dictionary<long, Dictionary<string, (double InputQty, double NGQty)>> source,
            long periodKey,
            string ngName,
            double inputQty,
            double ngQty)
        {
            if (!source.TryGetValue(periodKey, out var byNg))
            {
                byNg = new Dictionary<string, (double InputQty, double NGQty)>();
                source[periodKey] = byNg;
            }

            byNg.TryGetValue(ngName, out var existing);
            byNg[ngName] = (existing.InputQty + inputQty, existing.NGQty + ngQty);
            return byNg;
        }

        private List<(string ProcessType, string ProcessName, double TotalPPM)> CalculateProcessPPMs(List<AggregatedData> data)
        {
            return data
                .GroupBy(r => (r.ProcessType, r.ProcessName))
                .Select(g =>
                {
                    double totalPpm = g
                        .GroupBy(r => r.NGName)
                        .Sum(ngGroup =>
                        {
                            double inputQty = ngGroup.Sum(r => r.InputQty);
                            double ngQty = ngGroup.Sum(r => r.NGQty);
                            return inputQty > 0 ? Math.Round((ngQty / inputQty) * 1_000_000, 0) : 0;
                        });

                    return (g.Key.ProcessType, g.Key.ProcessName, totalPpm);
                })
                .ToList();
        }

        private void AddProcessRow(
            DataTable result,
            Dictionary<int, GroupAggregates> aggregatesByGroup,
            int no,
            string processType,
            string processName,
            int groupIndex)
        {
            if (!aggregatesByGroup.TryGetValue(groupIndex, out var aggregates))
            {
                return;
            }

            if (!aggregates.ByProcess.TryGetValue((processType, processName), out var cache))
            {
                cache = new ProcessPeriodCache();
            }

            string groupName = groupedModels.First(g => g.Index == groupIndex).GroupName ?? string.Empty;

            var row = result.NewRow();
            row["No"] = no;
            row["ProcessType"] = processType;
            row["ProcessName"] = processName;
            row[SEPARATOR_AFTER_PROCESSNAME] = string.Empty;
            row["GroupName"] = groupName;

            foreach (var date in cachedDates)
            {
                double ppm = CalculatePeriodPpm(cache.ByDate, date);
                row[date] = ppm > 0 ? ppm : DBNull.Value;
            }

            if (cachedDates.Count > 0 && cachedWeeks.Count > 0)
            {
                row[SEPARATOR_1] = string.Empty;
            }

            foreach (var week in cachedWeeks)
            {
                row[$"W{week}"] = CalculatePeriodPpm(cache.ByWeek, week);
            }

            if (cachedWeeks.Count > 0 && cachedMonths.Count > 0)
            {
                row[SEPARATOR_2] = string.Empty;
            }

            foreach (var month in cachedMonths)
            {
                row[$"M{month}"] = CalculatePeriodPpm(cache.ByMonth, month);
            }

            result.Rows.Add(row);
        }

        private static double CalculatePeriodPpm(Dictionary<string, Dictionary<string, (double InputQty, double NGQty)>> source, string periodKey)
        {
            if (!source.TryGetValue(periodKey, out var byNg))
            {
                return 0;
            }

            return byNg.Values.Sum(v => v.InputQty > 0 ? Math.Round((v.NGQty / v.InputQty) * 1_000_000, 0) : 0);
        }

        private static double CalculatePeriodPpm(Dictionary<long, Dictionary<string, (double InputQty, double NGQty)>> source, long periodKey)
        {
            if (!source.TryGetValue(periodKey, out var byNg))
            {
                return 0;
            }

            return byNg.Values.Sum(v => v.InputQty > 0 ? Math.Round((v.NGQty / v.InputQty) * 1_000_000, 0) : 0);
        }

        private sealed class AggregatedData
        {
            public required string ProcessType { get; init; }
            public required string ProcessName { get; init; }
            public required string NGName { get; init; }
            public required string Date { get; init; }
            public long Week { get; init; }
            public long Month { get; init; }
            public double InputQty { get; init; }
            public double NGQty { get; init; }
        }

        private sealed class ProcessPeriodCache
        {
            public Dictionary<string, Dictionary<string, (double InputQty, double NGQty)>> ByDate { get; } = new Dictionary<string, Dictionary<string, (double, double)>>();
            public Dictionary<long, Dictionary<string, (double InputQty, double NGQty)>> ByWeek { get; } = new Dictionary<long, Dictionary<string, (double, double)>>();
            public Dictionary<long, Dictionary<string, (double InputQty, double NGQty)>> ByMonth { get; } = new Dictionary<long, Dictionary<string, (double, double)>>();
        }
    }
}
