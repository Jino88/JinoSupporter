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
    /// Worst Report 생성 클래스 Ver2
    /// NGName 기준으로 지정된 그룹에서 불량률 상위 항목을 추출한다.
    /// </summary>
    public class clWorstReportMakerVer2 : clBaseReportMaker
    {
        private readonly int referenceGroupIndex;
        private readonly int topNGCount;

        private sealed class GroupAggregates
        {
            public required List<AggregatedData> Rows { get; init; }
            public required Dictionary<(string ProcessType, string ProcessName, string NGName, string Date), (double InputQty, double NGQty)> ByDate { get; init; }
            public required Dictionary<(string ProcessType, string ProcessName, string NGName, long Week), (double InputQty, double NGQty)> ByWeek { get; init; }
            public required Dictionary<(string ProcessType, string ProcessName, string NGName, long Month), (double InputQty, double NGQty)> ByMonth { get; init; }
        }

        public clWorstReportMakerVer2(string dbPath, List<clModelGroupData> groupedModels,
            int referenceGroupIndex, int topNGCount = 10)
            : base(dbPath, groupedModels)
        {
            this.referenceGroupIndex = referenceGroupIndex;
            this.topNGCount = topNGCount;
        }

        public clWorstReportMakerVer2(string dbPath, List<clModelGroupData> groupedModels,
            List<(int GroupIndex, DataTable Data)> preloadedGroupDataList,
            int referenceGroupIndex, int topNGCount = 10)
            : base(dbPath, groupedModels, preloadedGroupDataList)
        {
            this.referenceGroupIndex = referenceGroupIndex;
            this.topNGCount = topNGCount;
        }

        public override Task<DataTable> CreateReport()
        {
            return Task.Run(() =>
            {
                clLogger.Log("=== Start Creating Worst Report Ver2 (NGName based) ===");
                clLogger.Log($"Reference Group Index: {referenceGroupIndex}");
                clLogger.Log($"Top NG Count: {topNGCount}");

                ReportProgress(0, 0, "Loading group data...");
                var groupDataList = LoadOrReuseGroupData();

                ReportProgress(100, 80, "Creating pivot table...");
                var result = CreatePivotTable(groupDataList);

                ReportProgress(100, 100, "Worst report completed!");
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

            ReportProgress(30, 85, "Finding worst NGNames...");
            FillWorstNGDataRows(result, groupDataList);

            return result;
        }

        private void CreateTableStructure(DataTable result)
        {
            result.Columns.Add("No", typeof(int));
            result.Columns.Add("ProcessType", typeof(string));
            result.Columns.Add("ProcessName", typeof(string));
            result.Columns.Add("NGName", typeof(string));
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

        private void FillWorstNGDataRows(DataTable result, List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            clLogger.Log("Pre-aggregating data for performance...");
            var aggregatesByGroup = BuildAggregates(groupDataList);

            if (!aggregatesByGroup.TryGetValue(referenceGroupIndex, out var referenceGroupAggregates))
            {
                clLogger.LogWarning($"Reference group {referenceGroupIndex} not found!");
                return;
            }

            clLogger.Log($"Reference group {referenceGroupIndex} has {referenceGroupAggregates.Rows.Count} aggregated combinations");

            List<AggregatedData> referenceData;
            string periodType = CONSTANT.OPTION.RankingPeriodType;

            if (periodType == "Week")
            {
                int weekOffset = CONSTANT.OPTION.WorstRankingWeekOffset;
                if (weekOffset >= cachedWeeks.Count)
                {
                    clLogger.LogWarning($"Week offset {weekOffset} exceeds available weeks {cachedWeeks.Count}. Using oldest week.");
                    weekOffset = cachedWeeks.Count - 1;
                }

                long referenceWeek = cachedWeeks.Skip(weekOffset).First();
                referenceData = referenceGroupAggregates.Rows.Where(r => r.Week == referenceWeek).ToList();
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
                referenceData = referenceGroupAggregates.Rows.Where(r => r.Month == referenceMonth).ToList();
                clLogger.Log($"Filtering reference month (M{referenceMonth}, offset={monthOffset}): {referenceData.Count} combinations");
            }

            var topNGNames = CalculateNGPPMs(referenceData)
                .OrderByDescending(ng => ng.TotalPPM)
                .Take(topNGCount)
                .ToList();

            clLogger.Log($"Found {topNGNames.Count} top NG entries");

            int ngNo = 1;
            foreach (var ngInfo in topNGNames)
            {
                AddNGDataRow(result, aggregatesByGroup, ngNo, ngInfo, referenceGroupIndex);

                foreach (var group in groupedModels.Where(g => g.Index != referenceGroupIndex))
                {
                    AddNGDataRow(result, aggregatesByGroup, ngNo, ngInfo, group.Index);
                }

                ngNo++;
            }
        }

        private Dictionary<int, GroupAggregates> BuildAggregates(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var result = new Dictionary<int, GroupAggregates>();

            foreach (var (groupIndex, data) in groupDataList)
            {
                var aggregated = data.AsEnumerable()
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
                        ProcessType = g.Key.ProcessType,
                        ProcessName = g.Key.ProcessName,
                        NGName = g.Key.NGName,
                        Date = g.Key.Date,
                        Week = g.Key.Week,
                        Month = g.Key.Month,
                        InputQty = g.Sum(r => r.Field<double>(CONSTANT.QTYINPUT.NEW)),
                        NGQty = g.Sum(r => r.Field<double>(CONSTANT.QTYNG.NEW))
                    })
                    .ToList();

                result[groupIndex] = new GroupAggregates
                {
                    Rows = aggregated,
                    ByDate = BuildDateLookup(aggregated),
                    ByWeek = BuildWeekLookup(aggregated),
                    ByMonth = BuildMonthLookup(aggregated)
                };

                clLogger.Log($"  Group {groupIndex}: Pre-aggregated to {aggregated.Count} combinations");
            }

            return result;
        }

        private Dictionary<(string ProcessType, string ProcessName, string NGName, string Date), (double InputQty, double NGQty)> BuildDateLookup(
            List<AggregatedData> aggregated)
        {
            var lookup = new Dictionary<(string ProcessType, string ProcessName, string NGName, string Date), (double InputQty, double NGQty)>();

            foreach (var row in aggregated)
            {
                var key = (row.ProcessType ?? string.Empty, row.ProcessName ?? string.Empty, row.NGName ?? string.Empty, row.Date ?? string.Empty);
                lookup.TryGetValue(key, out var existing);
                lookup[key] = (existing.InputQty + row.InputQty, existing.NGQty + row.NGQty);
            }

            return lookup;
        }

        private Dictionary<(string ProcessType, string ProcessName, string NGName, long Week), (double InputQty, double NGQty)> BuildWeekLookup(
            List<AggregatedData> aggregated)
        {
            var lookup = new Dictionary<(string ProcessType, string ProcessName, string NGName, long Week), (double InputQty, double NGQty)>();

            foreach (var row in aggregated)
            {
                var key = (row.ProcessType ?? string.Empty, row.ProcessName ?? string.Empty, row.NGName ?? string.Empty, row.Week);
                lookup.TryGetValue(key, out var existing);
                lookup[key] = (existing.InputQty + row.InputQty, existing.NGQty + row.NGQty);
            }

            return lookup;
        }

        private Dictionary<(string ProcessType, string ProcessName, string NGName, long Month), (double InputQty, double NGQty)> BuildMonthLookup(
            List<AggregatedData> aggregated)
        {
            var lookup = new Dictionary<(string ProcessType, string ProcessName, string NGName, long Month), (double InputQty, double NGQty)>();

            foreach (var row in aggregated)
            {
                var key = (row.ProcessType ?? string.Empty, row.ProcessName ?? string.Empty, row.NGName ?? string.Empty, row.Month);
                lookup.TryGetValue(key, out var existing);
                lookup[key] = (existing.InputQty + row.InputQty, existing.NGQty + row.NGQty);
            }

            return lookup;
        }

        private List<(string NGName, string ProcessType, string ProcessName, double TotalPPM)> CalculateNGPPMs(List<AggregatedData> aggregatedData)
        {
            return aggregatedData
                .GroupBy(r => new { r.NGName, r.ProcessType, r.ProcessName })
                .Select(g =>
                {
                    double totalInput = g.Sum(r => r.InputQty);
                    double totalNG = g.Sum(r => r.NGQty);
                    double ppm = totalInput > 0 ? Math.Round((totalNG / totalInput) * 1000000, 0) : 0;

                    return (
                        NGName: g.Key.NGName ?? "Unknown",
                        ProcessType: g.Key.ProcessType ?? string.Empty,
                        ProcessName: g.Key.ProcessName ?? string.Empty,
                        TotalPPM: ppm);
                })
                .ToList();
        }

        private void AddNGDataRow(
            DataTable result,
            Dictionary<int, GroupAggregates> aggregatesByGroup,
            int ngNo,
            (string NGName, string ProcessType, string ProcessName, double TotalPPM) ngInfo,
            int groupIndex)
        {
            if (!aggregatesByGroup.TryGetValue(groupIndex, out var aggregates))
            {
                return;
            }

            var group = groupedModels.First(g => g.Index == groupIndex);
            string groupName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m).Take(2));

            var row = result.NewRow();
            row["No"] = ngNo;
            row["ProcessType"] = ngInfo.ProcessType;
            row["ProcessName"] = ngInfo.ProcessName;
            row["NGName"] = ngInfo.NGName;
            row["GroupName"] = groupName;

            foreach (var date in cachedDates)
            {
                aggregates.ByDate.TryGetValue((ngInfo.ProcessType, ngInfo.ProcessName, ngInfo.NGName, date), out var dateData);
                row[date] = dateData.InputQty > 0
                    ? Math.Round((dateData.NGQty / dateData.InputQty) * 1000000, 0)
                    : 0;
            }

            SetSeparatorValue(row, result);

            foreach (var week in cachedWeeks)
            {
                aggregates.ByWeek.TryGetValue((ngInfo.ProcessType, ngInfo.ProcessName, ngInfo.NGName, week), out var weekData);
                row[$"W{week}"] = weekData.InputQty > 0
                    ? Math.Round((weekData.NGQty / weekData.InputQty) * 1000000, 0)
                    : 0;
            }

            if (result.Columns.Contains(SEPARATOR_2))
            {
                row[SEPARATOR_2] = DBNull.Value;
            }

            foreach (var month in cachedMonths)
            {
                aggregates.ByMonth.TryGetValue((ngInfo.ProcessType, ngInfo.ProcessName, ngInfo.NGName, month), out var monthData);
                row[$"M{month}"] = monthData.InputQty > 0
                    ? Math.Round((monthData.NGQty / monthData.InputQty) * 1000000, 0)
                    : 0;
            }

            result.Rows.Add(row);
        }

        private sealed class AggregatedData
        {
            public string? ProcessType { get; set; }
            public string? ProcessName { get; set; }
            public string? NGName { get; set; }
            public string? Date { get; set; }
            public long Week { get; set; }
            public long Month { get; set; }
            public double InputQty { get; set; }
            public double NGQty { get; set; }
        }
    }
}
