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
    /// 모든 JSON의 모든 그룹에 대한 상세 리포트
    /// Input Qty, NG Qty, NG Rate 표시
    /// </summary>
    public class clAllGroupsDetailReportMaker : clBaseReportMaker
    {
        public clAllGroupsDetailReportMaker(string dbPath, List<clModelGroupData> groupedModels)
            : base(dbPath, groupedModels)
        {
        }

        public clAllGroupsDetailReportMaker(string dbPath, List<clModelGroupData> groupedModels, List<(int GroupIndex, DataTable Data)> preloadedGroupDataList)
            : base(dbPath, groupedModels, preloadedGroupDataList)
        {
        }

        public override Task<DataTable> CreateReport()
        {
            return Task.Run(() =>
            {
                clLogger.Log("=== Start Creating All Groups Detail Report ===");

                // 1. 그룹 데이터 로드
                ReportProgress(0, 0, "Loading group data...");
                var groupDataList = LoadOrReuseGroupData();

                // 2. 리포트 테이블 생성
                ReportProgress(100, 80, "Creating report table...");
                var result = CreateReportTable(groupDataList);

                return result;
            });
        }

        private DataTable CreateReportTable(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var result = new DataTable();

            ReportProgress(0, 80, "Analyzing period coverage...");

            // 전체 기간 수집
            cachedMonths = GetUniqueMonths(groupDataList);
            cachedWeeks = GetUniqueWeeks(groupDataList);
            cachedDates = GetUniqueDates(groupDataList);

            clLogger.Log($"Months: {cachedMonths.Count}, Weeks: {cachedWeeks.Count}, Dates: {cachedDates.Count}");

            ReportProgress(10, 82, "Creating table structure...");

            // 1~5열: 공백
            result.Columns.Add("Col1", typeof(string));
            result.Columns.Add("Col2", typeof(string));
            result.Columns.Add("Col3", typeof(string));
            result.Columns.Add("Col4", typeof(string));
            result.Columns.Add("Col5", typeof(string));

            // 6~11열: 기본 정보
            result.Columns.Add("Combination", typeof(string));  // ModelName_ProcessName_NgName
            result.Columns.Add("ModelName", typeof(string));
            result.Columns.Add("ProcessType", typeof(string));
            result.Columns.Add("Reason", typeof(string));
            result.Columns.Add("ProcessName", typeof(string));
            result.Columns.Add("NGName", typeof(string));

            // 12열부터: 일별 (내림차순) - Input Qty, NG Qty, NG Rate
            foreach (var date in cachedDates)
            {
                result.Columns.Add($"{date}_InputQty", typeof(double));
                result.Columns.Add($"{date}_NGQty", typeof(double));
                result.Columns.Add($"{date}_NGRate", typeof(double));
            }

            if (cachedDates.Count > 0 && cachedWeeks.Count > 0)
            {
                result.Columns.Add(SEPARATOR_1, typeof(object));
            }

            // 주별 (내림차순)
            foreach (var week in cachedWeeks)
            {
                result.Columns.Add($"W{week}_InputQty", typeof(double));
                result.Columns.Add($"W{week}_NGQty", typeof(double));
                result.Columns.Add($"W{week}_NGRate", typeof(double));
            }

            if (cachedWeeks.Count > 0 && cachedMonths.Count > 0)
            {
                result.Columns.Add(SEPARATOR_2, typeof(object));
            }

            // 월별 (내림차순)
            foreach (var month in cachedMonths)
            {
                result.Columns.Add($"M{month}_InputQty", typeof(double));
                result.Columns.Add($"M{month}_NGQty", typeof(double));
                result.Columns.Add($"M{month}_NGRate", typeof(double));
            }

            clLogger.Log($"Created {result.Columns.Count} columns");

            ReportProgress(30, 85, "Filling data rows...");

            // 데이터 행 채우기
            FillDataRows(result, groupDataList);

            ReportProgress(100, 100, "All groups detail report completed!");

            return result;
        }

        private void FillDataRows(DataTable result, List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            int totalRows = 0;
            int currentRow = 0;

            // 전체 행 수 계산
            foreach (var (groupIndex, data) in groupDataList)
            {
                var combinations = data.AsEnumerable()
                    .Select(r => new
                    {
                        ProcessType = r.Field<string>(CONSTANT.PROCESSTYPE.NEW),
                        ProcessName = r.Field<string>(CONSTANT.PROCESSNAME.NEW),
                        NGName = r.Field<string>(CONSTANT.NGNAME.NEW)
                    })
                    .Distinct()
                    .Count();
                totalRows += combinations;
            }

            clLogger.Log($"Total combinations across all groups: {totalRows}");

            // 각 그룹별 데이터 행 생성
            foreach (var (groupIndex, data) in groupDataList)
            {
                var group = groupedModels.First(g => g.Index == groupIndex);
                string modelName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m).Take(2));

                // 데이터를 미리 그룹화 (성능 최적화)
                var preAggregated = data.AsEnumerable()
                    .GroupBy(r => new
                    {
                        ProcessType = r.Field<string>(CONSTANT.PROCESSTYPE.NEW),
                        ProcessName = r.Field<string>(CONSTANT.PROCESSNAME.NEW),
                        NGName = r.Field<string>(CONSTANT.NGNAME.NEW),
                        Reason = r.Field<string>(CONSTANT.REASON.NEW),
                        Date = r.Field<string>(CONSTANT.PRODUCT_DATE.NEW),
                        Week = r.Field<long>(CONSTANT.WEEK.NEW),
                        Month = r.Field<long>(CONSTANT.MONTH.NEW)
                    })
                    .Select(g => new
                    {
                        g.Key.ProcessType,
                        g.Key.ProcessName,
                        g.Key.NGName,
                        g.Key.Reason,
                        g.Key.Date,
                        g.Key.Week,
                        g.Key.Month,
                        InputQty = g.Sum(r => r.Field<double>(CONSTANT.QTYINPUT.NEW)),
                        NGQty = g.Sum(r => r.Field<double>(CONSTANT.QTYNG.NEW))
                    })
                    .ToList();

                // 조합별로 그룹화
                var combinations = preAggregated
                    .GroupBy(r => new
                    {
                        r.ProcessType,
                        r.ProcessName,
                        r.NGName,
                        r.Reason
                    })
                    .OrderBy(g => g.Key.ProcessType)
                    .ThenBy(g => g.Key.ProcessName)
                    .ThenBy(g => g.Key.NGName)
                    .ToList();

                clLogger.Log($"Group {groupIndex} ({modelName}): {combinations.Count} combinations");

                foreach (var combo in combinations)
                {
                    currentRow++;

                    var row = result.NewRow();

                    // 1~5열: 공백
                    row["Col1"] = "";
                    row["Col2"] = "";
                    row["Col3"] = "";
                    row["Col4"] = "";
                    row["Col5"] = "";

                    // 6~11열: 기본 정보
                    row["Combination"] = $"{modelName}_{combo.Key.ProcessName}_{combo.Key.NGName}";
                    row["ModelName"] = modelName;
                    row["ProcessType"] = combo.Key.ProcessType ?? "";
                    row["Reason"] = combo.Key.Reason ?? "";
                    row["ProcessName"] = combo.Key.ProcessName ?? "";
                    row["NGName"] = combo.Key.NGName ?? "";

                    // 일별 데이터 (미리 집계된 데이터 사용)
                    foreach (var date in cachedDates)
                    {
                        var dateData = combo.Where(r => r.Date == date);
                        double inputQty = dateData.Sum(r => r.InputQty);
                        double ngQty = dateData.Sum(r => r.NGQty);
                        double ngRate = inputQty > 0 ? Math.Round((ngQty / inputQty) * 1000000, 0) : 0;

                        row[$"{date}_InputQty"] = inputQty;
                        row[$"{date}_NGQty"] = ngQty;
                        row[$"{date}_NGRate"] = ngRate;
                    }

                    if (result.Columns.Contains(SEPARATOR_1))
                    {
                        row[SEPARATOR_1] = DBNull.Value;
                    }

                    // 주별 데이터
                    foreach (var week in cachedWeeks)
                    {
                        var weekData = combo.Where(r => r.Week == week);
                        double inputQty = weekData.Sum(r => r.InputQty);
                        double ngQty = weekData.Sum(r => r.NGQty);
                        double ngRate = inputQty > 0 ? Math.Round((ngQty / inputQty) * 1000000, 0) : 0;

                        row[$"W{week}_InputQty"] = inputQty;
                        row[$"W{week}_NGQty"] = ngQty;
                        row[$"W{week}_NGRate"] = ngRate;
                    }

                    if (result.Columns.Contains(SEPARATOR_2))
                    {
                        row[SEPARATOR_2] = DBNull.Value;
                    }

                    // 월별 데이터
                    foreach (var month in cachedMonths)
                    {
                        var monthData = combo.Where(r => r.Month == month);
                        double inputQty = monthData.Sum(r => r.InputQty);
                        double ngQty = monthData.Sum(r => r.NGQty);
                        double ngRate = inputQty > 0 ? Math.Round((ngQty / inputQty) * 1000000, 0) : 0;

                        row[$"M{month}_InputQty"] = inputQty;
                        row[$"M{month}_NGQty"] = ngQty;
                        row[$"M{month}_NGRate"] = ngRate;
                    }

                    result.Rows.Add(row);

                    // 진행률 업데이트
                    if (currentRow % 100 == 0 || currentRow == totalRows)
                    {
                        int progress = (int)((currentRow / (double)totalRows) * 100);
                        int totalProgress = 85 + (int)(progress * 0.15);
                        ReportProgress(progress, totalProgress, $"Processing row {currentRow}/{totalRows}...");
                    }
                }
            }

            clLogger.Log($"Filled {result.Rows.Count} data rows");
        }
    }
}
