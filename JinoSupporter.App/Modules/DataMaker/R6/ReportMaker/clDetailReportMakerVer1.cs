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
    /// Detail Report 생성 클래스
    /// 모든 ProcessType/ProcessName/NGName 포함 + Summary 행
    /// </summary>
    public class clDetailReportMakerVer1 : clBaseReportMaker
    {
        public clDetailReportMakerVer1(string dbPath, List<clModelGroupData> groupedModels)
            : base(dbPath, groupedModels)
        {
        }

        public clDetailReportMakerVer1(string dbPath, List<clModelGroupData> groupedModels, List<(int GroupIndex, DataTable Data)> preloadedGroupDataList)
            : base(dbPath, groupedModels, preloadedGroupDataList)
        {
        }

        public override Task<DataTable> CreateReport()
        {
            return Task.Run(() =>
            {
                clLogger.Log($"=== Start Creating Detail Report ===");
                clLogger.Log($"Groups: {groupedModels.Count}");

                // 1. 그룹 데이터 로드 (0% → 80%)
                ReportProgress(0, 0, "Loading group data...");

                var groupDataList = LoadOrReuseGroupData();

                // 2. Pivot 테이블 생성 (80% → 100%)
                ReportProgress(100, 80, "Creating pivot table...");
                var result = CreatePivotTable(groupDataList);

                return result;
            });
        }

        private DataTable CreatePivotTable(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var result = new DataTable();

            ReportProgress(0, 80, "Analyzing period coverage...");

            // 전체 기간 수집
            cachedMonths = GetUniqueMonths(groupDataList);
            cachedWeeks = GetUniqueWeeks(groupDataList);
            cachedDates = GetUniqueDates(groupDataList);
            cachedColumnNames = new Dictionary<(string, int), string>();

            clLogger.Log($"Period coverage - Months: {string.Join(", ", cachedMonths)}, Weeks: {string.Join(", ", cachedWeeks)}, Dates: {cachedDates.Count} days");

            ReportProgress(10, 82, "Creating table structure...");

            // 기본 컬럼 생성
            result.Columns.Add("Lv.2", typeof(string));

            // Lv.2 다음에 빈 컬럼 추가
            result.Columns.Add(SEPARATOR_AFTER_LV2, typeof(object));

            result.Columns.Add("ProcessName", typeof(string));
            result.Columns.Add("Lv.3", typeof(string));
            result.Columns.Add("Unit", typeof(string));

            // Unit 다음에 빈 컬럼 추가
            result.Columns.Add(SEPARATOR_AFTER_UNIT, typeof(object));

            var columnNames = new HashSet<string>();

            // 1. 일자별 컬럼
            foreach (var date in cachedDates)
            {
                foreach (var group in groupedModels)
                {
                    string groupName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m));
                    string colName = $"{date}_{groupName}";
                    string uniqueColName = GetUniqueColumnName(colName, columnNames);
                    columnNames.Add(uniqueColName);
                    result.Columns.Add(uniqueColName, typeof(object));
                    cachedColumnNames[(date, group.Index)] = uniqueColName;
                }
            }

            clLogger.Log($"Created {cachedDates.Count * groupedModels.Count} date columns for {groupedModels.Count} groups");

            // 빈 컬럼 1 (공백으로 표시)
            if (cachedDates.Count > 0 && cachedWeeks.Count > 0)
            {
                result.Columns.Add(SEPARATOR_1, typeof(object));
            }

            // 2. 주별 컬럼
            foreach (var week in cachedWeeks)
            {
                foreach (var group in groupedModels)
                {
                    string groupName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m));
                    string colName = $"W{week}_{groupName}";
                    string uniqueColName = GetUniqueColumnName(colName, columnNames);
                    columnNames.Add(uniqueColName);
                    result.Columns.Add(uniqueColName, typeof(object));
                    cachedColumnNames[($"W{week}", group.Index)] = uniqueColName;
                }
            }

            clLogger.Log($"Created {cachedWeeks.Count * groupedModels.Count} week columns for {groupedModels.Count} groups");

            // 빈 컬럼 2 (공백으로 표시)
            if (cachedWeeks.Count > 0 && cachedMonths.Count > 0)
            {
                result.Columns.Add(SEPARATOR_2, typeof(object));
            }

            // 3. 월별 컬럼
            foreach (var month in cachedMonths)
            {
                foreach (var group in groupedModels)
                {
                    string groupName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m));
                    string colName = $"M{month}_{groupName}";
                    string uniqueColName = GetUniqueColumnName(colName, columnNames);
                    columnNames.Add(uniqueColName);
                    result.Columns.Add(uniqueColName, typeof(object));
                    cachedColumnNames[($"M{month}", group.Index)] = uniqueColName;
                }
            }

            clLogger.Log($"Created {cachedMonths.Count * groupedModels.Count} month columns for {groupedModels.Count} groups");

            ReportProgress(30, 85, "Adding header rows...");

            // Row 0: 기간 헤더
            AddPeriodHeaderRow(result);

            // Row 1: 그룹 헤더
            AddGroupHeaderRow(result);

            ReportProgress(40, 88, "Filling data rows...");

            // Row 2+: 데이터 행
            var processTypePPMs = FillDataRows(result, groupDataList);

            ReportProgress(80, 95, "Inserting summary rows...");

            // Row 2부터 집계 행 삽입 (TOTAL + 모든 ProcessType)
            int summaryRowCount = InsertSummaryRows(result, processTypePPMs);

            ReportProgress(90, 98, "Finalizing table...");

            // Summary rows 다음에 완전히 빈 행 1개 삽입
            var emptyRow = result.NewRow();

            // 완전히 빈 행
            emptyRow["Lv.2"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
            {
                emptyRow[SEPARATOR_AFTER_LV2] = DBNull.Value;
            }

            emptyRow["ProcessName"] = "";
            emptyRow["Lv.3"] = "";
            emptyRow["Unit"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
            {
                emptyRow[SEPARATOR_AFTER_UNIT] = DBNull.Value;
            }

            foreach (var kvp in cachedColumnNames)
            {
                emptyRow[kvp.Value] = DBNull.Value;
            }
            if (result.Columns.Contains(SEPARATOR_1))
            {
                emptyRow[SEPARATOR_1] = DBNull.Value;
            }
            if (result.Columns.Contains(SEPARATOR_2))
            {
                emptyRow[SEPARATOR_2] = DBNull.Value;
            }

            result.Rows.InsertAt(emptyRow, 2 + summaryRowCount);


            // Row 11에 레이블 행 추가 (모델, Process Name, NG NAME)
            var groupHeaderRow = result.NewRow();
            groupHeaderRow["Lv.2"] = "모델";

            if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
            {
                groupHeaderRow[SEPARATOR_AFTER_LV2] = DBNull.Value;
            }

            groupHeaderRow["ProcessName"] = "Process Name";
            groupHeaderRow["Lv.3"] = "NG NAME";
            groupHeaderRow["Unit"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
            {
                groupHeaderRow[SEPARATOR_AFTER_UNIT] = DBNull.Value;
            }

            // 데이터 컬럼들에 Row 1 (PeriodHeader) 날짜 값 복사
            foreach (var date in cachedDates)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[(date, group.Index)];
                    groupHeaderRow[colName] = result.Rows[0][colName];
                }
            }
            if (result.Columns.Contains(SEPARATOR_1))
            {
                groupHeaderRow[SEPARATOR_1] = result.Rows[0][SEPARATOR_1];
            }
            foreach (var week in cachedWeeks)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"W{week}", group.Index)];
                    groupHeaderRow[colName] = result.Rows[0][colName];
                }
            }
            if (result.Columns.Contains(SEPARATOR_2))
            {
                groupHeaderRow[SEPARATOR_2] = result.Rows[0][SEPARATOR_2];
            }
            foreach (var month in cachedMonths)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"M{month}", group.Index)];
                    groupHeaderRow[colName] = result.Rows[0][colName];
                }
            }

            result.Rows.InsertAt(groupHeaderRow, 2 + summaryRowCount + 1);

            // Row 12에 모델명 행 추가
            var modelNameRow = result.NewRow();
            modelNameRow["Lv.2"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
            {
                modelNameRow[SEPARATOR_AFTER_LV2] = DBNull.Value;
            }

            modelNameRow["ProcessName"] = "";
            modelNameRow["Lv.3"] = "";
            modelNameRow["Unit"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
            {
                modelNameRow[SEPARATOR_AFTER_UNIT] = DBNull.Value;
            }

            // 데이터 컬럼들에 Row 2 (GroupHeader) 모델명 값 복사
            foreach (var date in cachedDates)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[(date, group.Index)];
                    modelNameRow[colName] = result.Rows[1][colName];
                }
            }
            if (result.Columns.Contains(SEPARATOR_1))
            {
                modelNameRow[SEPARATOR_1] = result.Rows[1][SEPARATOR_1];
            }
            foreach (var week in cachedWeeks)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"W{week}", group.Index)];
                    modelNameRow[colName] = result.Rows[1][colName];
                }
            }
            if (result.Columns.Contains(SEPARATOR_2))
            {
                modelNameRow[SEPARATOR_2] = result.Rows[1][SEPARATOR_2];
            }
            foreach (var month in cachedMonths)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"M{month}", group.Index)];
                    modelNameRow[colName] = result.Rows[1][colName];
                }
            }

            result.Rows.InsertAt(modelNameRow, 2 + summaryRowCount + 2);

            ReportProgress(100, 100, "Detail report completed!");

            if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
            {
                result.Columns.Remove(SEPARATOR_AFTER_UNIT);
            }

            return result;
        }

        /// <summary>
        /// Summary 행 삽입 (TOTAL + 실제 존재하는 모든 ProcessType)
        /// </summary>
        /// <returns>삽입한 행 개수</returns>
        private int InsertSummaryRows(DataTable result, Dictionary<(string period, int groupIndex, string processType), double> processTypePPMs)
        {
            // 실제 데이터에 존재하는 모든 ProcessType 수집
            var actualProcessTypes = processTypePPMs.Keys
                .Select(k => k.processType)
                .Distinct()
                .OrderBy(pt => {
                    switch (pt)
                    {
                        case "SUB": return 1;
                        case "MAIN": return 2;
                        case "FUNCTION": return 3;
                        case "VISUAL": return 4;
                        default: return pt.StartsWith("SUB") ? 1 : 5; // SUB2, SUB3 등도 SUB 다음에 정렬
                    }
                })
                .ThenBy(pt => pt) // 같은 우선순위 내에서는 알파벳순
                .ToList();

            // TOTAL 행을 맨 앞에 추가
            var processTypeOrder = new List<string> { "TOTAL" };
            processTypeOrder.AddRange(actualProcessTypes);

            int insertIndex = 2; // Row 0: 기간, Row 1: 그룹명, Row 2부터 집계

            bool isFirstRow = true;

            foreach (var processType in processTypeOrder)
            {
                var summaryRow = result.NewRow();
                summaryRow["Lv.2"] = "";

                // Lv.2 다음 빈 컬럼
                if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
                {
                    summaryRow[SEPARATOR_AFTER_LV2] = DBNull.Value;
                }

                summaryRow["ProcessName"] = isFirstRow ? "NG PPM" : "";
                summaryRow["Lv.3"] = processType;
                summaryRow["Unit"] = "";  // Unit 값 제거

                // Unit 다음 빈 컬럼
                if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
                {
                    summaryRow[SEPARATOR_AFTER_UNIT] = DBNull.Value;
                }

                isFirstRow = false;

                // 일자별
                foreach (var date in cachedDates)
                {
                    foreach (var group in groupedModels)
                    {
                        string colName = cachedColumnNames[(date, group.Index)];
                        double totalPPM = 0;

                        if (processType == "TOTAL")
                        {
                            // 모든 실제 ProcessType의 합계
                            foreach (var pt in actualProcessTypes)
                            {
                                var key = (date, group.Index, pt);
                                if (processTypePPMs.ContainsKey(key))
                                    totalPPM += processTypePPMs[key];
                            }
                        }
                        else
                        {
                            var key = (date, group.Index, processType);
                            if (processTypePPMs.ContainsKey(key))
                                totalPPM = processTypePPMs[key];
                        }

                        summaryRow[colName] = totalPPM;
                    }
                }

                // 빈 컬럼 1
                if (result.Columns.Contains(SEPARATOR_1))
                {
                    summaryRow[SEPARATOR_1] = DBNull.Value;
                }

                // 주별
                foreach (var week in cachedWeeks)
                {
                    foreach (var group in groupedModels)
                    {
                        string colName = cachedColumnNames[($"W{week}", group.Index)];
                        string periodKey = $"W{week}";
                        double totalPPM = 0;

                        if (processType == "TOTAL")
                        {
                            // 모든 실제 ProcessType의 합계
                            foreach (var pt in actualProcessTypes)
                            {
                                var key = (periodKey, group.Index, pt);
                                if (processTypePPMs.ContainsKey(key))
                                    totalPPM += processTypePPMs[key];
                            }
                        }
                        else
                        {
                            var key = (periodKey, group.Index, processType);
                            if (processTypePPMs.ContainsKey(key))
                                totalPPM = processTypePPMs[key];
                        }

                        summaryRow[colName] = totalPPM;
                    }
                }

                // 빈 컬럼 2
                if (result.Columns.Contains(SEPARATOR_2))
                {
                    summaryRow[SEPARATOR_2] = DBNull.Value;
                }

                // 월별
                foreach (var month in cachedMonths)
                {
                    foreach (var group in groupedModels)
                    {
                        string colName = cachedColumnNames[($"M{month}", group.Index)];
                        string periodKey = $"M{month}";
                        double totalPPM = 0;

                        if (processType == "TOTAL")
                        {
                            // 모든 실제 ProcessType의 합계
                            foreach (var pt in actualProcessTypes)
                            {
                                var key = (periodKey, group.Index, pt);
                                if (processTypePPMs.ContainsKey(key))
                                    totalPPM += processTypePPMs[key];
                            }
                        }
                        else
                        {
                            var key = (periodKey, group.Index, processType);
                            if (processTypePPMs.ContainsKey(key))
                                totalPPM = processTypePPMs[key];
                        }

                        summaryRow[colName] = totalPPM;
                    }
                }

                result.Rows.InsertAt(summaryRow, insertIndex);
                insertIndex++;
            }

            return processTypeOrder.Count;
        }

        /// <summary>
        /// 모든 불량 데이터 행 채우기
        /// </summary>
        private Dictionary<(string period, int groupIndex, string processType), double> FillDataRows(DataTable result, List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            // 성능 최적화: 각 그룹별 데이터를 미리 집계하고 딕셔너리로 인덱싱
            clLogger.Log("Pre-aggregating data for performance...");
            var dateDataByGroup = new Dictionary<int, Dictionary<(string, string, string, string), (double InputQty, double NGQty)>>();
            var weekDataByGroup = new Dictionary<int, Dictionary<(string, string, string, long), (double InputQty, double NGQty)>>();
            var monthDataByGroup = new Dictionary<int, Dictionary<(string, string, string, long), (double InputQty, double NGQty)>>();
            var missingGroupWarnings = new HashSet<int>();

            foreach (var (groupIndex, data) in groupDataList)
            {
                // Date lookup: (ProcessType, ProcessName, NGName, Date) -> (InputQty, NGQty)
                var dateDict = data.AsEnumerable()
                    .GroupBy(r => (
                        ProcessType: r.Field<string>(CONSTANT.PROCESSTYPE.NEW),
                        ProcessName: r.Field<string>(CONSTANT.PROCESSNAME.NEW),
                        NGName: r.Field<string>(CONSTANT.NGNAME.NEW),
                        Date: r.Field<string>(CONSTANT.PRODUCT_DATE.NEW)
                    ))
                    .ToDictionary(
                        g => g.Key,
                        g => (
                            InputQty: g.Sum(r => r.Field<double>(CONSTANT.QTYINPUT.NEW)),
                            NGQty: g.Sum(r => r.Field<double>(CONSTANT.QTYNG.NEW))
                        )
                    );

                // Week lookup: (ProcessType, ProcessName, NGName, Week) -> (InputQty, NGQty)
                var weekDict = data.AsEnumerable()
                    .GroupBy(r => (
                        ProcessType: r.Field<string>(CONSTANT.PROCESSTYPE.NEW),
                        ProcessName: r.Field<string>(CONSTANT.PROCESSNAME.NEW),
                        NGName: r.Field<string>(CONSTANT.NGNAME.NEW),
                        Week: r.Field<long>(CONSTANT.WEEK.NEW)
                    ))
                    .ToDictionary(
                        g => g.Key,
                        g => (
                            InputQty: g.Sum(r => r.Field<double>(CONSTANT.QTYINPUT.NEW)),
                            NGQty: g.Sum(r => r.Field<double>(CONSTANT.QTYNG.NEW))
                        )
                    );

                // Month lookup: (ProcessType, ProcessName, NGName, Month) -> (InputQty, NGQty)
                var monthDict = data.AsEnumerable()
                    .GroupBy(r => (
                        ProcessType: r.Field<string>(CONSTANT.PROCESSTYPE.NEW),
                        ProcessName: r.Field<string>(CONSTANT.PROCESSNAME.NEW),
                        NGName: r.Field<string>(CONSTANT.NGNAME.NEW),
                        Month: r.Field<long>(CONSTANT.MONTH.NEW)
                    ))
                    .ToDictionary(
                        g => g.Key,
                        g => (
                            InputQty: g.Sum(r => r.Field<double>(CONSTANT.QTYINPUT.NEW)),
                            NGQty: g.Sum(r => r.Field<double>(CONSTANT.QTYNG.NEW))
                        )
                    );

                dateDataByGroup[groupIndex] = dateDict;
                weekDataByGroup[groupIndex] = weekDict;
                monthDataByGroup[groupIndex] = monthDict;
                clLogger.Log($"  Group {groupIndex}: Indexed {dateDict.Count} date combinations, {weekDict.Count} week combinations, {monthDict.Count} month combinations");
            }

            var processTypePPMs = new Dictionary<(string period, int groupIndex, string processType), double>();
            var processTypes = GetUniqueProcessTypes(groupDataList);

            var sortedProcessTypes = processTypes.OrderBy(pt => {
                switch (pt)
                {
                    case "SUB": return 1;
                    case "MAIN": return 2;
                    case "FUNCTION": return 3;
                    case "VISUAL": return 4;
                    default: return pt.StartsWith("SUB") ? 1 : 5; // SUB4, SUB5 등도 SUB 그룹으로
                }
            })
            .ThenBy(pt => pt) // 같은 우선순위 내에서는 알파벳순
            .ToList();

            // 전체 ProcessName 수 계산 (프로그레스용)
            int totalProcessNames = 0;
            foreach (var processType in sortedProcessTypes)
            {
                var processNames = GetProcessNames(groupDataList, processType);
                totalProcessNames += processNames.Count;
            }

            int currentProcessNameIndex = 0;

            foreach (var processType in sortedProcessTypes)
            {
                var processNames = GetProcessNames(groupDataList, processType);

                foreach (var processName in processNames)
                {
                    currentProcessNameIndex++;

                    var ngNames = GetNGNames(groupDataList, processType, processName);

                    // ProcessName별 진행률 표시
                    int taskProgress = (int)((currentProcessNameIndex / (double)totalProcessNames) * 100);
                    int totalProgress = 88 + (int)(taskProgress * 0.07); // 88% → 95%
                    clLogger.Log($"  [{currentProcessNameIndex}/{totalProcessNames}] Processing: {processType} - {processName} ({ngNames.Count} defects)");
                    ReportProgress(taskProgress, totalProgress, $"Step {currentProcessNameIndex}/{totalProcessNames}: {processType} - {processName}");

                    foreach (var ngName in ngNames)
                    {
                        var row = result.NewRow();
                        row["Lv.2"] = processType;

                        // Lv.2 다음 빈 컬럼
                        if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
                        {
                            row[SEPARATOR_AFTER_LV2] = DBNull.Value;
                        }

                        row["ProcessName"] = processName;
                        row["Lv.3"] = ngName;
                        row["Unit"] = "PPM";

                        // Unit 다음 빈 컬럼
                        if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
                        {
                            row[SEPARATOR_AFTER_UNIT] = DBNull.Value;
                        }

                        // 일자별
                        foreach (var date in cachedDates)
                        {
                            foreach (var group in groupedModels)
                            {
                                string colName = cachedColumnNames[(date, group.Index)];
                                if (!dateDataByGroup.TryGetValue(group.Index, out var dateDict))
                                {
                                    if (missingGroupWarnings.Add(group.Index))
                                    {
                                        clLogger.LogWarning($"DetailReport skipped missing group data. GroupIndex={group.Index}, GroupName='{group.GroupName}'");
                                    }

                                    row[colName] = 0d;
                                    continue;
                                }

                                var lookupKey = (processType, processName, ngName, date);
                                double ppm = 0;

                                if (dateDict.TryGetValue(lookupKey, out var data))
                                {
                                    ppm = data.InputQty > 0 ? Math.Round((data.NGQty / data.InputQty) * 1000000, 0) : 0;
                                }

                                row[colName] = ppm;

                                var key = (date, group.Index, processType);
                                if (!processTypePPMs.ContainsKey(key))
                                    processTypePPMs[key] = 0;
                                processTypePPMs[key] += ppm;
                            }
                        }

                        // 빈 컬럼 1
                        if (result.Columns.Contains(SEPARATOR_1))
                        {
                            row[SEPARATOR_1] = DBNull.Value;
                        }

                        // 주별
                        foreach (var week in cachedWeeks)
                        {
                            foreach (var group in groupedModels)
                            {
                                string colName = cachedColumnNames[($"W{week}", group.Index)];
                                if (!weekDataByGroup.TryGetValue(group.Index, out var weekDict))
                                {
                                    if (missingGroupWarnings.Add(group.Index))
                                    {
                                        clLogger.LogWarning($"DetailReport skipped missing group data. GroupIndex={group.Index}, GroupName='{group.GroupName}'");
                                    }

                                    row[colName] = 0d;
                                    continue;
                                }

                                var lookupKey = (processType, processName, ngName, week);
                                double ppm = 0;

                                if (weekDict.TryGetValue(lookupKey, out var data))
                                {
                                    ppm = data.InputQty > 0 ? Math.Round((data.NGQty / data.InputQty) * 1000000, 0) : 0;
                                }

                                row[colName] = ppm;

                                string periodKey = $"W{week}";
                                var key = (periodKey, group.Index, processType);
                                if (!processTypePPMs.ContainsKey(key))
                                    processTypePPMs[key] = 0;
                                processTypePPMs[key] += ppm;
                            }
                        }

                        // 빈 컬럼 2
                        if (result.Columns.Contains(SEPARATOR_2))
                        {
                            row[SEPARATOR_2] = DBNull.Value;
                        }

                        // 월별
                        foreach (var month in cachedMonths)
                        {
                            foreach (var group in groupedModels)
                            {
                                string colName = cachedColumnNames[($"M{month}", group.Index)];
                                if (!monthDataByGroup.TryGetValue(group.Index, out var monthDict))
                                {
                                    if (missingGroupWarnings.Add(group.Index))
                                    {
                                        clLogger.LogWarning($"DetailReport skipped missing group data. GroupIndex={group.Index}, GroupName='{group.GroupName}'");
                                    }

                                    row[colName] = 0d;
                                    continue;
                                }

                                var lookupKey = (processType, processName, ngName, month);
                                double ppm = 0;

                                if (monthDict.TryGetValue(lookupKey, out var data))
                                {
                                    ppm = data.InputQty > 0 ? Math.Round((data.NGQty / data.InputQty) * 1000000, 0) : 0;
                                }

                                row[colName] = ppm;

                                string periodKey = $"M{month}";
                                var key = (periodKey, group.Index, processType);
                                if (!processTypePPMs.ContainsKey(key))
                                    processTypePPMs[key] = 0;
                                processTypePPMs[key] += ppm;
                            }
                        }

                        result.Rows.Add(row);
                    }
                }
            }

            return processTypePPMs;
        }
    }
}
