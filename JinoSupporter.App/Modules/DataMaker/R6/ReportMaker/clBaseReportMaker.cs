using DataMaker.Logger;
using DataMaker.R6.Grouping;
using DataMaker.R6.SQLProcess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using static DataMaker.R6.CONSTANT;

namespace DataMaker.R6.Report
{
    #region Progress Event

    /// <summary>
    /// Report 진행 상황 이벤트 인자
    /// </summary>
    public class ProgressEventArgs : EventArgs
    {
        public int TaskProgress { get; set; }  // 현재 작업 진행률 (0-100)
        public int TotalProgress { get; set; }  // 전체 진행률 (0-100)
        public string Message { get; set; }
    }

    #endregion

    #region Base Report Maker

    /// <summary>
    /// Report 생성을 위한 베이스 클래스
    /// 공통 로직 포함
    /// </summary>
    public abstract class clBaseReportMaker
    {
        protected clSQLFileIO sql;
        protected List<clModelGroupData> groupedModels;
        protected List<long> cachedMonths;
        protected List<long> cachedWeeks;
        protected List<string> cachedDates;
        protected Dictionary<(string, int), string> cachedColumnNames;
        protected List<(int GroupIndex, DataTable Data)>? preloadedGroupDataList;

        // Progress 이벤트
        public event EventHandler<ProgressEventArgs> ProgressChanged;

        protected clBaseReportMaker(string dbPath, List<clModelGroupData> groupedModels)
        {
            sql = new clSQLFileIO(dbPath);
            this.groupedModels = groupedModels;
        }

        protected clBaseReportMaker(string dbPath, List<clModelGroupData> groupedModels, List<(int GroupIndex, DataTable Data)> preloadedGroupDataList)
            : this(dbPath, groupedModels)
        {
            this.preloadedGroupDataList = preloadedGroupDataList;
        }

        // Progress 보고 헬퍼 메서드
        protected void ReportProgress(int taskProgress, int totalProgress, string message)
        {
            ProgressChanged?.Invoke(this, new ProgressEventArgs
            {
                TaskProgress = taskProgress,
                TotalProgress = totalProgress,
                Message = message
            });
        }

        protected List<(int GroupIndex, DataTable Data)> LoadOrReuseGroupData(int progressStart = 0, int progressEnd = 80)
        {
            if (preloadedGroupDataList != null)
            {
                return preloadedGroupDataList;
            }

            var groupDataList = new List<(int GroupIndex, DataTable Data)>();
            int groupCount = groupedModels.Count;

            for (int i = 0; i < groupCount; i++)
            {
                var group = groupedModels[i];
                string tableName = CONSTANT.GetGroupTableName(group);

                clLogger.Log($"Loading {tableName}...");
                var data = sql.LoadTable(tableName);

                if (data != null && data.Rows.Count > 0)
                {
                    groupDataList.Add((group.Index, data));
                    clLogger.Log($"  Loaded {data.Rows.Count} rows from {tableName}");
                }
                else
                {
                    clLogger.Log($"  WARNING: {tableName} is empty or not found");
                }

                int taskProgress = groupCount == 0 ? 100 : (int)((i + 1) / (double)groupCount * 100);
                int totalProgress = progressStart + (int)((progressEnd - progressStart) * (taskProgress / 100.0));
                ReportProgress(taskProgress, totalProgress, $"Loading group {i + 1}/{groupCount}...");
            }

            return groupDataList;
        }

        #region Period Methods (월/주/일자)

        protected List<long> GetUniqueMonths(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            // 날짜 기준으로 YYYYMM 생성 (년도 고려한 정렬을 위해)
            var monthSet = new HashSet<string>();
            var monthsByGroup = new Dictionary<int, HashSet<string>>();

            foreach (var (groupIndex, data) in groupDataList)
            {
                if (data != null && data.Rows.Count > 0)
                {
                    clLogger.Log($"Group {groupIndex}: Processing {data.Rows.Count} rows for months");

                    if (!monthsByGroup.ContainsKey(groupIndex))
                        monthsByGroup[groupIndex] = new HashSet<string>();

                    foreach (DataRow row in data.Rows)
                    {
                        try
                        {
                            var dateStr = row.Field<string>(CONSTANT.PRODUCT_DATE.NEW);
                            if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out DateTime date))
                            {
                                string yyyyMM = date.ToString("yyyyMM");
                                monthSet.Add(yyyyMM);
                                monthsByGroup[groupIndex].Add(yyyyMM);
                            }
                        }
                        catch (Exception ex)
                        {
                            clLogger.Log($"Error reading date for month: {ex.Message}");
                        }
                    }

                    clLogger.Log($"  Group {groupIndex} months: {string.Join(", ", monthsByGroup[groupIndex].OrderByDescending(m => m))}");
                }
            }

            // YYYYMM을 내림차순 정렬 후 MM으로 변환
            var sortedMonths = monthSet.OrderByDescending(m => m).ToList();
            var result = sortedMonths.Select(m => long.Parse(m.Substring(4, 2))).ToList();

            clLogger.Log($"=== TOTAL months across all groups: {string.Join(", ", result)} (sorted by date) ===");
            clLogger.Log($"Will create {result.Count} month columns for each of {groupedModels.Count} groups = {result.Count * groupedModels.Count} total month columns");

            return result;
        }

        protected List<long> GetUniqueWeeks(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            // 날짜 기준으로 주차 생성
            var weekSet = new HashSet<string>();

            foreach (var (groupIndex, data) in groupDataList)
            {
                if (data != null && data.Rows.Count > 0)
                {
                    foreach (DataRow row in data.Rows)
                    {
                        try
                        {
                            var dateStr = row.Field<string>(CONSTANT.PRODUCT_DATE.NEW);
                            if (!string.IsNullOrEmpty(dateStr) && DateTime.TryParse(dateStr, out DateTime date))
                            {
                                int weekOfYear = System.Globalization.CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(
                                    date, System.Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday);
                                string yyyyWW = $"{date.Year:0000}{weekOfYear:00}";
                                weekSet.Add(yyyyWW);
                            }
                        }
                        catch { }
                    }
                }
            }

            // YYYYWW를 내림차순 정렬 후 WW로 변환
            var sortedWeeks = weekSet.OrderByDescending(w => w).ToList();
            var result = sortedWeeks.Select(w => long.Parse(w.Substring(4, 2))).ToList();

            return result;
        }

        protected List<string> GetUniqueDates(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var dates = new HashSet<string>();
            foreach (var (groupIndex, data) in groupDataList)
            {
                if (data != null && data.Rows.Count > 0)
                {
                    foreach (DataRow row in data.Rows)
                    {
                        try
                        {
                            var date = row.Field<string>(CONSTANT.PRODUCT_DATE.NEW);
                            if (!string.IsNullOrEmpty(date))
                                dates.Add(date);
                        }
                        catch { }
                    }
                }
            }
            return dates.OrderByDescending(d => d).ToList();
        }

        #endregion

        #region PPM Calculation Methods

        /// <summary>
        /// 통합 PPM 계산 메서드
        /// </summary>
        protected double CalculatePPM(IEnumerable<DataRow> filteredRows)
        {
            if (!filteredRows.Any())
                return 0;

            double totalInput = filteredRows.Sum(r => r.Field<double>(CONSTANT.QTYINPUT.NEW));
            double totalNG = filteredRows.Sum(r => r.Field<double>(CONSTANT.QTYNG.NEW));

            if (totalInput == 0)
                return 0;

            return Math.Round((totalNG / totalInput) * 1000000, 0);
        }

        /// <summary>
        /// 필터 조건을 적용하여 PPM 계산
        /// </summary>
        protected double GetPPMWithFilter(DataTable data, Func<DataRow, bool> filter)
        {
            if (data == null || data.Rows.Count == 0)
                return 0;

            var filtered = data.AsEnumerable().Where(filter);
            return CalculatePPM(filtered);
        }

        protected double GetMonthNGPPM(DataTable data, string processType, string processName, string ngName, long month)
        {
            return GetPPMWithFilter(data, row =>
                row.Field<string>(CONSTANT.PROCESSTYPE.NEW) == processType &&
                row.Field<string>(CONSTANT.PROCESSNAME.NEW) == processName &&
                row.Field<string>(CONSTANT.NGNAME.NEW) == ngName &&
                row.Field<long>(CONSTANT.MONTH.NEW) == month);
        }

        protected double GetWeekNGPPM(DataTable data, string processType, string processName, string ngName, long week)
        {
            return GetPPMWithFilter(data, row =>
                row.Field<string>(CONSTANT.PROCESSTYPE.NEW) == processType &&
                row.Field<string>(CONSTANT.PROCESSNAME.NEW) == processName &&
                row.Field<string>(CONSTANT.NGNAME.NEW) == ngName &&
                row.Field<long>(CONSTANT.WEEK.NEW) == week);
        }

        protected double GetPPM(DataTable data, string processType, string processName, string ngName, string date)
        {
            return GetPPMWithFilter(data, row =>
                row.Field<string>(CONSTANT.PROCESSTYPE.NEW) == processType &&
                row.Field<string>(CONSTANT.PROCESSNAME.NEW) == processName &&
                row.Field<string>(CONSTANT.NGNAME.NEW) == ngName &&
                row.Field<string>(CONSTANT.PRODUCT_DATE.NEW) == date);
        }

        #endregion

        #region Helper Methods

        protected string GetUniqueColumnName(string baseName, HashSet<string> existingNames)
        {
            string uniqueName = baseName;
            int counter = 1;

            while (existingNames.Contains(uniqueName))
            {
                uniqueName = $"{baseName}_{counter}";
                counter++;
            }

            return uniqueName;
        }

        protected List<string> GetUniqueProcessTypes(List<(int GroupIndex, DataTable Data)> groupDataList)
        {
            var processTypes = new HashSet<string>();
            foreach (var (groupIndex, data) in groupDataList)
            {
                if (data != null && data.Rows.Count > 0)
                {
                    foreach (DataRow row in data.Rows)
                    {
                        var pt = row.Field<string>(CONSTANT.PROCESSTYPE.NEW);
                        if (!string.IsNullOrEmpty(pt))
                            processTypes.Add(pt);
                    }
                }
            }
            return processTypes.ToList();
        }

        protected List<string> GetProcessNames(List<(int GroupIndex, DataTable Data)> groupDataList, string processType)
        {
            var processNames = new HashSet<string>();
            foreach (var (groupIndex, data) in groupDataList)
            {
                if (data != null && data.Rows.Count > 0)
                {
                    var names = data.AsEnumerable()
                        .Where(r => r.Field<string>(CONSTANT.PROCESSTYPE.NEW) == processType)
                        .Select(r => r.Field<string>(CONSTANT.PROCESSNAME.NEW))
                        .Where(n => !string.IsNullOrEmpty(n))
                        .Distinct();

                    foreach (var name in names)
                        processNames.Add(name);
                }
            }
            return processNames.ToList();
        }

        protected List<string> GetNGNames(List<(int GroupIndex, DataTable Data)> groupDataList, string processType, string processName)
        {
            var ngNames = new HashSet<string>();
            foreach (var (groupIndex, data) in groupDataList)
            {
                if (data != null && data.Rows.Count > 0)
                {
                    var names = data.AsEnumerable()
                        .Where(r => r.Field<string>(CONSTANT.PROCESSTYPE.NEW) == processType &&
                                   r.Field<string>(CONSTANT.PROCESSNAME.NEW) == processName)
                        .Select(r => r.Field<string>(CONSTANT.NGNAME.NEW))
                        .Where(n => !string.IsNullOrEmpty(n))
                        .Distinct();

                    foreach (var name in names)
                        ngNames.Add(name);
                }
            }
            return ngNames.ToList();
        }

        protected List<(string NgName, string NgCode)> GetNGNamesWithCodes(List<(int GroupIndex, DataTable Data)> groupDataList, string processType, string processName)
        {
            var ngItems = new HashSet<(string, string)>();
            foreach (var (groupIndex, data) in groupDataList)
            {
                if (data != null && data.Rows.Count > 0)
                {
                    var items = data.AsEnumerable()
                        .Where(r => r.Field<string>(CONSTANT.PROCESSTYPE.NEW) == processType &&
                                   r.Field<string>(CONSTANT.PROCESSNAME.NEW) == processName)
                        .Select(r => (
                            NgName: r.Field<string>(CONSTANT.NGNAME.NEW),
                            NgCode: r.Field<string>(CONSTANT.NGCODE.NEW)
                        ))
                        .Where(n => !string.IsNullOrEmpty(n.NgName))
                        .Distinct();

                    foreach (var item in items)
                        ngItems.Add(item);
                }
            }
            return ngItems.ToList();
        }

        #endregion

        #region Header Row Methods

        protected void AddPeriodHeaderRow(DataTable result)
        {
            var headerRow = result.NewRow();
            headerRow["Lv.2"] = "모델";

            if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
            {
                headerRow[SEPARATOR_AFTER_LV2] = "";
            }

            headerRow["ProcessName"] = "";
            headerRow["Lv.3"] = "";
            headerRow["Unit"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
            {
                headerRow[SEPARATOR_AFTER_UNIT] = "";
            }

            foreach (var date in cachedDates)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[(date, group.Index)];
                    headerRow[colName] = date;
                }
            }

            if (result.Columns.Contains("Separator_1") || result.Columns.Contains(" "))
            {
                if (result.Columns.Contains(" "))
                    headerRow[" "] = "";
                if (result.Columns.Contains("Separator_1"))
                    headerRow["Separator_1"] = "";
            }

            foreach (var week in cachedWeeks)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"W{week}", group.Index)];
                    headerRow[colName] = $"W{week}";
                }
            }

            if (result.Columns.Contains("Separator_2") || result.Columns.Contains("  "))
            {
                if (result.Columns.Contains("  "))
                    headerRow["  "] = "";
                if (result.Columns.Contains("Separator_2"))
                    headerRow["Separator_2"] = "";
            }

            foreach (var month in cachedMonths)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"M{month}", group.Index)];
                    headerRow[colName] = $"M{month}";
                }
            }

            result.Rows.Add(headerRow);
        }

        protected void AddGroupHeaderRow(DataTable result)
        {
            var headerRow = result.NewRow();
            headerRow["Lv.2"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
            {
                headerRow[SEPARATOR_AFTER_LV2] = "";
            }

            headerRow["ProcessName"] = "";
            headerRow["Lv.3"] = "";
            headerRow["Unit"] = "";

            if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
            {
                headerRow[SEPARATOR_AFTER_UNIT] = "";
            }

            foreach (var date in cachedDates)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[(date, group.Index)];
                    string groupName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m).Take(2));
                    headerRow[colName] = groupName;
                }
            }

            if (result.Columns.Contains("Separator_1") || result.Columns.Contains(" "))
            {
                if (result.Columns.Contains(" "))
                    headerRow[" "] = "";
                if (result.Columns.Contains("Separator_1"))
                    headerRow["Separator_1"] = "";
            }

            foreach (var week in cachedWeeks)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"W{week}", group.Index)];
                    string groupName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m).Take(2));
                    headerRow[colName] = groupName;
                }
            }

            if (result.Columns.Contains("Separator_2") || result.Columns.Contains("  "))
            {
                if (result.Columns.Contains("  "))
                    headerRow["  "] = "";
                if (result.Columns.Contains("Separator_2"))
                    headerRow["Separator_2"] = "";
            }

            foreach (var month in cachedMonths)
            {
                foreach (var group in groupedModels)
                {
                    string colName = cachedColumnNames[($"M{month}", group.Index)];
                    string groupName = group.GroupName ?? string.Join("_", group.ModelList.OrderBy(m => m).Take(2));
                    headerRow[colName] = groupName;
                }
            }

            result.Rows.Add(headerRow);
        }

        #endregion

        #region Period Data Helpers

        // Separator 컬럼명 상수 (공백으로 표시)
        protected const string SEPARATOR_AFTER_LV2 = "    ";  // Lv.2 다음 빈 열
        protected const string SEPARATOR_AFTER_UNIT = "   ";  // Unit 다음 빈 열
        protected const string SEPARATOR_1 = " ";
        protected const string SEPARATOR_2 = "  ";

        /// <summary>
        /// Separator 컬럼에 빈 값 설정
        /// </summary>
        protected void SetSeparatorValue(DataRow row, DataTable result)
        {
            if (result.Columns.Contains(SEPARATOR_AFTER_LV2))
                row[SEPARATOR_AFTER_LV2] = DBNull.Value;
            if (result.Columns.Contains(SEPARATOR_AFTER_UNIT))
                row[SEPARATOR_AFTER_UNIT] = DBNull.Value;
            if (result.Columns.Contains(SEPARATOR_1))
                row[SEPARATOR_1] = DBNull.Value;
            if (result.Columns.Contains(SEPARATOR_2))
                row[SEPARATOR_2] = DBNull.Value;
        }

        /// <summary>
        /// 기간별 데이터 채우기 통합 메서드
        /// </summary>
        protected void FillPeriodData<T>(DataRow row, DataTable result,
            List<T> periods, string periodPrefix,
            Func<T, double> getPPMFunc)
        {
            foreach (var period in periods)
            {
                string colName = $"{periodPrefix}{period}";
                if (result.Columns.Contains(colName))
                {
                    row[colName] = getPPMFunc(period);
                }
            }
        }

        #endregion

        #region Abstract Methods (각 Report에서 구현)

        /// <summary>
        /// Report 생성 메서드 (각 파생 클래스에서 구현)
        /// </summary>
        public abstract Task<DataTable> CreateReport();

        #endregion

        public void Dispose()
        {
            sql?.Dispose();
        }
    }

    #endregion
}
