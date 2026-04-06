using DataMaker.Logger;
using DataMaker.R6;
using DataMaker.R6.Grouping;
using DataMaker.R6.SQLProcess;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows;
using static DataMaker.R6.CONSTANT;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DataMaker.R6.PreProcessor
{
    /// <summary>
    /// ProcTable 생성 클래스
    /// OriginalTable → ProcTable (선택된 모델들만 필터링)
    /// </summary>
    public class clMakeProcTable
    {
        private clSQLFileIO sql;
        private List<string> selectedModels;

        public clMakeProcTable(string dbPath)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
                throw new ArgumentException("Database path cannot be null or empty.", nameof(dbPath));

            sql = new clSQLFileIO(dbPath);
        }

        /// <summary>
        /// ProcTable 생성 (선택된 모델들만)
        /// </summary>
        /// <param name="modelList">선택된 모델명 리스트 (MATERIALNAME)</param>
        public void Run(List<string> modelList)
        {
            if (modelList == null || modelList.Count == 0)
            {
                clLogger.LogWarning("No models selected. ProcTable will not be created.");
                return;
            }

            this.selectedModels = modelList.Distinct().ToList(); // 중복 제거

            try
            {
                clLogger.Log("=== Start Creating ProcTable ===");
                clLogger.Log($"Selected models count: {selectedModels.Count}");
                clLogger.Log($"Models: {string.Join(", ", selectedModels.Take(10))}" +
                           (selectedModels.Count > 10 ? "..." : ""));

                // 0. 기존 테이블 구조 검증 (NGCODE 컬럼 존재 여부 확인)
                ValidateTableStructure();

                // 1. 기존 ProcTable 삭제
                DropExistingTable();

                // 2. ProcTable 구조 생성
                CreateTableStructure();

                // 3. 선택된 모델들의 데이터만 복사 및 중복 제거
                CopySelectedModelsData();

                // 4. 필터링된 ProcTable에 후처리 적용
                SetDerivedColumns();
                MapLookupColumns();

                // 5. 완료
                int rowCount = sql.GetRowCount(OPTION_TABLE_NAME.PROC);
                clLogger.Log($"ProcTable created successfully with {rowCount} rows");
                clLogger.Log("=== ProcTable Creation Completed ===");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error in clMakeProcTable.Run");
                throw;
            }
        }

        /// <summary>
        /// MainForm의 selectedModels 형식으로 실행
        /// </summary>
        /// <param name="groupedModels">그룹별 모델 리스트</param>
        public void Run(List<clModelGroupData> groupedModels)
        {
            if (groupedModels == null || groupedModels.Count == 0)
            {
                clLogger.LogWarning("No grouped models provided. ProcTable will not be created.");
                return;
            }

            // 모든 그룹의 모델들을 하나의 리스트로 통합 (중복 제거)
            var allModels = new HashSet<string>();

            foreach (var group in groupedModels)
            {
                clLogger.Log($"Group {group.Index}: {group.ModelList.Count} models");

                foreach (var model in group.ModelList)
                {
                    allModels.Add(model);
                }
            }

            clLogger.Log($"Total unique models across all groups: {allModels.Count}");

            // ProcTable 생성
            Run(allModels.ToList());
        }

        /// <summary>
        /// 0. 기존 테이블 구조 검증 (NGCODE 컬럼 존재 확인)
        /// </summary>
        private void ValidateTableStructure()
        {
            if (sql.IsTableExist(OPTION_TABLE_NAME.PROC))
            {
                var columns = sql.GetColumns(OPTION_TABLE_NAME.PROC);

                // NGCODE 컬럼이 없으면 구조가 오래된 것
                if (!columns.Contains(CONSTANT.NGCODE.NEW))
                {
                    clLogger.LogWarning($"  ⚠ Detected old ProcTable structure without NGCODE column");
                    clLogger.Log($"  → Auto-regenerating ProcTable with updated schema...");
                }
            }
        }

        /// <summary>
        /// 1. 기존 ProcTable 삭제
        /// </summary>
        private void DropExistingTable()
        {
            if (sql.IsTableExist(OPTION_TABLE_NAME.PROC))
            {
                clLogger.Log("  - Dropping existing ProcTable");
                sql.DropTable(OPTION_TABLE_NAME.PROC);
            }
        }

        /// <summary>
        /// 2. ProcTable 구조 생성
        /// </summary>
        private void CreateTableStructure()
        {
            clLogger.Log("  - Creating ProcTable structure");

            if (!sql.IsTableExist(OPTION_TABLE_NAME.ORG))
            {
                throw new InvalidOperationException("OriginalTable does not exist. Please load the DB first.");
            }

            var orgColumns = sql.GetColumns(OPTION_TABLE_NAME.ORG);
            if (!orgColumns.Contains(CONSTANT.NGCODE.NEW))
            {
                throw new InvalidOperationException(
                    $"OriginalTable does not have NGCODE column. " +
                    $"Please refresh the DB with the updated schema.");
            }

            string[] columnNames = clOption.GetAccessTableColumns().Select(c => c.ColumnName).ToArray();
            string[] columnTypes = clOption.GetAccessTableColumns()
                .Select(c => GetSQLiteType(c.ColumnType))
                .ToArray();
            sql.CreateTable(OPTION_TABLE_NAME.PROC, columnNames, columnTypes);

            clLogger.Log("    - ProcTable structure created");
        }

        /// <summary>
        /// 3. 선택된 모델들의 데이터만 OriginalTable에서 ProcTable로 복사
        /// </summary>
        private void CopySelectedModelsData()
        {
            clLogger.Log("  - Copying data for selected models");

            sql.Processor.CopyCommonColumns(
                sourceTable: OPTION_TABLE_NAME.ORG,
                targetTable: OPTION_TABLE_NAME.PROC,
                columnName: CONSTANT.ModelName_WithLineShift.NEW,
                values: selectedModels
            );

            var normalizedData = ApplySourceNormalizationAndDuplicateSelection();
            var deduplicatedData = RemoveDuplicates(normalizedData);

            sql.DropTable(OPTION_TABLE_NAME.PROC);
            CreateTableStructure();
            sql.Writer.Write(OPTION_TABLE_NAME.PROC, deduplicatedData);

            clLogger.Log($"    - Prepared {deduplicatedData.Rows.Count} rows for {selectedModels.Count} models");
        }

        private DataTable ApplySourceNormalizationAndDuplicateSelection()
        {
            clLogger.Log("  - Applying source normalization and duplicate selection in ProcTable");

            var procData = sql.LoadTable(OPTION_TABLE_NAME.PROC);
            var selectedRows = SelectRowsForProcTable(procData);
            var normalizedData = procData.Clone();

            foreach (var sourceRow in selectedRows)
            {
                var procRow = normalizedData.NewRow();

                foreach (DataColumn col in normalizedData.Columns)
                {
                    if (!sourceRow.Table.Columns.Contains(col.ColumnName))
                    {
                        continue;
                    }

                    string value = sourceRow[col.ColumnName]?.ToString() ?? string.Empty;
                    if (col.ColumnName == CONSTANT.PROCESSNAME.NEW || col.ColumnName == CONSTANT.NGNAME.NEW)
                    {
                        procRow[col.ColumnName] = CONSTANT.Normalize(value);
                    }
                    else
                    {
                        procRow[col.ColumnName] = sourceRow[col.ColumnName];
                    }
                }

                normalizedData.Rows.Add(procRow);
            }

            clLogger.Log($"    - Source duplicate selection completed: {normalizedData.Rows.Count} rows kept");
            return normalizedData;
        }

        private void SetDerivedColumns()
        {
            clLogger.Log("  - Setting derived column values in ProcTable");
            sql.Processor.SetEmptyColumnsValueInProcTable(OPTION_TABLE_NAME.PROC);
        }

        private void MapLookupColumns()
        {
            clLogger.Log("  - Mapping ProcessType and Reason in ProcTable");

            var procData = sql.LoadTable(OPTION_TABLE_NAME.PROC);
            if (procData == null || procData.Rows.Count == 0)
            {
                clLogger.Log("    - ProcTable is empty, skipping lookup mapping");
                return;
            }

            var processTypeLookup = BuildProcessTypeLookup();
            var reasonLookup = BuildReasonLookup();

            int processTypeMappedCount = 0;
            int reasonMappedCount = 0;

            foreach (DataRow row in procData.Rows)
            {
                if (TryMapProcessType(row, processTypeLookup))
                {
                    processTypeMappedCount++;
                }

                if (TryMapReason(row, reasonLookup))
                {
                    reasonMappedCount++;
                }
            }

            sql.DropTable(OPTION_TABLE_NAME.PROC);
            CreateTableStructure();
            sql.Writer.Write(OPTION_TABLE_NAME.PROC, procData);

            clLogger.Log($"    - ProcessType mapping completed in ProcTable ({processTypeMappedCount} rows)");
            clLogger.Log($"    - Reason mapping completed in ProcTable ({reasonMappedCount} rows)");
        }

        private Dictionary<(string MaterialName, string ProcessCode, string ProcessName), string> BuildProcessTypeLookup()
        {
            if (!sql.IsTableExist(OPTION_TABLE_NAME.ROUTING))
            {
                clLogger.LogWarning("    - ProcessTypeTable not found, skipping ProcessType mapping");
                return new Dictionary<(string MaterialName, string ProcessCode, string ProcessName), string>();
            }

            var routingData = sql.LoadTable(OPTION_TABLE_NAME.ROUTING);
            var lookup = new Dictionary<(string MaterialName, string ProcessCode, string ProcessName), string>();

            foreach (DataRow row in routingData.Rows)
            {
                string materialName = row["모델명"]?.ToString() ?? string.Empty;
                string processCode = row["ProcessCode"]?.ToString() ?? string.Empty;
                string processName = row["ProcessName"]?.ToString() ?? string.Empty;
                string processType = row["ProcessType"]?.ToString() ?? string.Empty;

                lookup[(materialName, processCode, processName)] = processType;
            }

            return lookup;
        }

        private Dictionary<(string ProcessName, string NgName), string> BuildReasonLookup()
        {
            if (!sql.IsTableExist(OPTION_TABLE_NAME.REASON))
            {
                clLogger.LogWarning("    - ReasonTable not found, skipping Reason mapping");
                return new Dictionary<(string ProcessName, string NgName), string>();
            }

            var reasonData = sql.LoadTable(OPTION_TABLE_NAME.REASON);
            var lookup = new Dictionary<(string ProcessName, string NgName), string>();

            foreach (DataRow row in reasonData.Rows)
            {
                string processName = row["processName"]?.ToString() ?? string.Empty;
                string ngName = row["NgName"]?.ToString() ?? string.Empty;
                string reason = row["Reason"]?.ToString() ?? string.Empty;

                lookup[(processName, ngName)] = reason;
            }

            return lookup;
        }

        private bool TryMapProcessType(
            DataRow row,
            Dictionary<(string MaterialName, string ProcessCode, string ProcessName), string> processTypeLookup)
        {
            if (processTypeLookup.Count == 0)
            {
                return false;
            }

            string materialName = row[CONSTANT.MATERIALNAME.NEW]?.ToString() ?? string.Empty;
            string processCode = row[CONSTANT.PROCESSCODE.NEW]?.ToString() ?? string.Empty;
            string processName = row[CONSTANT.PROCESSNAME.NEW]?.ToString() ?? string.Empty;

            if (!processTypeLookup.TryGetValue((materialName, processCode, processName), out string processType))
            {
                return false;
            }

            row[CONSTANT.PROCESSTYPE.NEW] = processType;
            return true;
        }

        private bool TryMapReason(
            DataRow row,
            Dictionary<(string ProcessName, string NgName), string> reasonLookup)
        {
            if (reasonLookup.Count == 0)
            {
                return false;
            }

            string processName = row[CONSTANT.PROCESSNAME.NEW]?.ToString() ?? string.Empty;
            string ngName = row[CONSTANT.NGNAME.NEW]?.ToString() ?? string.Empty;

            if (!reasonLookup.TryGetValue((processName, ngName), out string reason))
            {
                return false;
            }

            row[CONSTANT.REASON.NEW] = reason;
            return true;
        }

        /// <summary>
        /// 4. 중복 제거 - DataTable 방식
        /// 같은 PRODUCTION_LINE, 날짜, ProcessName, NGName, MaterialName이면 대표값 1개로 합침
        /// QTYINPUT: 평균값, QTYNG: 합계값
        /// </summary>
        private DataTable RemoveDuplicates(DataTable data)
        {
            clLogger.Log("  - Removing duplicates (ProductionLine, Date, ProcessName, NGName, MaterialName)");

            int beforeCount = data.Rows.Count;
            clLogger.Log($"    - Before: {beforeCount} rows");

            if (data.Rows.Count == 0)
            {
                clLogger.Log("    - No data to process");
                return data;
            }

            // 중복 판단: PRODUCTION_LINE, 날짜, ProcessName, NGName, NGCode, MaterialName 기준
            var groups = data.AsEnumerable()
                .GroupBy(row => new
                {
                    ProductionLine = row.Field<string>(CONSTANT.PRODUCTION_LINE.NEW),
                    Date = row.Field<string>(CONSTANT.PRODUCT_DATE.NEW),
                    ProcessName = row.Field<string>(CONSTANT.PROCESSNAME.NEW),
                    NGName = row.Field<string>(CONSTANT.NGNAME.NEW),
                    NGCode = row.Field<string>(CONSTANT.NGCODE.NEW),
                    MaterialName = row.Field<string>(CONSTANT.ModelName_WithLineShift.NEW)
                });

            // 새 DataTable 생성
            var newData = data.Clone();
            int duplicateGroupCount = 0;
            int totalMergedRows = 0;

            foreach (var group in groups)
            {
                if (group.Count() == 1)
                {
                    // 중복 없음 - 그대로 유지
                    newData.ImportRow(group.First());
                }
                else
                {
                    // 중복 있음 - 새 행 생성 (데이터 집계)
                    duplicateGroupCount++;
                    totalMergedRows += group.Count();

                    var firstRow = group.First();
                    var newRow = newData.NewRow();

                    // 모든 컬럼 복사 (기본값)
                    foreach (DataColumn col in data.Columns)
                    {
                        newRow[col.ColumnName] = firstRow[col.ColumnName];
                    }

                    // QTYINPUT: 평균값
                    double avgInput = group.Average(r => Convert.ToDouble(r[CONSTANT.QTYINPUT.NEW]));
                    newRow[CONSTANT.QTYINPUT.NEW] = avgInput;

                    // QTYNG: 합계값
                    double sumNG = group.Sum(r => Convert.ToDouble(r[CONSTANT.QTYNG.NEW]));
                    newRow[CONSTANT.QTYNG.NEW] = sumNG;

                    newData.Rows.Add(newRow);
                }
            }

            int afterCount = newData.Rows.Count;
            int removedCount = beforeCount - afterCount;

            clLogger.Log($"    - After: {afterCount} rows");
            clLogger.Log($"    - Aggregated {duplicateGroupCount} groups (merged {totalMergedRows} rows into {duplicateGroupCount} rows)");
            clLogger.Log($"    - Total reduction: {removedCount} rows");
            return newData;
        }

        /// <summary>
        /// ProcTable 존재 여부 확인
        /// </summary>
        public bool IsProcTableExists()
        {
            return sql.IsTableExist(OPTION_TABLE_NAME.PROC);
        }

        /// <summary>
        /// ProcTable 행 개수 반환
        /// </summary>
        public int GetProcTableRowCount()
        {
            if (!IsProcTableExists())
                return 0;

            return sql.GetRowCount(OPTION_TABLE_NAME.PROC);
        }

        /// <summary>
        /// ProcTable의 고유 모델 목록 반환
        /// </summary>
        public List<string> GetModelsInProcTable()
        {
            if (!IsProcTableExists())
                return new List<string>();

            return sql.GetUniqueValues(OPTION_TABLE_NAME.PROC, CONSTANT.MATERIALNAME.NEW);
        }

        /// <summary>
        /// 리소스 해제
        /// </summary>
        public void Dispose()
        {
            sql?.Dispose();
        }

        private List<DataRow> SelectRowsForProcTable(DataTable procData)
        {
            var selectedRows = new List<DataRow>();

            if (procData == null || procData.Rows.Count == 0)
            {
                return selectedRows;
            }

            string[] requiredColumns =
            {
                CONSTANT.PRODUCTION_LINE.NEW,
                CONSTANT.PROCESSCODE.NEW,
                CONSTANT.PROCESSNAME.NEW,
                CONSTANT.NGNAME.NEW,
                CONSTANT.MATERIALNAME.NEW,
                CONSTANT.PRODUCT_DATE.NEW,
                CONSTANT.SHIFT.NEW,
                CONSTANT.QTYINPUT.NEW,
                CONSTANT.QTYNG.NEW
            };

            bool hasAllColumns = requiredColumns.All(c => procData.Columns.Contains(c));
            if (!hasAllColumns)
            {
                clLogger.LogWarning("    - Source duplicate filtering skipped: required columns missing in ProcTable");
                return procData.AsEnumerable().ToList();
            }

            int duplicateGroupCount = 0;
            int filteredOutRows = 0;
            int manualSelectionGroupCount = 0;

            var groups = procData.AsEnumerable().GroupBy(row => new
            {
                ProductionLine = GetStringValue(row, CONSTANT.PRODUCTION_LINE.NEW),
                ProcessCode = GetStringValue(row, CONSTANT.PROCESSCODE.NEW),
                ProcessName = CONSTANT.Normalize(GetStringValue(row, CONSTANT.PROCESSNAME.NEW)),
                NgName = CONSTANT.Normalize(GetStringValue(row, CONSTANT.NGNAME.NEW)),
                MaterialName = GetStringValue(row, CONSTANT.MATERIALNAME.NEW),
                ProductDate = GetStringValue(row, CONSTANT.PRODUCT_DATE.NEW),
                Shift = GetStringValue(row, CONSTANT.SHIFT.NEW)
            });

            foreach (var group in groups)
            {
                var groupRows = group.ToList();
                if (groupRows.Count <= 1)
                {
                    selectedRows.AddRange(groupRows);
                    continue;
                }

                duplicateGroupCount++;

                string groupSummary =
                    $"PRODUCTION_LINE={group.Key.ProductionLine}, " +
                    $"PROCESSCODE={group.Key.ProcessCode}, " +
                    $"PROCESSNAME={group.Key.ProcessName}, " +
                    $"NGNAME={group.Key.NgName}, " +
                    $"MATERIALNAME={group.Key.MaterialName}, " +
                    $"PRODUCT_DATE={group.Key.ProductDate}, " +
                    $"SHIFT={group.Key.Shift}";

                var selectedRow = SelectRowByComparingTwoRows(groupRows, groupSummary);
                if (selectedRow == null)
                {
                    clLogger.LogWarning("    - Manual selection returned no row. Dropping the group.");
                    filteredOutRows += groupRows.Count;
                    continue;
                }

                manualSelectionGroupCount++;
                selectedRows.Add(selectedRow);
                filteredOutRows += groupRows.Count - 1;
            }

            clLogger.Log($"    - Source duplicate groups checked: {duplicateGroupCount}");
            clLogger.Log($"    - Source rows filtered out: {filteredOutRows}");
            if (manualSelectionGroupCount > 0)
            {
                clLogger.Log($"    - Source groups resolved by manual comparison: {manualSelectionGroupCount}");
            }

            return selectedRows;
        }

        private DataRow? SelectRowByComparingTwoRows(List<DataRow> rows, string groupSummary)
        {
            if (rows == null || rows.Count == 0)
            {
                return null;
            }

            if (rows.Count == 1)
            {
                return rows[0];
            }

            DataRow selected = rows[rows.Count - 1];
            for (int i = rows.Count - 2; i >= 0; i--)
            {
                if (TryResolveRowsWithoutPrompt(rows[i], selected, out DataRow resolved))
                {
                    selected = resolved;
                    continue;
                }

                selected = PromptUserToChooseRow(rows[i], selected, groupSummary);
            }

            return selected;
        }

        private DataRow PromptUserToChooseRow(DataRow optionA, DataRow optionB, string groupSummary)
        {
            MessageBoxResult ShowDialogAndGetChoice()
            {
                string optionAText = FormatRowForSelection(optionA);
                string optionBText = FormatRowForSelection(optionB);

                return MessageBox.Show(
                    "중복 데이터가 있습니다. 아래 2개 행 중 하나를 선택해주세요.\n\n" +
                    $"{groupSummary}\n\n" +
                    "[A]\n" + optionAText + "\n\n" +
                    "[B]\n" + optionBText + "\n\n" +
                    "YES: A 선택\n" +
                    "NO: B 선택\n" +
                    "CANCEL: B 선택",
                    "중복 행 선택",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);
            }

            if (Application.Current?.Dispatcher == null)
            {
                return ShowDialogAndGetChoice() == MessageBoxResult.Yes ? optionA : optionB;
            }

            MessageBoxResult result = Application.Current.Dispatcher.CheckAccess()
                ? ShowDialogAndGetChoice()
                : Application.Current.Dispatcher.Invoke(ShowDialogAndGetChoice);

            return result == MessageBoxResult.Yes ? optionA : optionB;
        }

        private string FormatRowForSelection(DataRow row)
        {
            string qtyInput = GetStringValue(row, CONSTANT.QTYINPUT.NEW);
            string qtyNg = GetStringValue(row, CONSTANT.QTYNG.NEW);
            string shift = row.Table.Columns.Contains(CONSTANT.SHIFT.NEW)
                ? GetStringValue(row, CONSTANT.SHIFT.NEW)
                : string.Empty;

            return $"QTYINPUT={qtyInput}, QTYNG={qtyNg}, SHIFT={shift}";
        }

        private static string GetStringValue(DataRow row, string columnName)
        {
            return row[columnName]?.ToString()?.Trim() ?? string.Empty;
        }

        private static bool AreRowsEquivalent(DataRow left, DataRow right)
        {
            if (left == null || right == null || left.Table == null || right.Table == null)
            {
                return false;
            }

            return
                AreColumnValuesEquivalent(left, right, CONSTANT.QTYINPUT.NEW, isNumeric: true) &&
                AreColumnValuesEquivalent(left, right, CONSTANT.QTYNG.NEW, isNumeric: true) &&
                AreColumnValuesEquivalent(left, right, CONSTANT.SHIFT.NEW, isNumeric: false);
        }

        private static bool TryResolveRowsWithoutPrompt(DataRow optionA, DataRow optionB, out DataRow selectedRow)
        {
            selectedRow = optionB;

            if (AreRowsEquivalent(optionA, optionB))
            {
                return true;
            }

            if (!TryGetNumericColumnValue(optionA, CONSTANT.QTYINPUT.NEW, out decimal qtyInputA) ||
                !TryGetNumericColumnValue(optionB, CONSTANT.QTYINPUT.NEW, out decimal qtyInputB) ||
                qtyInputA != qtyInputB)
            {
                return false;
            }

            if (!TryGetNumericColumnValue(optionA, CONSTANT.QTYNG.NEW, out decimal qtyNgA) ||
                !TryGetNumericColumnValue(optionB, CONSTANT.QTYNG.NEW, out decimal qtyNgB))
            {
                return false;
            }

            bool isAZero = qtyNgA == 0m;
            bool isBZero = qtyNgB == 0m;

            if (isAZero != isBZero)
            {
                selectedRow = isAZero ? optionB : optionA;
                return true;
            }

            if (isAZero && isBZero)
            {
                return false;
            }

            selectedRow = CreateMergedRowWithSummedQty(optionA, optionB, qtyInputA, qtyNgA + qtyNgB);
            return true;
        }

        private static bool TryGetNumericColumnValue(DataRow row, string columnName, out decimal value)
        {
            value = 0m;

            if (row == null || row.Table == null || !row.Table.Columns.Contains(columnName))
            {
                return false;
            }

            return TryParseDecimal(row[columnName], out value);
        }

        private static DataRow CreateMergedRowWithSummedQty(DataRow optionA, DataRow optionB, decimal qtyInput, decimal mergedQtyNg)
        {
            DataRow mergedRow = CloneRow(optionB);
            SetNumericColumnValue(mergedRow, CONSTANT.QTYINPUT.NEW, qtyInput);
            SetNumericColumnValue(mergedRow, CONSTANT.QTYNG.NEW, mergedQtyNg);
            return mergedRow;
        }

        private static DataRow CloneRow(DataRow sourceRow)
        {
            if (sourceRow.Table == null)
            {
                return sourceRow;
            }

            DataRow clonedRow = sourceRow.Table.NewRow();
            clonedRow.ItemArray = (object[])sourceRow.ItemArray.Clone();
            return clonedRow;
        }

        private static void SetNumericColumnValue(DataRow row, string columnName, decimal value)
        {
            if (row.Table == null || !row.Table.Columns.Contains(columnName))
            {
                return;
            }

            DataColumn column = row.Table.Columns[columnName]!;
            Type columnType = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType;

            try
            {
                if (columnType == typeof(decimal))
                {
                    row[columnName] = value;
                }
                else if (columnType == typeof(double))
                {
                    row[columnName] = (double)value;
                }
                else if (columnType == typeof(float))
                {
                    row[columnName] = (float)value;
                }
                else if (columnType == typeof(long))
                {
                    row[columnName] = decimal.ToInt64(value);
                }
                else if (columnType == typeof(int))
                {
                    row[columnName] = decimal.ToInt32(value);
                }
                else if (columnType == typeof(short))
                {
                    row[columnName] = decimal.ToInt16(value);
                }
                else if (columnType == typeof(byte))
                {
                    row[columnName] = decimal.ToByte(value);
                }
                else
                {
                    row[columnName] = value.ToString(CultureInfo.InvariantCulture);
                }
            }
            catch
            {
                row[columnName] = value.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static bool AreColumnValuesEquivalent(DataRow left, DataRow right, string columnName, bool isNumeric)
        {
            if (!left.Table.Columns.Contains(columnName) || !right.Table.Columns.Contains(columnName))
            {
                return false;
            }

            object leftValue = left[columnName];
            object rightValue = right[columnName];

            if (isNumeric)
            {
                bool leftParsed = TryParseDecimal(leftValue, out decimal leftNumber);
                bool rightParsed = TryParseDecimal(rightValue, out decimal rightNumber);

                if (leftParsed && rightParsed)
                {
                    return leftNumber == rightNumber;
                }
            }

            string leftText = NormalizeValueForComparison(leftValue);
            string rightText = NormalizeValueForComparison(rightValue);
            return string.Equals(leftText, rightText, StringComparison.Ordinal);
        }

        private static string NormalizeValueForComparison(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }

            return value.ToString()?.Trim() ?? string.Empty;
        }

        private static bool TryParseDecimal(object value, out decimal parsed)
        {
            parsed = 0m;

            if (value == null || value == DBNull.Value)
            {
                return false;
            }

            if (value is decimal directDecimal)
            {
                parsed = directDecimal;
                return true;
            }

            string text = value.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed) ||
                   decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out parsed);
        }

        private static string GetSQLiteType(Type type)
        {
            if (type == typeof(int) || type == typeof(long) || type == typeof(short) ||
                type == typeof(byte) || type == typeof(bool))
                return "INTEGER";

            if (type == typeof(float) || type == typeof(double) || type == typeof(decimal))
                return "REAL";

            if (type == typeof(DateTime))
                return "TEXT";

            if (type == typeof(byte[]))
                return "BLOB";

            return "TEXT";
        }
    }
}
