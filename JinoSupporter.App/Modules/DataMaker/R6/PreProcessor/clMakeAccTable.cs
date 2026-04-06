using DataMaker.Logger;
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
    /// AccessTable 생성 클래스
    /// OriginalTable → AccessTable 변환
    /// - ProcessType 매핑
    /// - Reason 매핑
    /// - 파생 컬럼 생성 (MONTH, WEEK, LINE_REMOVE, etc.)
    /// </summary>
    public class clMakeAccTable : IDisposable
    {
        public clSQLFileIO sql { get; set; }
        public List<(string ColumnName, Type ColumnType)> Columns { get; set; }
        private string _dbPath; // dbPath 저장

        public clMakeAccTable(string dbPath, List<(string, Type)> columns)
        {
            if (string.IsNullOrWhiteSpace(dbPath))
                throw new ArgumentException("Database path cannot be null or empty.", nameof(dbPath));

            if (columns == null || columns.Count == 0)
                throw new ArgumentException("Columns cannot be null or empty.", nameof(columns));

            _dbPath = dbPath; // 저장
            sql = new clSQLFileIO(dbPath);
            Columns = columns;
        }

        /// <summary>
        /// AccessTable 생성 프로세스 실행
        /// </summary>
        public void Run()
        {
            try
            {
                // 0. 기존 테이블 구조 검증 (NGCODE 컬럼 존재 여부 확인)
                ValidateTableStructure();

                // 1. 기존 AccessTable 삭제
                DropExistingTable();

                // 2. AccessTable 구조 생성
                CreateTableStructure();

                // 3. OriginalTable 데이터 복사 (Normalize 적용)
                CopyDataFromOriginalTable();

                // 4. 모델 필터링용 최소 파생 컬럼만 설정
                SetMinimalDerivedColumns();

                clLogger.Log("AccessTable creation completed successfully");
            }
            catch (Exception ex)
            {
                clLogger.LogException(ex, "Error in clMakeAccTable.Run");
                throw;
            }
        }

        /// <summary>
        /// 0. 기존 테이블 구조 검증 (NGCODE 컬럼 존재 확인)
        /// </summary>
        private void ValidateTableStructure()
        {
            if (sql.IsTableExist(OPTION_TABLE_NAME.ACC))
            {
                var columns = sql.GetColumns(OPTION_TABLE_NAME.ACC);

                // NGCODE 컬럼이 없으면 구조가 오래된 것
                if (!columns.Contains(CONSTANT.NGCODE.NEW))
                {
                    clLogger.LogWarning($"  ⚠ Detected old AccessTable structure without NGCODE column");
                    clLogger.Log($"  → Auto-regenerating AccessTable with updated schema...");
                }
            }
        }

        /// <summary>
        /// 1. 기존 AccessTable 삭제
        /// </summary>
        private void DropExistingTable()
        {
            if (sql.IsTableExist(OPTION_TABLE_NAME.ACC))
            {
                clLogger.Log("  - Dropping existing AccessTable");
                sql.DropTable(OPTION_TABLE_NAME.ACC);

                // 검증: 테이블이 정말로 삭제되었는지 확인
                if (sql.IsTableExist(OPTION_TABLE_NAME.ACC))
                {
                    clLogger.LogWarning("    ⚠ WARNING: Table still exists after DROP! Retrying...");

                    // 강제로 연결 닫고 다시 열기 (dbPath를 먼저 저장)
                    sql.Dispose();
                    sql = new clSQLFileIO(_dbPath);

                    // 다시 시도
                    if (sql.IsTableExist(OPTION_TABLE_NAME.ACC))
                    {
                        sql.DropTable(OPTION_TABLE_NAME.ACC);

                        // 마지막 확인
                        if (sql.IsTableExist(OPTION_TABLE_NAME.ACC))
                        {
                            throw new InvalidOperationException("Failed to drop AccessTable after multiple attempts.");
                        }
                    }
                }

                clLogger.Log("    ✓ AccessTable dropped successfully");
            }
        }

        /// <summary>
        /// 2. AccessTable 구조 생성
        /// </summary>
        private void CreateTableStructure()
        {
            clLogger.Log("  - Creating AccessTable structure");

            string[] columnNames = Columns.Select(c => c.ColumnName).ToArray();
            string[] columnTypes = Columns.Select(c => GetSQLiteType(c.ColumnType)).ToArray();

            sql.CreateTable(OPTION_TABLE_NAME.ACC, columnNames, columnTypes);
        }

        /// <summary>
        /// 3. OriginalTable → AccessTable 데이터 복사 (공통 컬럼)
        /// ProcessName, NGName에 Normalize 적용
        /// </summary>
        private void CopyDataFromOriginalTable()
        {
            clLogger.Log("  - Copying data from OriginalTable with Normalize");

            // OriginalTable 데이터 로드
            var orgData = sql.LoadTable(OPTION_TABLE_NAME.ORG);
            var accData = sql.LoadTable(OPTION_TABLE_NAME.ACC);

            var selectedOrgRows = SelectRowsForAccTable(orgData);

            foreach (System.Data.DataRow orgRow in selectedOrgRows)
            {
                var accRow = accData.NewRow();

                // 공통 컬럼 복사
                foreach (System.Data.DataColumn col in orgData.Columns)
                {
                    if (accData.Columns.Contains(col.ColumnName))
                    {
                        string value = orgRow[col.ColumnName]?.ToString() ?? "";

                        // ProcessName, NGName에는 Normalize 적용
                        if (col.ColumnName == CONSTANT.PROCESSNAME.NEW || col.ColumnName == CONSTANT.NGNAME.NEW)
                        {
                            accRow[col.ColumnName] = CONSTANT.Normalize(value);
                        }
                        else
                        {
                            accRow[col.ColumnName] = value;
                        }
                    }
                }

                accData.Rows.Add(accRow);
            }

            // 한 번에 저장
            sql.Writer.Write(OPTION_TABLE_NAME.ACC, accData);

            int copiedRows = sql.GetRowCount(OPTION_TABLE_NAME.ACC);
            clLogger.Log($"    - Copied {copiedRows} rows with Normalize");
        }

        private List<DataRow> SelectRowsForAccTable(DataTable orgData)
        {
            var selectedRows = new List<DataRow>();

            if (orgData == null || orgData.Rows.Count == 0)
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

            bool hasAllColumns = requiredColumns.All(c => orgData.Columns.Contains(c));
            if (!hasAllColumns)
            {
                clLogger.LogWarning("    - Duplicate filtering skipped: required columns missing in OriginalTable");
                return orgData.AsEnumerable().ToList();
            }

            var sourceRows = orgData.AsEnumerable().ToList();

            int duplicateGroupCount = 0;
            int filteredOutRows = 0;
            int manualSelectionGroupCount = 0;

            var groups = sourceRows.GroupBy(row => new
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

            clLogger.Log($"    - Duplicate groups checked: {duplicateGroupCount}");
            clLogger.Log($"    - Rows filtered out by duplicate rule: {filteredOutRows}");
            if (manualSelectionGroupCount > 0)
            {
                clLogger.Log($"    - Groups resolved by manual comparison: {manualSelectionGroupCount}");
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

            // 최신 행을 기본 후보로 두고, 이전 행들과 2개씩 비교하며 사용자가 선택
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

            MessageBoxResult result;
            if (Application.Current?.Dispatcher != null)
            {
                result = Application.Current.Dispatcher.CheckAccess()
                    ? ShowDialogAndGetChoice()
                    : Application.Current.Dispatcher.Invoke(ShowDialogAndGetChoice);
            }
            else
            {
                result = ShowDialogAndGetChoice();
            }

            return result == MessageBoxResult.Yes ? optionA : optionB;
        }

        private string FormatRowForSelection(DataRow row)
        {
            string qtyInput = GetStringValue(row, CONSTANT.QTYINPUT.NEW);
            string qtyNg = GetStringValue(row, CONSTANT.QTYNG.NEW);
            string shift = row.Table.Columns.Contains(CONSTANT.SHIFT.NEW)
                ? GetStringValue(row, CONSTANT.SHIFT.NEW)
                : "";

            return $"QTYINPUT={qtyInput}, QTYNG={qtyNg}, SHIFT={shift}";
        }

        private static string GetStringValue(DataRow row, string columnName)
        {
            return row[columnName]?.ToString()?.Trim() ?? string.Empty;
        }

        private static bool AreRowsEquivalent(DataRow left, DataRow right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            if (left.Table == null || right.Table == null)
            {
                return false;
            }

            // 같은 중복 그룹에서는 키 컬럼이 이미 동일하므로,
            // 사용자 선택 기준이 되는 값(QTYINPUT/QTYNG/SHIFT)만 비교해
            // 동일하면 팝업 없이 1개만 유지한다.
            return
                AreColumnValuesEquivalent(left, right, CONSTANT.QTYINPUT.NEW, isNumeric: true) &&
                AreColumnValuesEquivalent(left, right, CONSTANT.QTYNG.NEW, isNumeric: true) &&
                AreColumnValuesEquivalent(left, right, CONSTANT.SHIFT.NEW, isNumeric: false);
        }

        private static bool TryResolveRowsWithoutPrompt(DataRow optionA, DataRow optionB, out DataRow selectedRow)
        {
            selectedRow = optionB;

            // 완전히 동일(사용자 선택 기준상)하면 하나만 유지
            if (AreRowsEquivalent(optionA, optionB))
            {
                return true;
            }

            // QTYINPUT이 같고 QTYNG가 0 / 비0으로 갈리면 비0 행을 자동 선택
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

            // 0 / 비0 조합은 비0 자동 선택
            if (isAZero != isBZero)
            {
                selectedRow = isAZero ? optionB : optionA;
                return true;
            }

            // 둘 다 0이면 자동 규칙 없음(필요시 수동 비교)
            if (isAZero && isBZero)
            {
                return false;
            }

            // QTYINPUT이 같고 QTYNG가 둘 다 비0이면 QTYNG 합산 후 1행으로 병합
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
            DataRow baseRow = optionB;
            DataRow mergedRow = CloneRow(baseRow);

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

        private static bool IsNonZero(DataRow row, string columnName)
        {
            object value = row[columnName];
            if (value == null || value == DBNull.Value)
            {
                return false;
            }

            if (value is double d) return d != 0d;
            if (value is float f) return f != 0f;
            if (value is decimal m) return m != 0m;
            if (value is int i) return i != 0;
            if (value is long l) return l != 0L;

            string text = value.ToString();
            if (string.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedInvariant))
            {
                return parsedInvariant != 0d;
            }

            if (double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out double parsedCurrent))
            {
                return parsedCurrent != 0d;
            }

            return false;
        }

        /// <summary>
        /// 4. 데이터 정규화 (PROCESSNAME, NGNAME)
        /// </summary>
        private void NormalizeData()
        {
            clLogger.Log("  - Normalizing data (PROCESSNAME, NGNAME)");

            sql.Processor.NormalizeColumns(
                OPTION_TABLE_NAME.ACC,
                new string[] { CONSTANT.PROCESSNAME.NEW, CONSTANT.NGNAME.NEW }
            );
        }

        /// <summary>
        /// 모델 필터링에 필요한 최소 파생 컬럼만 설정
        /// </summary>
        private void SetMinimalDerivedColumns()
        {
            clLogger.Log("  - Setting minimal derived values for AccessTable (LineShift only)");
            sql.Processor.SetLineShiftColumnValue(OPTION_TABLE_NAME.ACC);
        }

        /// <summary>
        /// .NET Type → SQLite Type 변환
        /// </summary>
        private string GetSQLiteType(Type type)
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

        /// <summary>
        /// 리소스 해제
        /// </summary>
        public void Dispose()
        {
            sql?.Dispose();
        }
    }
}
