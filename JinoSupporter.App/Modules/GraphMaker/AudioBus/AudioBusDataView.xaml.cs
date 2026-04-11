using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Text.Json;
using OfficeOpenXml;
using ExcelDataReader;
using UserControl = System.Windows.Controls.UserControl;
using CheckBox = System.Windows.Controls.CheckBox;
using MessageBox = System.Windows.MessageBox;

namespace GraphMaker
{
    public partial class AudioBusDataView : GraphViewBase
    {
        private List<FileInfo> excelFiles = new List<FileInfo>();
        private Dictionary<FileInfo, List<string>> fileSheetMapping = new Dictionary<FileInfo, List<string>>();
        private Dictionary<string, CheckBox> sheetCheckBoxes = new Dictionary<string, CheckBox>();
        public AudioBusDataView()
        {
            InitializeComponent();
            // EPPlus 라이선스 설정 (비상업적 용도)
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            // ExcelDataReader를 위한 Encoding 공급자 등록 (.xls 파일 지원)
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            UpdateWorkflowSummary();
            NotifyWebModuleSnapshotChanged();
        }

        private void AddLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                LogTextBox.ScrollToEnd();
                NotifyWebModuleSnapshotChanged();
            });
        }

        private void ClearLog()
        {
            Dispatcher.Invoke(() =>
            {
                LogTextBox.Clear();
                NotifyWebModuleSnapshotChanged();
            });
        }

        private void FolderDropBox_FolderSelected(object sender, JinoSupporter.Controls.FolderSelectedEventArgs e)
        {
            try
            {
                LoadExcelFilesFromFolder(e.FolderPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"폴더 선택 중 오류 발생:\n{ex.Message}", "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadExcelFilesFromFolder(string folderPath)
        {
            try
            {
                CurrentFolderText.Text = folderPath;
                excelFiles.Clear();
                fileSheetMapping.Clear();
                FileListBox.Items.Clear();

                // .xls 및 .xlsx 파일 검색
                var directory = new DirectoryInfo(folderPath);
                var xlsFiles = directory.GetFiles("*.xls", SearchOption.TopDirectoryOnly);
                var xlsxFiles = directory.GetFiles("*.xlsx", SearchOption.TopDirectoryOnly);

                excelFiles.AddRange(xlsFiles);
                excelFiles.AddRange(xlsxFiles);

                // 임시 파일 제외 (~ 로 시작하는 파일)
                excelFiles = excelFiles.Where(f => !f.Name.StartsWith("~")).ToList();

                if (excelFiles.Count == 0)
                {
                    MessageBox.Show("선택한 폴더에 Excel 파일이 없습니다.", "알림",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 각 파일의 시트 이름 추출
                foreach (var file in excelFiles)
                {
                    try
                    {
                        var sheetNames = GetSheetNames(file.FullName);
                        fileSheetMapping[file] = sheetNames;
                        FileListBox.Items.Add(file);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"파일 '{file.Name}' 읽기 오류:\n{ex.Message}", "경고",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }

                FileCountText.Text = $"파일: {excelFiles.Count}개";
                StatusText.Text = $"{excelFiles.Count}개의 Excel 파일 로드 완료";
                UpdateWorkflowSummary();

                // 모든 파일에 공통으로 존재하는 시트 표시
                DisplayCommonSheets();
                NotifyWebModuleSnapshotChanged();
            }
            catch (Exception ex)
            {
                StatusText.Text = "파일 로드 중 오류 발생";
                UpdateWorkflowSummary();
                MessageBox.Show($"파일 로드 중 오류 발생:\n{ex.Message}", "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateWorkflowSummary()
        {
            FileCountValueText.Text = excelFiles.Count.ToString("N0");
            if (string.IsNullOrWhiteSpace(CurrentFolderText.Text))
            {
                CurrentFolderText.Text = "(No folder)";
            }
        }

        private List<string> GetSheetNames(string filePath)
        {
            var sheetNames = new List<string>();
            var extension = Path.GetExtension(filePath).ToLower();

            if (extension == ".xlsx")
            {
                // .xlsx 파일은 EPPlus 사용
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        sheetNames.Add(worksheet.Name);
                    }
                }
            }
            else if (extension == ".xls")
            {
                // .xls 파일은 ExcelDataReader 사용
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var readerConfig = new ExcelDataReader.ExcelReaderConfiguration
                    {
                        FallbackEncoding = Encoding.GetEncoding(1252) // Windows-1252 인코딩
                    };

                    using (var reader = ExcelReaderFactory.CreateReader(stream, readerConfig))
                    {
                        do
                        {
                            if (!string.IsNullOrWhiteSpace(reader.Name))
                            {
                                sheetNames.Add(reader.Name);
                            }
                        } while (reader.NextResult());
                    }
                }
            }

            return sheetNames;
        }

        private void DisplayCommonSheets()
        {
            SheetCheckBoxPanel.Children.Clear();
            sheetCheckBoxes.Clear();

            if (fileSheetMapping.Count == 0)
            {
                return;
            }

            // 모든 고유한 시트 이름과 각 시트가 몇 개 파일에 있는지 계산
            var allSheets = fileSheetMapping.Values.SelectMany(sheets => sheets).Distinct().ToList();
            var sheetFileCounts = new Dictionary<string, int>();

            foreach (var sheetName in allSheets)
            {
                int count = fileSheetMapping.Values.Count(sheets => sheets.Contains(sheetName));
                sheetFileCounts[sheetName] = count;
            }

            if (sheetFileCounts.Count == 0)
            {
                var noSheetText = new TextBlock
                {
                    Text = "시트를 찾을 수 없습니다.",
                    Foreground = System.Windows.Media.Brushes.Red,
                    Margin = new Thickness(5),
                    TextWrapping = TextWrapping.Wrap
                };
                SheetCheckBoxPanel.Children.Add(noSheetText);
                StatusText.Text = "시트 없음";
                return;
            }

            // 파일 개수가 많은 순서대로 정렬
            var sortedSheets = sheetFileCounts.OrderByDescending(kvp => kvp.Value).ThenBy(kvp => kvp.Key);

            foreach (var kvp in sortedSheets)
            {
                var sheetName = kvp.Key;
                var fileCount = kvp.Value;

                var checkBox = new CheckBox
                {
                    Content = $"{sheetName} ({fileCount}개 파일)",
                    Margin = new Thickness(5),
                    FontSize = 12
                };

                sheetCheckBoxes[sheetName] = checkBox;
                SheetCheckBoxPanel.Children.Add(checkBox);
            }

            StatusText.Text = $"{sheetFileCounts.Count}개의 시트 발견";
        }

        private void ExtractDataButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 선택된 시트 확인
                var selectedSheetNames = sheetCheckBoxes
                    .Where(kvp => kvp.Value.IsChecked == true)
                    .Select(kvp => kvp.Key)
                    .ToList();

                if (selectedSheetNames.Count == 0)
                {
                    MessageBox.Show("시트를 선택해주세요.", "알림",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 저장 경로 선택
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel Files (*.xlsx)|*.xlsx",
                    FileName = "AudioBus_Data_Merged.xlsx",
                    Title = "병합된 데이터 저장 위치 선택"
                };

                if (saveDialog.ShowDialog() != true)
                {
                    return;
                }

                StatusText.Text = "데이터 추출 중...";
                ClearLog();
                AddLog("데이터 추출 시작...");

                string outputFilePath = saveDialog.FileName;

                // 기존 파일이 있으면 삭제 (파일이 열려있는지 확인)
                if (File.Exists(outputFilePath))
                {
                    try
                    {
                        File.Delete(outputFilePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"기존 파일을 삭제할 수 없습니다. 파일이 다른 프로그램에서 열려있는지 확인하세요.\n\n{ex.Message}",
                            "오류", MessageBoxButton.OK, MessageBoxImage.Error);
                        StatusText.Text = "저장 실패";
                        return;
                    }
                }

                // 각 시트별로 데이터 추출 및 병합
                using (var package = new ExcelPackage())
                {
                    int processedSheetCount = 0;

                    foreach (var sheetName in selectedSheetNames)
                    {
                        StatusText.Text = $"데이터 추출 중... ({sheetName})";

                        // 모든 파일에서 해당 시트의 데이터 읽기
                        var allFileData = new Dictionary<string, List<(double X, double Y)>>();

                        int successCount = 0;
                        int failCount = 0;

                        foreach (var file in excelFiles)
                        {
                            try
                            {
                                var data = ReadSheetData(file.FullName, sheetName);
                                if (data.Count > 0)
                                {
                                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(file.Name);
                                    allFileData[fileNameWithoutExt] = data;
                                    successCount++;
                                }
                                else
                                {
                                    AddLog($"[실패] {file.Name} - 데이터 없음");
                                    failCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                AddLog($"[오류] {file.Name} - {ex.Message}");
                                failCount++;
                            }
                        }

                        if (allFileData.Count > 0)
                        {
                            // 데이터 병합 및 Excel 시트에 작성
                            StatusText.Text = $"'{sheetName}' 시트 작성 중... ({successCount}개 파일)";
                            CreateMergedSheet(package, sheetName, allFileData);
                            processedSheetCount++;
                        }
                        else
                        {
                            MessageBox.Show($"'{sheetName}' 시트에서 데이터를 읽을 수 없습니다.\n\n성공: {successCount}개\n실패: {failCount}개\n\n파일 구조를 확인해주세요.",
                                "경고", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }

                    if (processedSheetCount == 0)
                    {
                        MessageBox.Show("추출할 데이터가 없습니다. 선택한 시트에 유효한 데이터가 있는지 확인하세요.",
                            "경고", MessageBoxButton.OK, MessageBoxImage.Warning);
                        StatusText.Text = "데이터 없음";
                        return;
                    }

                    // 파일 저장
                    StatusText.Text = "파일 저장 중...";
                    var fileInfo = new FileInfo(outputFilePath);
                    package.SaveAs(fileInfo);
                }

                StatusText.Text = "데이터 추출 완료";
                MessageBox.Show($"데이터가 성공적으로 저장되었습니다.\n\n파일: {outputFilePath}\n시트 개수: {selectedSheetNames.Count}개",
                    "완료", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"데이터 추출 중 오류 발생:\n{ex.Message}\n\n{ex.StackTrace}", "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusText.Text = "오류 발생";
            }
        }

        private List<(double X, double Y)> ReadSheetData(string filePath, string sheetName)
        {
            var data = new List<(double X, double Y)>();
            var extension = Path.GetExtension(filePath).ToLower();

            // Impedance 시트인지 확인
            bool isImpedanceSheet = sheetName.Contains("Impedanc", StringComparison.OrdinalIgnoreCase);

            if (extension == ".xlsx")
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null) return data;

                    if (isImpedanceSheet)
                    {
                        // Impedance 시트: 4행 1열에 X값 1개, 4행 2열에 Y값 1개
                        var xValueCell = worksheet.Cells[4, 1].Value;
                        var yValueCell = worksheet.Cells[4, 2].Value;

                        string fileName = Path.GetFileName(filePath);
                        AddLog($"[Impedance] {fileName}: X셀={xValueCell}, Y셀={yValueCell}");

                        if (xValueCell != null && yValueCell != null)
                        {
                            if (double.TryParse(xValueCell.ToString(), out double x) &&
                                double.TryParse(yValueCell.ToString(), out double y))
                            {
                                data.Add((x, y));
                                AddLog($"  ✓ 성공: X={x}, Y={y}");
                            }
                            else
                            {
                                AddLog($"  ✗ 파싱 실패 (숫자가 아님)");
                            }
                        }
                        else
                        {
                            AddLog($"  ✗ 셀이 비어있음 (null)");
                        }
                    }
                    else
                    {
                        // 일반 시트: 5행부터 시작, 1열=X, 2열=Y
                        int row = 5;
                        while (row <= worksheet.Dimension?.End.Row)
                        {
                            var xValueCell = worksheet.Cells[row, 1].Value;
                            var yValueCell = worksheet.Cells[row, 2].Value;

                            if (xValueCell != null && yValueCell != null)
                            {
                                if (double.TryParse(xValueCell.ToString(), out double x) &&
                                    double.TryParse(yValueCell.ToString(), out double y))
                                {
                                    data.Add((x, y));
                                }
                            }
                            else
                            {
                                // 빈 행이면 중단
                                break;
                            }

                            row++;
                        }
                    }
                }
            }
            else if (extension == ".xls")
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    var readerConfig = new ExcelDataReader.ExcelReaderConfiguration
                    {
                        FallbackEncoding = Encoding.GetEncoding(1252) // Windows-1252 인코딩
                    };

                    using (var reader = ExcelReaderFactory.CreateReader(stream, readerConfig))
                    {
                        // 해당 시트로 이동
                        do
                        {
                            if (reader.Name == sheetName)
                            {
                                if (isImpedanceSheet)
                                {
                                    // Impedance 시트: 4행 1열에 X값 1개, 4행 2열에 Y값 1개
                                    // 4행까지 읽기 (0-based이므로 3번 읽기)
                                    for (int i = 0; i < 3; i++)
                                    {
                                        if (!reader.Read()) return data;
                                    }

                                    // 4행 읽기 (X값과 Y값)
                                    if (!reader.Read()) return data;
                                    var xValueCell = reader.GetValue(0);
                                    var yValueCell = reader.GetValue(1);

                                    string fileName = Path.GetFileName(filePath);
                                    AddLog($"[Impedance .xls] {fileName}: X셀={xValueCell}, Y셀={yValueCell}");

                                    if (xValueCell != null && yValueCell != null)
                                    {
                                        if (double.TryParse(xValueCell.ToString(), out double x) &&
                                            double.TryParse(yValueCell.ToString(), out double y))
                                        {
                                            data.Add((x, y));
                                            AddLog($"  ✓ 성공: X={x}, Y={y}");
                                        }
                                        else
                                        {
                                            AddLog($"  ✗ 파싱 실패 (숫자가 아님)");
                                        }
                                    }
                                    else
                                    {
                                        AddLog($"  ✗ 셀이 비어있음 (null)");
                                    }
                                }
                                else
                                {
                                    // 일반 시트: 5행까지 건너뛰기 (0-based이므로 4번 읽기)
                                    for (int i = 0; i < 4; i++)
                                    {
                                        if (!reader.Read()) return data;
                                    }

                                    // 데이터 읽기
                                    while (reader.Read())
                                    {
                                        var xValueCell = reader.GetValue(0);
                                        var yValueCell = reader.GetValue(1);

                                        if (xValueCell != null && yValueCell != null)
                                        {
                                            if (double.TryParse(xValueCell.ToString(), out double x) &&
                                                double.TryParse(yValueCell.ToString(), out double y))
                                            {
                                                data.Add((x, y));
                                            }
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                        } while (reader.NextResult());
                    }
                }
            }

            return data;
        }

        private string SanitizeSheetName(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return "Sheet";
            }

            // Excel 시트 이름에서 사용 불가능한 문자 제거: : \ / ? * [ ]
            var invalidChars = new char[] { ':', '\\', '/', '?', '*', '[', ']' };
            string sanitized = sheetName;

            foreach (var c in invalidChars)
            {
                sanitized = sanitized.Replace(c, '_');
            }

            // 시트 이름 길이 제한 (31자)
            if (sanitized.Length > 31)
            {
                sanitized = sanitized.Substring(0, 31);
            }

            return sanitized;
        }

        private void CreateMergedSheet(ExcelPackage package, string sheetName, Dictionary<string, List<(double X, double Y)>> allFileData)
        {
            // 시트 이름 정리
            string sanitizedSheetName = SanitizeSheetName(sheetName);

            // 중복된 시트 이름 처리
            string uniqueSheetName = sanitizedSheetName;
            int counter = 1;
            while (package.Workbook.Worksheets[uniqueSheetName] != null)
            {
                uniqueSheetName = $"{sanitizedSheetName.Substring(0, Math.Min(28, sanitizedSheetName.Length))}_{counter}";
                counter++;
            }

            var worksheet = package.Workbook.Worksheets.Add(uniqueSheetName);
            var fileNames = allFileData.Keys.OrderBy(name => name).ToList();

            // 각 파일의 X축 값을 그룹으로 나눔 (X축이 같은 파일끼리 그룹화)
            // 소수점 2자리로 반올림하여 비교
            var groups = new List<List<string>>();
            var processedFiles = new HashSet<string>();

            foreach (var fileName in fileNames)
            {
                if (processedFiles.Contains(fileName)) continue;

                var currentGroup = new List<string> { fileName };
                processedFiles.Add(fileName);

                var currentXValues = allFileData[fileName]
                    .Select(d => Math.Round(d.X, 2))
                    .OrderBy(x => x)
                    .ToList();

                // 같은 X축을 가진 다른 파일 찾기
                foreach (var otherFileName in fileNames)
                {
                    if (processedFiles.Contains(otherFileName)) continue;

                    var otherXValues = allFileData[otherFileName]
                        .Select(d => Math.Round(d.X, 2))
                        .OrderBy(x => x)
                        .ToList();

                    if (currentXValues.SequenceEqual(otherXValues))
                    {
                        currentGroup.Add(otherFileName);
                        processedFiles.Add(otherFileName);
                    }
                }

                groups.Add(currentGroup);
            }

            // 각 그룹별로 데이터 작성
            int col = 1;
            foreach (var group in groups)
            {
                if (col > 1)
                {
                    col++; // 그룹 사이에 빈 열
                }

                // X열 헤더
                worksheet.Cells[1, col].Value = "X";

                // 첫 번째 파일의 X축 값 사용 (그룹 내 모든 파일이 같은 X축 값을 가짐)
                var firstFileName = group[0];
                var xValues = allFileData[firstFileName]
                    .OrderBy(d => d.X)
                    .Select(d => Math.Round(d.X, 2))
                    .ToList();

                // X축 데이터 작성
                for (int i = 0; i < xValues.Count; i++)
                {
                    worksheet.Cells[i + 2, col].Value = xValues[i];
                }

                col++;

                // 각 파일의 Y값 작성
                foreach (var fileName in group)
                {
                    worksheet.Cells[1, col].Value = fileName;

                    // 소수점 2자리로 반올림한 X값을 키로 사용
                    var dataDict = allFileData[fileName]
                        .OrderBy(d => d.X)
                        .ToDictionary(d => Math.Round(d.X, 2), d => d.Y);

                    for (int i = 0; i < xValues.Count; i++)
                    {
                        if (dataDict.TryGetValue(xValues[i], out double yValue))
                        {
                            worksheet.Cells[i + 2, col].Value = yValue;
                        }
                    }

                    col++;
                }
            }

            // 열 너비 자동 조정
            try
            {
                if (worksheet.Dimension != null)
                {
                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                }
            }
            catch
            {
                // AutoFitColumns 실패 시 무시 (데이터는 이미 작성됨)
            }
        }

        public void HandleWebDroppedFiles(IReadOnlyList<string> droppedPaths)
        {
            string? folderPath = droppedPaths.FirstOrDefault(Directory.Exists);
            if (!string.IsNullOrWhiteSpace(folderPath))
            {
                LoadExcelFilesFromFolder(folderPath);
                return;
            }

            string? firstExcelFile = droppedPaths.FirstOrDefault(path =>
                string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(firstExcelFile))
            {
                string? parent = Path.GetDirectoryName(firstExcelFile);
                if (!string.IsNullOrWhiteSpace(parent) && Directory.Exists(parent))
                {
                    LoadExcelFilesFromFolder(parent);
                }
            }
        }

        public object GetWebModuleSnapshot()
        {
            return new
            {
                moduleType = "GraphMakerAudioBus",
                fileCount = FileCountText.Text ?? string.Empty,
                status = StatusText.Text ?? string.Empty,
                files = excelFiles.Select(file => new
                {
                    name = file.Name,
                    fullName = file.FullName,
                    sheetCount = fileSheetMapping.TryGetValue(file, out List<string>? sheets) ? sheets.Count : 0
                }).ToArray(),
                sheetFileCounts = fileSheetMapping.Values
                    .SelectMany(sheets => sheets)
                    .GroupBy(name => name)
                    .OrderBy(group => group.Key)
                    .Select(group => new
                    {
                        name = group.Key,
                        fileCount = group.Count()
                    }).ToArray(),
                sheets = sheetCheckBoxes.Select(entry => new
                {
                    name = entry.Key,
                    isChecked = entry.Value.IsChecked == true
                }).ToArray(),
                logText = LogTextBox.Text ?? string.Empty,
                previewColumns = new[] { "File", "Sheet Count" },
                previewRows = excelFiles.Select(file => new[]
                {
                    file.Name,
                    (fileSheetMapping.TryGetValue(file, out List<string>? sheets) ? sheets.Count : 0).ToString()
                }).ToArray()
            };
        }

        public object UpdateWebModuleState(JsonElement payload)
        {
            if (payload.TryGetProperty("selectedSheets", out JsonElement selectedSheetsElement) &&
                selectedSheetsElement.ValueKind == JsonValueKind.Array)
            {
                HashSet<string> selected = selectedSheetsElement.EnumerateArray()
                    .Select(item => item.GetString() ?? string.Empty)
                    .ToHashSet(StringComparer.Ordinal);

                foreach ((string name, CheckBox checkBox) in sheetCheckBoxes)
                {
                    checkBox.IsChecked = selected.Contains(name);
                }
            }

            NotifyWebModuleSnapshotChanged();
            return GetWebModuleSnapshot();
        }

        public object InvokeWebModuleAction(string action)
        {
            switch (action)
            {
                case "select-folder":
                    FolderDropBoxControl.Browse();
                    break;
                case "extract-data":
                    ExtractDataButton_Click(this, new RoutedEventArgs());
                    break;
            }

            return GetWebModuleSnapshot();
        }

    }
}
