using System.IO;
using System.Runtime.InteropServices;

namespace InferenceDataAIService.Wpf;

internal static class ExcelRangeNavigator
{
    internal static Task OpenReadOnlyAsync(
        string sourcePath,
        string sheetName,
        string rangeAddress) =>
        Task.Run(() => OpenReadOnly(sourcePath, sheetName, rangeAddress));

    private static void OpenReadOnly(
        string sourcePath,
        string sheetName,
        string rangeAddress)
    {
        var workbookPath = Path.GetFullPath(sourcePath);
        if (!File.Exists(workbookPath))
            throw new FileNotFoundException(
                "근거 원본 Excel 파일을 찾을 수 없습니다.",
                workbookPath);
        if (string.IsNullOrWhiteSpace(sheetName)
            || string.IsNullOrWhiteSpace(rangeAddress))
            throw new InvalidOperationException(
                "근거의 시트 또는 범위가 비어 있습니다.");

        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException(
                "Microsoft Excel COM을 사용할 수 없습니다.");
        object? applicationObject = null;
        object? workbooksObject = null;
        object? workbookObject = null;
        object? worksheetsObject = null;
        object? worksheetObject = null;
        object? rangeObject = null;
        try
        {
            applicationObject = Activator.CreateInstance(excelType)
                ?? throw new InvalidOperationException(
                    "Microsoft Excel을 시작하지 못했습니다.");
            dynamic application = applicationObject;
            application.Visible = true;
            application.DisplayAlerts = false;
            workbooksObject = application.Workbooks;
            dynamic workbooks = workbooksObject;
            workbookObject = workbooks.Open(workbookPath, ReadOnly: true);
            dynamic workbook = workbookObject;
            worksheetsObject = workbook.Worksheets;
            dynamic worksheets = worksheetsObject;
            worksheetObject = worksheets[sheetName];
            dynamic worksheet = worksheetObject;
            worksheet.Activate();
            rangeObject = worksheet.Range[rangeAddress];
            application.Goto(rangeObject, true);
            application.DisplayAlerts = true;
        }
        catch
        {
            if (applicationObject is not null)
            {
                try
                {
                    dynamic application = applicationObject;
                    application.DisplayAlerts = true;
                    application.Quit();
                }
                catch
                {
                    // Preserve the original navigation failure.
                }
            }
            throw;
        }
        finally
        {
            Release(rangeObject);
            Release(worksheetObject);
            Release(worksheetsObject);
            Release(workbookObject);
            Release(workbooksObject);
            Release(applicationObject);
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
            Marshal.FinalReleaseComObject(value);
    }
}
