using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

internal static class ExcelSelectionHelper
{
    public static void QueueSelectCell(string sheetName, string cellAddress, string dialogTitle)
    {
        if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(cellAddress))
        {
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            Excel.Worksheet sheet = null;
            Excel.Range range = null;
            try
            {
                var excelApp = (Excel.Application)ExcelDnaUtil.Application;
                sheet = excelApp.Sheets[sheetName] as Excel.Worksheet;
                range = sheet.Range[cellAddress];
                sheet.Activate();
                SelectRangeWithComfortableScroll(excelApp, range);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"セルの選択に失敗しました。\n\nシート: {sheetName}\nセル: {cellAddress}\n\n{ex.Message}",
                    string.IsNullOrWhiteSpace(dialogTitle) ? "セル選択" : dialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                ReleaseExcelComObject(range);
                ReleaseExcelComObject(sheet);
            }
        });
    }

    private static void SelectRangeWithComfortableScroll(Excel.Application excelApp, Excel.Range range)
    {
        Excel.Window activeWindow = null;
        Excel.Range visibleRange = null;
        Excel.Range visibleRows = null;
        Excel.Range leftVisibleRange = null;
        Excel.Range leftVisibleColumns = null;

        try
        {
            range.Select();

            activeWindow = excelApp.ActiveWindow;
            if (activeWindow == null)
            {
                return;
            }

            visibleRange = activeWindow.VisibleRange;
            visibleRows = visibleRange.Rows as Excel.Range;
            int visibleRowCount = Math.Max(1, visibleRows.Count);

            int targetRow = range.Row;
            int targetColumn = range.Column;

            activeWindow.ScrollRow = Math.Max(1, targetRow - (visibleRowCount / 2));
            activeWindow.ScrollColumn = 1;

            leftVisibleRange = activeWindow.VisibleRange;
            leftVisibleColumns = leftVisibleRange.Columns as Excel.Range;
            int visibleColumnCount = Math.Max(1, leftVisibleColumns.Count);

            if (targetColumn > visibleColumnCount)
            {
                activeWindow.ScrollColumn = Math.Max(1, targetColumn - visibleColumnCount + 1);
            }

            range.Select();
        }
        finally
        {
            ReleaseExcelComObject(leftVisibleColumns);
            ReleaseExcelComObject(leftVisibleRange);
            ReleaseExcelComObject(visibleRows);
            ReleaseExcelComObject(visibleRange);
            ReleaseExcelComObject(activeWindow);
        }
    }

    private static void ReleaseExcelComObject(object comObject)
    {
        if (comObject == null)
        {
            return;
        }

        try
        {
            if (Marshal.IsComObject(comObject))
            {
                Marshal.ReleaseComObject(comObject);
            }
        }
        catch
        {
        }
    }
}
