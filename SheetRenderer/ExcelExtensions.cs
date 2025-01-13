using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

using ExcelDna.Integration;

using ExcelDna.Integration.CustomUI;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;


public static class ExcelExtensions
{

    public static void AutoFitColumnsIfNarrower(this Excel.Range range)
    {
        Excel.Worksheet sheet = range.Worksheet;

        // 各列の元の幅を保存
        double[] originalWidths = new double[range.Columns.Count];
        int index = 0;
        foreach (Excel.Range column in range.Columns)
        {
            originalWidths[index++] = column.ColumnWidth;
        }

        // 一度にAutoFitを適用
        range.Columns.AutoFit();

        // 各列の幅をチェックして、元の幅よりも広くなった場合は元に戻す
        index = 0;
        foreach (Excel.Range column in range.Columns)
        {
            if (column.ColumnWidth > originalWidths[index])
            {
                column.ColumnWidth = originalWidths[index];
            }
            index++;
        }
    }

    public static Excel.Range GetRange(this Excel.Worksheet sheet, int startRow, int startColumn, int rowCount, int columnCount)
    {
        Excel.Range startCell = (Excel.Range)sheet.Cells[startRow, startColumn];
        Excel.Range endCell = (Excel.Range)sheet.Cells[startRow + rowCount - 1, startColumn + columnCount - 1];
        Excel.Range range = sheet.Range[startCell, endCell];
        return range;
    }

    public static object[,] GetValuesAs2DArray(object range)
    {
        if (range is object[,] array)
        {
            // 既に配列の場合はそのまま返す
            return array;
        }
        else if (range is object singleValue)
        {
            // 1つのセルの場合、1-originのように見える2次元配列として返す
            // 実際の配列のサイズは1x1
            var result = Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
            result.SetValue(singleValue, 1, 1);
            return (object[,])result;
        }

        // 何もない場合は空の1x1の2次元配列を返す
        var emptyResult = Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
        emptyResult.SetValue(null, 1, 1);
        return (object[,])emptyResult;
    }

    public static IEnumerable<object> GetColumnWithOffset(this Excel.Worksheet worksheet, string address, int columnOffset)
    {
        // 指定されたアドレスの範囲を取得
        var range = worksheet.Range[address];

        // 範囲の開始列を取得
        int startColumn = range.Column;

        // オフセット後の列番号を計算
        int targetColumn = startColumn + columnOffset;

        // 指定された範囲の行を基準にして、対象列を取得
        var offsetColumn = worksheet.Range[worksheet.Cells[range.Row, targetColumn], worksheet.Cells[range.Row + range.Rows.Count - 1, targetColumn]];

        // 2次元配列として範囲を取得
        var values = GetValuesAs2DArray(offsetColumn.Value2);

        // 2次元配列をList<object>に変換
        var result = new List<object>();
        for (int i = 1; i <= values.GetLength(0); i++)
        {
            result.Add(values[i, 1]);
        }

        return result;
    }

    // columnIndex は 0-origin
    public static IEnumerable<object> GetColumnValues(this Excel.Range range, int columnIndex)
    {
        // 指定された列インデックスが範囲外の場合に対応
        int totalColumns = range.Columns.Count;
        Excel.Range columnRange;

        if (columnIndex >= totalColumns)
        {
            // 範囲外の列インデックスの場合、指定された列を元の範囲の行に合わせて取得
            int offsetColumns = columnIndex - totalColumns + 1;
            columnRange = range.Offset[0, offsetColumns].Resize[range.Rows.Count, 1];
        }
        else
        {
            // 指定された列の範囲を取得
            columnRange = range.Columns[columnIndex + 1];
        }

        object value = columnRange.Value2;
        if (value is object[,])
        {
            object[,] values = (object[,])value;
            for (int i = 1; i <= values.GetLength(0); i++)
            {
                yield return values[i, 1];
            }
        }
        else if (value != null)
        {
            yield return value;
        }
    }

    // 指定されたワークシートオブジェクトと列アドレスから列インデックスを取得する
    public static int ColumnAddressToIndex(this Excel.Worksheet worksheet, string columnAddress)
    {
        Excel.Range range = worksheet.Range[columnAddress + "1"];
        return range.Column;
    }

    public static void DeleteRows(this Excel.Worksheet worksheet, int startRow, int rowCount)
    {
        Excel.Range rows = worksheet.Rows[startRow].Resize[rowCount];
        rows.Delete();
    }

    public static void InsertRowsAndCopyFormulas(this Excel.Worksheet worksheet, int startRow, int rowCount)
    {
        // 行を挿入
        Excel.Range insertRange = worksheet.Rows[startRow].Resize[rowCount];
        insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

        // 挿入した行の上の行から数式をコピー
        Excel.Range sourceRange = worksheet.Rows[startRow - 1];

        // 1行目にフォーマットをコピー
        Excel.Range firstRowDestination = worksheet.Rows[startRow];
        sourceRange.Copy(firstRowDestination);

        // シートの使用範囲を取得
        Excel.Range sheetUsedRange = worksheet.UsedRange;

        // UsedRange の列範囲を処理
        int startColumn = sheetUsedRange.Column;
        int endColumn = sheetUsedRange.Column + sheetUsedRange.Columns.Count - 1;

        // 1行目に数式を持たないセルを空欄にする
        for (int column = startColumn; column <= endColumn; column++)
        {
            Excel.Range sourceCell = sourceRange.Cells[1, column];
            if (!sourceCell.HasFormula) // 数式がない場合
            {
                firstRowDestination.Cells[1, column].Value = null;
            }
        }

        // 2行目以降がある場合、1行目からコピー
        if (rowCount > 1)
        {
            // firstRowDestination を1行下にずらし、サイズを設定
            Excel.Range secondToLastRange = firstRowDestination.Offset[1].Resize[rowCount - 1];
            firstRowDestination.Copy(secondToLastRange);
        }

        UpdateNamedRanges(worksheet, startRow, rowCount);
    }

    // 名前付き範囲を自動的に検出し、必要であれば更新
    private static void UpdateNamedRanges(Excel.Worksheet worksheet, int startRow, int rowCount)
    {
        foreach (Excel.Name name in worksheet.Names)
        {
            Excel.Range range = name.RefersToRange;
            int rangeLastRow = range.Row + range.Rows.Count - 1;

            // 名前付き範囲が自動で拡張されないケースに対応
            // 挿入位置が名前付き範囲の最終行に一致するかチェック
            if (startRow == rangeLastRow)
            {
                // 名前付き範囲を更新
                int newLastRow = rangeLastRow + rowCount; // 追加された行数だけ最終行をシフト

                // 新しい範囲を計算して更新
                int rangeLastColumn = range.Column + range.Columns.Count - 1;
                var newRange = worksheet.Range[range.Cells[1, 1], worksheet.Cells[newLastRow, rangeLastColumn]];
                worksheet.Names.Item(name.Name).RefersTo = "=" + worksheet.Name + "!" + newRange.Address;
            }
        }
    }

    // 指定した範囲のセルに同じ値を代入します
    public static void SetRangeValue<T>(this Excel.Worksheet sheet, int startRow, int startColumn, int rowCount, int columnCount, T value)
    {
        Excel.Range startCell = (Excel.Range)sheet.Cells[startRow, startColumn];
        Excel.Range range = startCell.Resize[rowCount, columnCount];

        range.Value = value;
    }

    public static void SetValueInSheet<T>(this Excel.Worksheet sheet, int startRow, int startColumn, T[,] array)
    {
        Excel.Range range = sheet.Cells[startRow, startColumn] as Excel.Range;
        range = range.Resize[array.GetLength(0), array.GetLength(1)];
        range.Value = array;
    }

    public static void SetValueInSheet<T>(this Excel.Worksheet sheet, int startRow, int startColumn, IEnumerable<T> source, bool isRow = true)
    {
        int length = source.Count();
        T[,] array = new T[isRow ? 1 : length, isRow ? length : 1];
        int index = 0;

        foreach (var item in source)
        {
            if (isRow)
            {
                array[0, index] = item;
            }
            else
            {
                array[index, 0] = item;
            }
            index++;
        }

        SetValueInSheet(sheet, startRow, startColumn, array);
    }

    // リストの値を行として貼り付けます
    public static void SetValueInSheetAsRow<T>(this Excel.Worksheet sheet, int startRow, int startColumn, IEnumerable<T> source)
    {
        SetValueInSheet(sheet, startRow, startColumn, source, true);
    }

    // リストの値を列として貼り付けます
    public static void SetValueInSheetAsColumn<T>(this Excel.Worksheet sheet, int startRow, int startColumn, IEnumerable<T> source)
    {
        SetValueInSheet(sheet, startRow, startColumn, source, false);
    }

    public static Excel.Worksheet GetSheetIfExists(this Excel.Workbook workbook, string sheetName)
    {
        foreach (Excel.Worksheet sheet in workbook.Sheets)
        {
            if (sheet.Name == sheetName)
            {
                return sheet;
            }
        }
        return null;
    }

    public static string GenerateTempSheetName(this Excel.Workbook workbook)
    {
        string tempName = "TempSheetName";
        int counter = 1;

        // 一時的な名前が既存のシート名と重複しないようにチェック
        while (workbook.GetSheetIfExists(tempName) != null)
        {
            tempName = "TempSheetName" + counter;
            counter++;
        }

        return tempName;
    }

    public static Excel.Name GetNamedRange(this Excel.Worksheet sheet, string name)
    {
        try
        {
            Excel.Name namedRange = sheet.Names.Item(name);
            return namedRange;
        }
        catch (Exception)
        {
            return null; // エラーが発生した場合は null を返します
        }
    }

    public static void DeleteCustomProperty(this Excel.Worksheet sheet, string propertyName)
    {
        var customProperties = sheet.CustomProperties;
        foreach (Excel.CustomProperty property in customProperties)
        {
            if (property.Name == propertyName)
            {
                property.Delete();  // 既存のプロパティを削除
                return;
            }
        }
        // 既存のプロパティが見つからない場合は何もしない
    }

    public static void SetCustomProperty(this Excel.Worksheet sheet, string propertyName, string propertyValue)
    {
        if (string.IsNullOrEmpty(propertyValue))
        {
            sheet.DeleteCustomProperty(propertyName);
            return;
        }

        // プロパティ値がnullまたは空でない場合、新規追加または更新
        var customProperties = sheet.CustomProperties;
        foreach (Excel.CustomProperty property in customProperties)
        {
            if (property.Name == propertyName)
            {
                property.Value = propertyValue;
                return;
            }
        }
        customProperties.Add(propertyName, propertyValue);
    }

    public static string GetCustomProperty(this Excel.Worksheet sheet, string propertyName)
    {
        var customProperties = sheet.CustomProperties;
        foreach (Excel.CustomProperty property in customProperties)
        {
            if (property.Name == propertyName)
            {
                return property.Value.ToString();
            }
        }
        return null;
    }

    public static dynamic GetCustomPropertyObject(this Excel.Workbook workbook, string propertyName)
    {
        dynamic properties = workbook.CustomDocumentProperties;
        foreach (dynamic prop in properties)
        {
            if (prop.Name == propertyName)
            {
                return prop;
            }
        }
        return null;
    }

    public static void SetCustomProperty(this Excel.Workbook workbook, string propertyName, string propertyValue)
    {
        dynamic prop = GetCustomPropertyObject(workbook, propertyName);
        if (prop != null)
        {
            prop.Value = propertyValue;
        }
        else
        {
            workbook.CustomDocumentProperties.Add(propertyName, false, Office.MsoDocProperties.msoPropertyTypeString, propertyValue);
        }
    }

    public static void SetCustomProperty<T>(this Excel.Workbook workbook, string propertyName, T propertyValue)
    {
        var serializer = new SerializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .Build();
        string yaml = serializer.Serialize(propertyValue);

        SetCustomProperty(workbook, propertyName, yaml);
    }

    public static string GetCustomProperty(this Excel.Workbook workbook, string propertyName)
    {
        dynamic prop = GetCustomPropertyObject(workbook, propertyName);
        return prop != null ? prop.Value : null;
    }

    public static T GetCustomProperty<T>(this Excel.Workbook workbook, string propertyName)
    {
        string yaml = GetCustomProperty(workbook, propertyName);
        if (yaml != null)
        {
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .Build();
            return deserializer.Deserialize<T>(yaml);
        }
        return default(T);
    }

    public static (int row, int column)? GetActiveCellPosition(this Excel.Application excelApp)
    {
        Excel.Range activeCell = null;

        try
        {
            activeCell = excelApp.ActiveCell;
        }
        catch (COMException)
        {
            // ActiveCell が存在しない場合のエラーハンドリング
            return null;
        }

        if (activeCell != null)
        {
            return (activeCell.Row, activeCell.Column);
        }

        return null;
    }

    public static void SetActiveCellPosition(this Excel.Application excelApp, (int row, int column)? position)
    {
        if (position.HasValue)
        {
            Excel.Worksheet activeSheet = excelApp.ActiveSheet;
            activeSheet.Cells[position.Value.row, position.Value.column].Select();
        }
    }

    public static (int horizontalScroll, int verticalScroll)? GetScrollPosition(this Excel.Application excelApp)
    {
        Excel.Window activeWindow = excelApp.ActiveWindow;
        return (activeWindow.ScrollColumn, activeWindow.ScrollRow);
    }

    public static void SetScrollPosition(this Excel.Application excelApp, (int horizontalScroll, int verticalScroll)? scrollPosition)
    {
        if (scrollPosition.HasValue)
        {
            Excel.Window activeWindow = excelApp.ActiveWindow;
            activeWindow.ScrollColumn = scrollPosition.Value.horizontalScroll;
            activeWindow.ScrollRow = scrollPosition.Value.verticalScroll;
        }
    }

    public static double GetActiveSheetZoom(this Excel.Application excelApp)
    {
        Excel.Worksheet activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;
        return activeSheet.Application.ActiveWindow.Zoom;
    }

    public static void SetActiveSheetZoom(this Excel.Application excelApp, double zoomLevel)
    {
        Excel.Worksheet activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;
        activeSheet.Application.ActiveWindow.Zoom = zoomLevel;
    }

}
