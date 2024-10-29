﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        Excel.Range destinationRange = worksheet.Rows[startRow].Resize[rowCount];
        sourceRange.Copy(destinationRange);
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

}
