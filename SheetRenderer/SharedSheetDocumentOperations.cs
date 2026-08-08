using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

internal static class SharedSheetDiffBuilder
{
    public static List<SharedSheetDiffEntry> BuildEntries(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument,
        SharedSheetDocument displayDocument,
        Func<object, object, bool> valuesEqual,
        Func<object, object> normalizeValue)
    {
        var entries = new List<SharedSheetDiffEntry>();
        if (!SharedSheetRowOperations.CanUseRowIds(localDocument) ||
            valuesEqual == null ||
            normalizeValue == null)
        {
            return entries;
        }

        int columnCount = Math.Max(
            SharedSheetRowOperations.GetColumnCount(localDocument),
            Math.Max(
                SharedSheetRowOperations.GetColumnCount(baseDocument),
                SharedSheetRowOperations.GetColumnCount(remoteDocument)));
        HashSet<int> ignoredColumns = SharedSheetRowOperations.GetIgnoredColumns(localDocument);
        Dictionary<string, object[]> localRows = SharedSheetRowOperations.CreateRowMap(localDocument);
        Dictionary<string, object[]> baseRows = SharedSheetRowOperations.CreateRowMap(baseDocument);
        Dictionary<string, object[]> remoteRows = remoteDocument == null
            ? new Dictionary<string, object[]>(baseRows, StringComparer.Ordinal)
            : SharedSheetRowOperations.CreateRowMap(remoteDocument);
        List<string> rowOrder = SharedSheetRowOperations.BuildRowOrder(
            localDocument,
            remoteDocument,
            baseDocument);

        SharedSheetDocument displaySource = displayDocument ?? localDocument;
        Dictionary<string, int> displayRows = CreateDisplayRowMap(displaySource);
        int? startColumn = TryGetStartColumn(displaySource == null ? null : displaySource.RangeAddress);
        bool hasExplicitRemote = remoteDocument != null &&
            !ReferenceEquals(remoteDocument, baseDocument);

        foreach (string rowId in rowOrder)
        {
            object[] localRow;
            bool hasLocalRow = localRows.TryGetValue(rowId, out localRow);
            object[] baseRow;
            bool hasBaseRow = baseRows.TryGetValue(rowId, out baseRow);
            object[] remoteRow;
            bool hasRemoteRow = remoteRows.TryGetValue(rowId, out remoteRow);

            bool localRowDeleted = hasBaseRow && !hasLocalRow;
            bool remoteRowDeleted = hasBaseRow && hasExplicitRemote && !hasRemoteRow;
            if (localRowDeleted || remoteRowDeleted)
            {
                bool localRowChanged = hasLocalRow &&
                    !RowsEqual(baseRow, localRow, columnCount, ignoredColumns, valuesEqual);
                bool remoteRowChanged = hasRemoteRow &&
                    !RowsEqual(baseRow, remoteRow, columnCount, ignoredColumns, valuesEqual);

                string stateLabel;
                if (localRowDeleted && remoteRowDeleted)
                {
                    stateLabel = "双方で行削除";
                }
                else if (localRowDeleted && remoteRowChanged)
                {
                    stateLabel = "競合（行削除 vs 編集）";
                }
                else if (localRowDeleted)
                {
                    stateLabel = "行削除";
                }
                else if (localRowChanged)
                {
                    stateLabel = "競合（編集 vs 行削除）";
                }
                else
                {
                    stateLabel = "共有先で行削除";
                }

                entries.Add(new SharedSheetDiffEntry
                {
                    SheetId = localDocument.SheetId,
                    SheetName = localDocument.SheetName,
                    RowId = rowId,
                    CellAddress = null,
                    StateLabel = stateLabel,
                    IsRowDeletion = localRowDeleted,
                    IsRowLevelChange = true,
                    BaseRowState = "存在",
                    LocalRowState = hasLocalRow
                        ? (localRowChanged ? "存在（編集あり）" : "存在")
                        : "削除",
                    RemoteRowState = hasExplicitRemote
                        ? (hasRemoteRow
                            ? (remoteRowChanged ? "存在（編集あり）" : "存在")
                            : "削除")
                        : null,
                    HasRemoteValue = hasExplicitRemote
                });
                continue;
            }

            for (int column = 0; column < columnCount; column++)
            {
                if (ignoredColumns.Contains(column))
                {
                    continue;
                }

                object baseValue = SharedSheetRowOperations.GetCellValue(baseRow, column);
                object localValue = SharedSheetRowOperations.GetCellValue(localRow, column);
                object remoteValue = remoteRow == null
                    ? baseValue
                    : SharedSheetRowOperations.GetCellValue(remoteRow, column);

                if (valuesEqual(baseValue, localValue) && valuesEqual(localValue, remoteValue))
                {
                    continue;
                }

                int displayRow;
                bool hasDisplayRow = displayRows.TryGetValue(rowId, out displayRow);
                string displayAddress = hasDisplayRow && startColumn.HasValue
                    ? GetColumnLetter(startColumn.Value + column) + displayRow
                    : "?";

                entries.Add(new SharedSheetDiffEntry
                {
                    SheetId = localDocument.SheetId,
                    SheetName = localDocument.SheetName,
                    RowId = rowId,
                    CellAddress = displayAddress,
                    StateLabel = BuildCellStateLabel(baseValue, localValue, remoteValue, valuesEqual),
                    BaseValue = normalizeValue(baseValue),
                    LocalValue = normalizeValue(localValue),
                    RemoteValue = normalizeValue(remoteValue),
                    HasRemoteValue = remoteDocument != null &&
                        !ReferenceEquals(remoteDocument, baseDocument) &&
                        !valuesEqual(remoteValue, baseValue)
                });
            }
        }

        return entries;
    }

    private static bool RowsEqual(
        object[] left,
        object[] right,
        int columnCount,
        HashSet<int> ignoredColumns,
        Func<object, object, bool> valuesEqual)
    {
        if (left == null || right == null)
        {
            return left == null && right == null;
        }

        for (int column = 0; column < columnCount; column++)
        {
            if (!ignoredColumns.Contains(column) &&
                !valuesEqual(
                    SharedSheetRowOperations.GetCellValue(left, column),
                    SharedSheetRowOperations.GetCellValue(right, column)))
            {
                return false;
            }
        }

        return true;
    }

    private static string BuildCellStateLabel(
        object baseValue,
        object localValue,
        object remoteValue,
        Func<object, object, bool> valuesEqual)
    {
        if (valuesEqual(localValue, baseValue) && valuesEqual(remoteValue, baseValue))
        {
            return "変更なし";
        }

        if (valuesEqual(localValue, baseValue))
        {
            return "共有先変更";
        }

        if (valuesEqual(remoteValue, baseValue))
        {
            return "ローカル変更";
        }

        return valuesEqual(localValue, remoteValue) ? "同一変更" : "競合";
    }

    private static Dictionary<string, int> CreateDisplayRowMap(SharedSheetDocument document)
    {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        if (!SharedSheetRowOperations.CanUseRowIds(document))
        {
            return result;
        }

        int? startRow = TryGetStartRow(document.RangeAddress);
        if (!startRow.HasValue)
        {
            return result;
        }

        for (int index = 0; index < document.RowIds.Length; index++)
        {
            string rowId = SharedSheetRowOperations.NormalizeRowId(document.RowIds[index]);
            if (!string.IsNullOrWhiteSpace(rowId) && !result.ContainsKey(rowId))
            {
                result[rowId] = startRow.Value + index;
            }
        }

        return result;
    }

    private static int? TryGetStartRow(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            return null;
        }

        Match match = Regex.Match(rangeAddress, @"\$?[A-Za-z]+\$?(\d+)");
        int value;
        return match.Success && int.TryParse(match.Groups[1].Value, out value)
            ? (int?)value
            : null;
    }

    private static int? TryGetStartColumn(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
        {
            return null;
        }

        Match match = Regex.Match(rangeAddress, @"\$?([A-Za-z]+)\$?\d+");
        if (!match.Success)
        {
            return null;
        }

        int column = 0;
        foreach (char character in match.Groups[1].Value.ToUpperInvariant())
        {
            if (character < 'A' || character > 'Z')
            {
                return null;
            }

            column = (column * 26) + (character - 'A' + 1);
        }

        return column > 0 ? (int?)column : null;
    }

    private static string GetColumnLetter(int column)
    {
        if (column <= 0)
        {
            return "?";
        }

        string result = string.Empty;
        while (column > 0)
        {
            int remainder = (column - 1) % 26;
            result = (char)('A' + remainder) + result;
            column = (column - remainder - 1) / 26;
        }

        return result;
    }
}

internal sealed class SharedSheetUploadMergeResult
{
    public SharedSheetDocument MergedDocument { get; set; }
    public int ConflictCount { get; set; }
}

internal static class SharedSheetUploadMergeEngine
{
    public static SharedSheetUploadMergeResult Merge(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument,
        Func<object, object, bool> valuesEqual,
        Func<object, object> normalizeValue)
    {
        if (localDocument == null)
        {
            return null;
        }

        bool canMerge =
            SharedSheetRowOperations.CanUseRowIds(localDocument) &&
            (baseDocument == null || SharedSheetRowOperations.CanUseRowIds(baseDocument)) &&
            (remoteDocument == null || SharedSheetRowOperations.CanUseRowIds(remoteDocument)) &&
            valuesEqual != null &&
            normalizeValue != null;
        if (!canMerge)
        {
            return new SharedSheetUploadMergeResult
            {
                MergedDocument = localDocument,
                ConflictCount = 1
            };
        }

        int columnCount = Math.Max(
            SharedSheetRowOperations.GetColumnCount(localDocument),
            Math.Max(
                SharedSheetRowOperations.GetColumnCount(baseDocument),
                SharedSheetRowOperations.GetColumnCount(remoteDocument)));
        HashSet<int> ignoredColumns = SharedSheetRowOperations.GetIgnoredColumns(localDocument);
        Dictionary<string, object[]> localRows = SharedSheetRowOperations.CreateRowMap(localDocument);
        Dictionary<string, object[]> baseRows = SharedSheetRowOperations.CreateRowMap(baseDocument);
        Dictionary<string, object[]> remoteRows = remoteDocument == null
            ? new Dictionary<string, object[]>(baseRows, StringComparer.Ordinal)
            : SharedSheetRowOperations.CreateRowMap(remoteDocument);
        List<string> rowOrder = SharedSheetRowOperations.BuildRowOrder(
            localDocument,
            remoteDocument,
            baseDocument);
        var mergedRows = new List<object[]>();
        var mergedRowIds = new List<object>();
        int conflictCount = 0;

        foreach (string rowId in rowOrder)
        {
            object[] baseRow;
            bool hasBaseRow = baseRows.TryGetValue(rowId, out baseRow);
            object[] localRow;
            bool hasLocalRow = localRows.TryGetValue(rowId, out localRow);
            object[] remoteRow;
            bool hasRemoteRow = remoteRows.TryGetValue(rowId, out remoteRow);

            if (hasBaseRow && !hasLocalRow && !hasRemoteRow)
            {
                continue;
            }

            if (hasBaseRow && !hasLocalRow)
            {
                if (RowsEqual(baseRow, remoteRow, columnCount, ignoredColumns, valuesEqual))
                {
                    continue;
                }

                AddRow(mergedRows, mergedRowIds, rowId, remoteRow, columnCount, normalizeValue);
                conflictCount++;
                continue;
            }

            if (hasBaseRow && !hasRemoteRow)
            {
                if (RowsEqual(baseRow, localRow, columnCount, ignoredColumns, valuesEqual))
                {
                    continue;
                }

                AddRow(mergedRows, mergedRowIds, rowId, localRow, columnCount, normalizeValue);
                conflictCount++;
                continue;
            }

            if (!hasBaseRow && !hasLocalRow && hasRemoteRow)
            {
                AddRow(mergedRows, mergedRowIds, rowId, remoteRow, columnCount, normalizeValue);
                continue;
            }

            if (!hasBaseRow && hasLocalRow && !hasRemoteRow)
            {
                AddRow(mergedRows, mergedRowIds, rowId, localRow, columnCount, normalizeValue);
                continue;
            }

            if (!hasLocalRow && !hasRemoteRow)
            {
                continue;
            }

            var mergedRow = new object[columnCount];
            for (int column = 0; column < columnCount; column++)
            {
                object baseValue = SharedSheetRowOperations.GetCellValue(baseRow, column);
                object localValue = SharedSheetRowOperations.GetCellValue(localRow, column);
                object remoteValue = SharedSheetRowOperations.GetCellValue(remoteRow, column);

                if (ignoredColumns.Contains(column))
                {
                    mergedRow[column] = normalizeValue(localValue);
                }
                else if (valuesEqual(localValue, baseValue))
                {
                    mergedRow[column] = normalizeValue(remoteValue);
                }
                else if (valuesEqual(remoteValue, baseValue) || valuesEqual(localValue, remoteValue))
                {
                    mergedRow[column] = normalizeValue(localValue);
                }
                else
                {
                    mergedRow[column] = normalizeValue(localValue);
                    conflictCount++;
                }
            }

            mergedRows.Add(mergedRow);
            mergedRowIds.Add(rowId);
        }

        return new SharedSheetUploadMergeResult
        {
            MergedDocument = new SharedSheetDocument
            {
                Project = localDocument.Project,
                SheetId = localDocument.SheetId,
                SheetName = localDocument.SheetName,
                RangeAddress = localDocument.RangeAddress,
                RangeInfo = CloneRangeInfo(localDocument.RangeInfo),
                RowIds = mergedRowIds.ToArray(),
                Values = mergedRows.ToArray()
            },
            ConflictCount = conflictCount
        };
    }

    private static bool RowsEqual(
        object[] left,
        object[] right,
        int columnCount,
        HashSet<int> ignoredColumns,
        Func<object, object, bool> valuesEqual)
    {
        if (left == null || right == null)
        {
            return left == null && right == null;
        }

        for (int column = 0; column < columnCount; column++)
        {
            if (!ignoredColumns.Contains(column) &&
                !valuesEqual(
                    SharedSheetRowOperations.GetCellValue(left, column),
                    SharedSheetRowOperations.GetCellValue(right, column)))
            {
                return false;
            }
        }

        return true;
    }

    private static void AddRow(
        List<object[]> rows,
        List<object> rowIds,
        string rowId,
        object[] source,
        int columnCount,
        Func<object, object> normalizeValue)
    {
        var row = new object[columnCount];
        for (int column = 0; column < columnCount; column++)
        {
            row[column] = normalizeValue(SharedSheetRowOperations.GetCellValue(source, column));
        }

        rows.Add(row);
        rowIds.Add(rowId);
    }

    private static SharedRangeInfo CloneRangeInfo(SharedRangeInfo source)
    {
        return source == null
            ? null
            : new SharedRangeInfo
            {
                IdColumnOffset = source.IdColumnOffset,
                IgnoreColumnOffsets = source.IgnoreColumnOffsets == null
                    ? new HashSet<int>()
                    : new HashSet<int>(source.IgnoreColumnOffsets)
            };
    }
}

internal static class SharedSheetRowOperations
{
    public static bool CanUseRowIds(SharedSheetDocument document)
    {
        if (document == null || document.Values == null || document.RowIds == null ||
            document.RowIds.Length != document.Values.Length)
        {
            return false;
        }

        if (document.RowIds.Length == 0)
        {
            return true;
        }

        return document.RowIds.Any(value => !string.IsNullOrWhiteSpace(NormalizeRowId(value)));
    }

    public static string NormalizeRowId(object value)
    {
        if (value == null)
        {
            return null;
        }

        string text = value.ToString();
        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    public static int GetColumnCount(SharedSheetDocument document)
    {
        return document == null || document.Values == null || document.Values.Length == 0
            ? 0
            : document.Values.Max(row => row == null ? 0 : row.Length);
    }

    public static object GetCellValue(object[] row, int column)
    {
        return row == null || column < 0 || column >= row.Length ? null : row[column];
    }

    public static HashSet<int> GetIgnoredColumns(SharedSheetDocument document)
    {
        return document == null || document.RangeInfo == null ||
            document.RangeInfo.IgnoreColumnOffsets == null
            ? new HashSet<int>()
            : new HashSet<int>(document.RangeInfo.IgnoreColumnOffsets);
    }

    public static Dictionary<string, object[]> CreateRowMap(SharedSheetDocument document)
    {
        var result = new Dictionary<string, object[]>(StringComparer.Ordinal);
        if (!CanUseRowIds(document))
        {
            return result;
        }

        for (int index = 0; index < document.RowIds.Length; index++)
        {
            string rowId = NormalizeRowId(document.RowIds[index]);
            if (!string.IsNullOrWhiteSpace(rowId))
            {
                result[rowId] = document.Values[index] ?? new object[0];
            }
        }

        return result;
    }

    public static List<string> BuildRowOrder(params SharedSheetDocument[] documents)
    {
        var result = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (SharedSheetDocument document in documents ?? new SharedSheetDocument[0])
        {
            if (!CanUseRowIds(document))
            {
                continue;
            }

            foreach (object value in document.RowIds)
            {
                string rowId = NormalizeRowId(value);
                if (!string.IsNullOrWhiteSpace(rowId) && seen.Add(rowId))
                {
                    result.Add(rowId);
                }
            }
        }

        return result;
    }
}
