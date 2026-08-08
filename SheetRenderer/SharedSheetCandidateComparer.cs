using System;
using System.Collections.Generic;
using System.Text.Json.Nodes;

internal static class SharedSheetCandidateComparer
{
    public static bool AreEquivalentIgnoringExcludedColumnValues(
        SharedSheetDocument left,
        SharedSheetDocument right,
        Func<object, object> normalizeValue)
    {
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        if (left == null || right == null || normalizeValue == null)
        {
            return false;
        }

        if (!string.Equals(left.Project, right.Project, StringComparison.Ordinal) ||
            !string.Equals(left.SheetId, right.SheetId, StringComparison.Ordinal) ||
            !string.Equals(left.SheetName, right.SheetName, StringComparison.Ordinal) ||
            !string.Equals(left.RangeAddress, right.RangeAddress, StringComparison.Ordinal) ||
            !AreRangeInfosEquivalent(left.RangeInfo, right.RangeInfo))
        {
            return false;
        }

        object[] leftRowIds = left.RowIds ?? new object[0];
        object[] rightRowIds = right.RowIds ?? new object[0];
        if (leftRowIds.Length != rightRowIds.Length)
        {
            return false;
        }

        for (int row = 0; row < leftRowIds.Length; row++)
        {
            if (!AreValuesEquivalent(leftRowIds[row], rightRowIds[row], normalizeValue))
            {
                return false;
            }
        }

        object[][] leftRows = left.Values ?? new object[0][];
        object[][] rightRows = right.Values ?? new object[0][];
        if (leftRows.Length != rightRows.Length)
        {
            return false;
        }

        HashSet<int> ignoredColumns = GetIgnoredColumns(left.RangeInfo);
        for (int row = 0; row < leftRows.Length; row++)
        {
            object[] leftValues = leftRows[row] ?? new object[0];
            object[] rightValues = rightRows[row] ?? new object[0];
            if (leftValues.Length != rightValues.Length)
            {
                return false;
            }

            for (int column = 0; column < leftValues.Length; column++)
            {
                if (!ignoredColumns.Contains(column) &&
                    !AreValuesEquivalent(leftValues[column], rightValues[column], normalizeValue))
                {
                    return false;
                }
            }
        }

        return true;
    }

    private static bool AreValuesEquivalent(
        object left,
        object right,
        Func<object, object> normalizeValue)
    {
        object normalizedLeft = normalizeValue(left);
        object normalizedRight = normalizeValue(right);
        if (normalizedLeft == null || normalizedRight == null)
        {
            return normalizedLeft == null && normalizedRight == null;
        }

        JsonNode leftNode = JsonValue.Create(normalizedLeft);
        JsonNode rightNode = JsonValue.Create(normalizedRight);
        return leftNode != null &&
            rightNode != null &&
            string.Equals(leftNode.ToJsonString(), rightNode.ToJsonString(), StringComparison.Ordinal);
    }

    private static bool AreRangeInfosEquivalent(SharedRangeInfo left, SharedRangeInfo right)
    {
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        if (left == null || right == null || left.IdColumnOffset != right.IdColumnOffset)
        {
            return false;
        }

        return GetIgnoredColumns(left).SetEquals(GetIgnoredColumns(right));
    }

    private static HashSet<int> GetIgnoredColumns(SharedRangeInfo rangeInfo)
    {
        return rangeInfo == null || rangeInfo.IgnoreColumnOffsets == null
            ? new HashSet<int>()
            : new HashSet<int>(rangeInfo.IgnoreColumnOffsets);
    }
}
