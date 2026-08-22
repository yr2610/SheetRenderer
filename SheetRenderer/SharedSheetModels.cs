using System;
using System.Collections.Generic;

internal sealed class SharedSheetDocument
{
    public string Project { get; set; }
    public string SheetId { get; set; }
    public string SheetName { get; set; }
    public string RangeAddress { get; set; }
    public SharedRangeInfo RangeInfo { get; set; }
    public object[] RowIds { get; set; }
    public object[][] Values { get; set; }
    public string Hash { get; set; }
}

internal sealed class SharedSheetDiffEntry
{
    public string SheetId { get; set; }
    public string SheetName { get; set; }
    public string RowId { get; set; }
    public string CellAddress { get; set; }
    public string StateLabel { get; set; }
    public object BaseValue { get; set; }
    public object LocalValue { get; set; }
    public object RemoteValue { get; set; }
    public bool HasRemoteValue { get; set; }
    public bool IsRowDeletion { get; set; }
    public bool IsRowLevelChange { get; set; }
    public string BaseRowState { get; set; }
    public string LocalRowState { get; set; }
    public string RemoteRowState { get; set; }

    public string CellAddressText
    {
        get { return string.IsNullOrWhiteSpace(CellAddress) ? "-" : CellAddress; }
    }

    public string BaseText
    {
        get { return IsRowLevelChange ? (BaseRowState ?? string.Empty) : FormatValue(BaseValue); }
    }

    public string LocalText
    {
        get { return IsRowLevelChange ? (LocalRowState ?? string.Empty) : FormatValue(LocalValue); }
    }

    public string RemoteText
    {
        get
        {
            if (!HasRemoteValue)
            {
                return string.Empty;
            }

            return IsRowLevelChange ? (RemoteRowState ?? string.Empty) : FormatValue(RemoteValue);
        }
    }

    private static string FormatValue(object value)
    {
        if (value == null || value == DBNull.Value)
        {
            return "(null)";
        }

        if (value is DateTime dateTimeValue)
        {
            return dateTimeValue.ToString("yyyy-MM-dd HH:mm:ss");
        }

        return value.ToString();
    }
}

internal sealed class SharedRangeInfo
{
    public int? IdColumnOffset { get; set; }
    public HashSet<int> IgnoreColumnOffsets { get; set; }
}

internal sealed class SharedProjectManifest
{
    public string Project { get; set; }
    public string UpdatedAt { get; set; }
    public List<SharedProjectManifestEntry> Sheets { get; set; }
}

internal sealed class SharedProjectManifestEntry
{
    public string SheetId { get; set; }
    public string SheetName { get; set; }
    public string Hash { get; set; }
}

internal sealed class SharedSheetSyncState
{
    public List<SharedSheetSyncStateEntry> Sheets { get; set; }
}

internal sealed class SharedSheetSyncStateEntry
{
    public string SheetId { get; set; }
    public string BaseHash { get; set; }
}
