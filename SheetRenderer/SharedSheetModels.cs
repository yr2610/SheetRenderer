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
