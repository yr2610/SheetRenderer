using System;
using System.Collections.Generic;

internal static class SharedSheetCandidateComparerTests
{
    private static int Main()
    {
        AssertEquivalent("identical", CreateDocument(), CreateDocument(), true);

        SharedSheetDocument sharedCellChanged = CreateDocument();
        sharedCellChanged.Values[0][0] = "changed";
        AssertEquivalent("shared cell change", sharedCellChanged, CreateDocument(), false);

        SharedSheetDocument ignoredCellChanged = CreateDocument();
        ignoredCellChanged.Values[0][1] = "local-only";
        AssertEquivalent("ignored cell change", ignoredCellChanged, CreateDocument(), true);

        SharedSheetDocument sharedCellTypeChanged = CreateDocument();
        sharedCellTypeChanged.Values[0][0] = 1d;
        SharedSheetDocument sharedCellTextValue = CreateDocument();
        sharedCellTextValue.Values[0][0] = "1";
        AssertEquivalent("shared cell type change", sharedCellTypeChanged, sharedCellTextValue, false);

        SharedSheetDocument rowAdded = CreateDocument();
        rowAdded.RowIds = new object[] { "row-1", "row-2", "row-3" };
        rowAdded.Values = new[]
        {
            new object[] { "A", "ignored-A" },
            new object[] { "B", "ignored-B" },
            new object[] { "C", "ignored-C" }
        };
        AssertEquivalent("row addition", rowAdded, CreateDocument(), false);

        SharedSheetDocument rowDeleted = CreateDocument();
        rowDeleted.RowIds = new object[] { "row-1" };
        rowDeleted.Values = new[] { new object[] { "A", "ignored-A" } };
        AssertEquivalent("row deletion", rowDeleted, CreateDocument(), false);

        SharedSheetDocument rowsReordered = CreateDocument();
        rowsReordered.RowIds = new object[] { "row-2", "row-1" };
        rowsReordered.Values = new[]
        {
            new object[] { "B", "ignored-B" },
            new object[] { "A", "ignored-A" }
        };
        AssertEquivalent("row order change", rowsReordered, CreateDocument(), false);

        SharedSheetDocument sheetRenamed = CreateDocument();
        sheetRenamed.SheetName = "Renamed";
        AssertEquivalent("sheet name change", sheetRenamed, CreateDocument(), false);

        SharedSheetDocument rangeChanged = CreateDocument();
        rangeChanged.RangeAddress = "$A$1:$C$2";
        AssertEquivalent("range change", rangeChanged, CreateDocument(), false);

        SharedSheetDocument settingsChanged = CreateDocument();
        settingsChanged.RangeInfo.IgnoreColumnOffsets = new HashSet<int> { 0, 1 };
        AssertEquivalent("range settings change", settingsChanged, CreateDocument(), false);

        SharedSheetDocumentOperationsTests.Run();
        SharedManifestSafetyTests.Run();

        Console.WriteLine("shared sheet candidate comparer tests passed");
        return 0;
    }

    private static SharedSheetDocument CreateDocument()
    {
        return new SharedSheetDocument
        {
            Project = "project",
            SheetId = "sheet-id",
            SheetName = "Sheet1",
            RangeAddress = "$A$1:$B$2",
            RangeInfo = new SharedRangeInfo
            {
                IdColumnOffset = 0,
                IgnoreColumnOffsets = new HashSet<int> { 1 }
            },
            RowIds = new object[] { "row-1", "row-2" },
            Values = new[]
            {
                new object[] { "A", "ignored-A" },
                new object[] { "B", "ignored-B" }
            }
        };
    }

    private static void AssertEquivalent(
        string scenario,
        SharedSheetDocument left,
        SharedSheetDocument right,
        bool expected)
    {
        bool actual = SharedSheetCandidateComparer.AreEquivalentIgnoringExcludedColumnValues(
            left,
            right,
            NormalizeValue);
        if (actual != expected)
        {
            throw new InvalidOperationException(
                scenario + ": expected " + expected + ", actual " + actual);
        }
    }

    private static object NormalizeValue(object value)
    {
        return value;
    }
}
