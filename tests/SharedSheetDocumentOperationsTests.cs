using System;
using System.Collections.Generic;
using System.Linq;

internal static class SharedSheetDocumentOperationsTests
{
    public static void Run()
    {
        TestDiffEntries();
        TestRowPresenceMergeCases();
        Console.WriteLine("shared sheet diff and merge tests passed");
    }

    private static void TestDiffEntries()
    {
        SharedSheetDocument baseDocument = CreateDocument(
            ("A", "value-A", "ignored-A"),
            ("B", "value-B", "ignored-B"),
            ("C", "value-C", "ignored-C"));

        AssertDiffCount("identical documents", baseDocument, CloneDocument(baseDocument), 0);

        SharedSheetDocument cellChanged = CloneDocument(baseDocument);
        cellChanged.Values[0][1] = "changed-A";
        List<SharedSheetDiffEntry> cellEntries = BuildDiff(baseDocument, cellChanged);
        Assert(cellEntries.Count == 1 && !cellEntries[0].IsRowDeletion,
            "normal cell change must remain a cell diff");

        SharedSheetDocument ignoredChanged = CloneDocument(baseDocument);
        ignoredChanged.Values[0][2] = "local-only";
        AssertDiffCount("ignored column change", baseDocument, ignoredChanged, 0);

        SharedSheetDocument oneDeleted = CreateDocument(
            ("B", "value-B", "ignored-B"),
            ("C", "value-C", "ignored-C"));
        List<SharedSheetDiffEntry> oneDeletedEntries = BuildDiff(baseDocument, oneDeleted);
        Assert(oneDeletedEntries.Count == 1,
            "one deleted row must produce exactly one diff entry");
        AssertRowDeletion(oneDeletedEntries[0], "A");

        SharedSheetDocument twoDeleted = CreateDocument(
            ("C", "value-C", "ignored-C"));
        List<SharedSheetDiffEntry> twoDeletedEntries = BuildDiff(baseDocument, twoDeleted);
        Assert(twoDeletedEntries.Count == 2,
            "two deleted rows must produce exactly two diff entries");
        Assert(twoDeletedEntries.All(entry => entry.IsRowDeletion),
            "deleted rows must be represented as row-level entries");
        Assert(twoDeletedEntries.Select(entry => entry.RowId).SequenceEqual(new[] { "A", "B" }),
            "deleted row IDs must retain base order");

        List<SharedSheetDiffEntry> finalRowDeletedEntries = BuildDiff(
            CreateDocument(("A", "value-A", "ignored-A")),
            CreateDocument());
        Assert(finalRowDeletedEntries.Count == 1,
            "deleting the final row must still produce one row deletion entry");
        AssertRowDeletion(finalRowDeletedEntries[0], "A");

        SharedSheetDocument rowPresenceBase = CreateDocument(
            ("A", "base-A", "ignored-A"),
            ("B", "base-B", "ignored-B"));
        SharedSheetDocument localDeletedA = CreateDocument(
            ("B", "base-B", "ignored-B"));
        SharedSheetDocument remoteEditedA = CreateDocument(
            ("A", "remote-edit-A", "ignored-A"),
            ("B", "base-B", "ignored-B"));
        List<SharedSheetDiffEntry> localDeleteRemoteEditEntries = BuildDiff(
            rowPresenceBase,
            localDeletedA,
            remoteEditedA);
        AssertRowPresenceDiff(
            localDeleteRemoteEditEntries,
            "競合（行削除 vs 編集）",
            "削除",
            "存在（編集あり）");

        SharedSheetDocument localEditedA = CreateDocument(
            ("A", "local-edit-A", "ignored-A"),
            ("B", "base-B", "ignored-B"));
        SharedSheetDocument remoteDeletedA = CreateDocument(
            ("B", "base-B", "ignored-B"));
        List<SharedSheetDiffEntry> localEditRemoteDeleteEntries = BuildDiff(
            rowPresenceBase,
            localEditedA,
            remoteDeletedA);
        AssertRowPresenceDiff(
            localEditRemoteDeleteEntries,
            "競合（編集 vs 行削除）",
            "存在（編集あり）",
            "削除");

        List<SharedSheetDiffEntry> remoteDeleteEntries = BuildDiff(
            rowPresenceBase,
            CloneDocument(rowPresenceBase),
            remoteDeletedA);
        AssertRowPresenceDiff(
            remoteDeleteEntries,
            "共有先で行削除",
            "存在",
            "削除");

        SharedSheetDocument rowAdded = CreateDocument(
            ("A", "value-A", "ignored-A"),
            ("B", "value-B", "ignored-B"),
            ("C", "value-C", "ignored-C"),
            ("D", "value-D", "ignored-D"));
        List<SharedSheetDiffEntry> addedEntries = BuildDiff(baseDocument, rowAdded);
        Assert(addedEntries.Count > 0 &&
            addedEntries.Any(entry => entry.RowId == "D" && !entry.IsRowDeletion),
            "new rows must retain the existing cell-diff behavior");
    }

    private static void TestRowPresenceMergeCases()
    {
        SharedSheetDocument baseDocument = CreateDocument(
            ("A", "base-A", "ignored-A"),
            ("B", "base-B", "ignored-B"));

        AssertMerge(
            "local-only deletion",
            baseDocument,
            CreateDocument(("B", "base-B", "ignored-B")),
            CreateDocument(("A", "base-A", "ignored-A"), ("B", "remote-B", "ignored-B")),
            expectedAExists: false,
            expectedConflicts: 0);

        AssertMerge(
            "remote-only deletion",
            baseDocument,
            CreateDocument(("A", "base-A", "ignored-A"), ("B", "base-B", "ignored-B")),
            CreateDocument(("B", "base-B", "ignored-B")),
            expectedAExists: false,
            expectedConflicts: 0);

        AssertMerge(
            "both delete",
            baseDocument,
            CreateDocument(("B", "base-B", "ignored-B")),
            CreateDocument(("B", "base-B", "ignored-B")),
            expectedAExists: false,
            expectedConflicts: 0);

        SharedSheetUploadMergeResult allRowsDeleted = Merge(
            CreateDocument(("A", "base-A", "ignored-A")),
            CreateDocument(),
            CreateDocument());
        Assert(allRowsDeleted.ConflictCount == 0 &&
            allRowsDeleted.MergedDocument.RowIds.Length == 0,
            "deleting the final row on both sides must produce an empty merge without conflict");

        SharedSheetUploadMergeResult finalRemoteDeletion = Merge(
            CreateDocument(("A", "base-A", "ignored-A")),
            CreateDocument(("A", "base-A", "ignored-A")),
            CreateDocument());
        Assert(finalRemoteDeletion.ConflictCount == 0 &&
            finalRemoteDeletion.MergedDocument.RowIds.Length == 0 &&
            finalRemoteDeletion.MergedDocument.Values.Length == 0,
            "remote deletion of the final unchanged row must produce a structurally valid empty merge");

        SharedSheetUploadMergeResult finalLocalDeletion = Merge(
            CreateDocument(("A", "base-A", "ignored-A")),
            CreateDocument(),
            CreateDocument(("A", "base-A", "ignored-A")));
        Assert(finalLocalDeletion.ConflictCount == 0 &&
            finalLocalDeletion.MergedDocument.RowIds.Length == 0 &&
            finalLocalDeletion.MergedDocument.Values.Length == 0,
            "local deletion of the final unchanged row must produce a structurally valid empty merge");

        SharedSheetUploadMergeResult localDeleteRemoteEdit = AssertMerge(
            "local deletion versus remote edit",
            baseDocument,
            CreateDocument(("B", "base-B", "ignored-B")),
            CreateDocument(("A", "remote-edit-A", "ignored-A"), ("B", "base-B", "ignored-B")),
            expectedAExists: true,
            expectedConflicts: 1);
        Assert(GetSharedValue(localDeleteRemoteEdit.MergedDocument, "A") == "remote-edit-A",
            "delete/edit conflict must retain the edited remote row in the blocked merge result");

        SharedSheetUploadMergeResult localEditRemoteDelete = AssertMerge(
            "local edit versus remote deletion",
            baseDocument,
            CreateDocument(("A", "local-edit-A", "ignored-A"), ("B", "base-B", "ignored-B")),
            CreateDocument(("B", "base-B", "ignored-B")),
            expectedAExists: true,
            expectedConflicts: 1);
        Assert(GetSharedValue(localEditRemoteDelete.MergedDocument, "A") == "local-edit-A",
            "edit/delete conflict must retain the edited local row in the blocked merge result");

        SharedSheetDocument noABase = CreateDocument(("B", "base-B", "ignored-B"));
        AssertMerge(
            "remote row addition",
            noABase,
            CreateDocument(("B", "base-B", "ignored-B")),
            CreateDocument(("B", "base-B", "ignored-B"), ("A", "remote-new-A", "ignored-A")),
            expectedAExists: true,
            expectedConflicts: 0);

        AssertMerge(
            "local row addition",
            noABase,
            CreateDocument(("B", "base-B", "ignored-B"), ("A", "local-new-A", "ignored-A")),
            CreateDocument(("B", "base-B", "ignored-B")),
            expectedAExists: true,
            expectedConflicts: 0);

        TestExistingCellMergeBehavior();
    }

    private static void TestExistingCellMergeBehavior()
    {
        SharedSheetDocument baseDocument = CreateDocument(("A", "base", "ignored-base"));

        SharedSheetUploadMergeResult localChanged = Merge(
            baseDocument,
            CreateDocument(("A", "local", "ignored-local")),
            CreateDocument(("A", "base", "ignored-remote")));
        Assert(localChanged.ConflictCount == 0 &&
            GetSharedValue(localChanged.MergedDocument, "A") == "local",
            "existing local-only cell changes must retain the existing merge behavior");
        Assert(localChanged.MergedDocument.Values[0][2].ToString() == "ignored-local",
            "ignored column values must continue to come from local");

        SharedSheetUploadMergeResult remoteChanged = Merge(
            baseDocument,
            CreateDocument(("A", "base", "ignored-local")),
            CreateDocument(("A", "remote", "ignored-remote")));
        Assert(remoteChanged.ConflictCount == 0 &&
            GetSharedValue(remoteChanged.MergedDocument, "A") == "remote",
            "existing remote-only cell changes must retain the existing merge behavior");

        SharedSheetUploadMergeResult bothChanged = Merge(
            baseDocument,
            CreateDocument(("A", "local", "ignored-local")),
            CreateDocument(("A", "remote", "ignored-remote")));
        Assert(bothChanged.ConflictCount == 1,
            "different edits to an existing cell must remain a conflict");
    }

    private static SharedSheetUploadMergeResult AssertMerge(
        string scenario,
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument,
        bool expectedAExists,
        int expectedConflicts)
    {
        SharedSheetUploadMergeResult result = Merge(baseDocument, localDocument, remoteDocument);
        Assert(result != null && result.MergedDocument != null,
            scenario + ": merge result must exist");
        Assert(ContainsRow(result.MergedDocument, "A") == expectedAExists,
            scenario + ": unexpected row A presence");
        Assert(result.ConflictCount == expectedConflicts,
            scenario + ": expected conflicts=" + expectedConflicts +
            ", actual=" + result.ConflictCount);
        return result;
    }

    private static SharedSheetUploadMergeResult Merge(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument)
    {
        return SharedSheetUploadMergeEngine.Merge(
            baseDocument,
            localDocument,
            remoteDocument,
            ValuesEqual,
            NormalizeValue);
    }

    private static List<SharedSheetDiffEntry> BuildDiff(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument)
    {
        return BuildDiff(baseDocument, localDocument, baseDocument);
    }

    private static List<SharedSheetDiffEntry> BuildDiff(
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        SharedSheetDocument remoteDocument)
    {
        return SharedSheetDiffBuilder.BuildEntries(
            baseDocument,
            localDocument,
            remoteDocument,
            localDocument,
            ValuesEqual,
            NormalizeValue);
    }

    private static void AssertDiffCount(
        string scenario,
        SharedSheetDocument baseDocument,
        SharedSheetDocument localDocument,
        int expectedCount)
    {
        int actualCount = BuildDiff(baseDocument, localDocument).Count;
        Assert(actualCount == expectedCount,
            scenario + ": expected diff count=" + expectedCount +
            ", actual=" + actualCount);
    }

    private static void AssertRowDeletion(SharedSheetDiffEntry entry, string expectedRowId)
    {
        Assert(entry.IsRowDeletion && entry.IsRowLevelChange,
            "deleted row entry must be row-level");
        Assert(entry.RowId == expectedRowId, "deleted row entry must expose its row ID");
        Assert(entry.StateLabel == "行削除", "deleted row entry must use the row deletion state");
        Assert(entry.BaseText == "存在" && entry.LocalText == "削除",
            "deleted row entry must expose Base and Local row presence");
        Assert(string.IsNullOrWhiteSpace(entry.CellAddress),
            "deleted row entry must not contain a cell address");
        Assert(entry.CellAddressText == "-", "deleted row entry must display a dash for its address");
    }

    private static void AssertRowPresenceDiff(
        List<SharedSheetDiffEntry> entries,
        string expectedState,
        string expectedLocalText,
        string expectedRemoteText)
    {
        Assert(entries.Count == 1,
            expectedState + ": row presence change must produce exactly one entry");
        SharedSheetDiffEntry entry = entries[0];
        Assert(entry.IsRowLevelChange, expectedState + ": entry must be row-level");
        Assert(entry.RowId == "A", expectedState + ": entry must expose row ID A");
        Assert(entry.StateLabel == expectedState,
            expectedState + ": unexpected state label " + entry.StateLabel);
        Assert(entry.BaseText == "存在", expectedState + ": Base row state must be visible");
        Assert(entry.LocalText == expectedLocalText,
            expectedState + ": unexpected Local row state " + entry.LocalText);
        Assert(entry.HasRemoteValue && entry.RemoteText == expectedRemoteText,
            expectedState + ": unexpected Remote row state " + entry.RemoteText);
        Assert(string.IsNullOrWhiteSpace(entry.CellAddress),
            expectedState + ": row-level entry must not contain a cell address");
    }

    private static SharedSheetDocument CreateDocument(
        params (string Id, string Value, string IgnoredValue)[] rows)
    {
        return new SharedSheetDocument
        {
            Project = "project",
            SheetId = "sheet-id",
            SheetName = "Sheet1",
            RangeAddress = "$A$2:$C$20",
            RangeInfo = new SharedRangeInfo
            {
                IdColumnOffset = 0,
                IgnoreColumnOffsets = new HashSet<int> { 2 }
            },
            RowIds = rows.Select(row => (object)row.Id).ToArray(),
            Values = rows.Select(row => new object[]
            {
                row.Id,
                row.Value,
                row.IgnoredValue
            }).ToArray()
        };
    }

    private static SharedSheetDocument CloneDocument(SharedSheetDocument source)
    {
        return new SharedSheetDocument
        {
            Project = source.Project,
            SheetId = source.SheetId,
            SheetName = source.SheetName,
            RangeAddress = source.RangeAddress,
            RangeInfo = new SharedRangeInfo
            {
                IdColumnOffset = source.RangeInfo.IdColumnOffset,
                IgnoreColumnOffsets = new HashSet<int>(source.RangeInfo.IgnoreColumnOffsets)
            },
            RowIds = source.RowIds.ToArray(),
            Values = source.Values.Select(row => row.ToArray()).ToArray()
        };
    }

    private static bool ContainsRow(SharedSheetDocument document, string rowId)
    {
        return document.RowIds.Any(value => string.Equals(value == null ? null : value.ToString(), rowId, StringComparison.Ordinal));
    }

    private static string GetSharedValue(SharedSheetDocument document, string rowId)
    {
        for (int index = 0; index < document.RowIds.Length; index++)
        {
            if (string.Equals(document.RowIds[index].ToString(), rowId, StringComparison.Ordinal))
            {
                return document.Values[index][1] == null ? null : document.Values[index][1].ToString();
            }
        }

        return null;
    }

    private static bool ValuesEqual(object left, object right)
    {
        return left == null || right == null
            ? left == null && right == null
            : left.Equals(right);
    }

    private static object NormalizeValue(object value)
    {
        return value;
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}
