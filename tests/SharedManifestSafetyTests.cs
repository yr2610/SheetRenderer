using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

internal static class SharedManifestSafetyTests
{
    public static void Run()
    {
        MetadataSuccessReturnsFoundWithoutFallbacks().GetAwaiter().GetResult();
        MetadataNotFoundRawFoundWithPathCommitReturnsFoundNotCreate().GetAwaiter().GetResult();
        MetadataAndRawNotFoundTreeBlobFoundReturnsFoundNotCreate().GetAwaiter().GetResult();
        MetadataMissingContentOrEncodingFallsBack().GetAwaiter().GetResult();
        ContentFoundWithoutFileLastCommitIsIndeterminateAndNotCreate().GetAwaiter().GetResult();
        AllRoutesConfirmedMissingAllowsCreateOnlyAfterValidatedProjectAndRef().GetAwaiter().GetResult();
        InvalidBase64FallsBackAndNeverBecomesNotFound().GetAwaiter().GetResult();
        InvalidManifestJsonIsIndeterminateAndNotCreate().GetAwaiter().GetResult();
        AuthenticationServerAndNetworkFailuresAreIndeterminate().GetAwaiter().GetResult();
        TreeBlobReadFailureIsIndeterminateAndNotCreate().GetAwaiter().GetResult();
        Tree404ClassificationRequiresTypedValidation();
        UpdateLastCommitMismatchIsClassifiedAsConcurrentUpdate();
        InitialCreateExistingFileIsClassifiedAsConcurrentCreate();
        UnrelatedHttp400IsNotClassifiedAsConflict();
        ExistingManifestCommitPlanUsesUpdateAndExactLastCommitId();
        ConfirmedMissingManifestCommitPlanUsesCreateWithoutLastCommitId();
        MetadataInvalidContentUsesMetadataCommitAsPinnedRef().GetAwaiter().GetResult();
        MetadataNotFoundPinsCommitBeforeContent().GetAwaiter().GetResult();
        BranchAdvanceAfterPinKeepsContentAndLastCommitAtPinnedCommit().GetAwaiter().GetResult();
        MismatchedContentAndCommitSnapshotIsRejected().GetAwaiter().GetResult();
        PinnedCommitWithoutPinnedContentIsIndeterminate().GetAwaiter().GetResult();
        EmptyCommitHistoryWithBranchContentIsIndeterminate().GetAwaiter().GetResult();
        TreePagingFindsManifestOnSecondPage().GetAwaiter().GetResult();
        TreePagingChecksEmptySecondPageBeforeNotFound().GetAwaiter().GetResult();
        UnsupportedEndpointClassificationAndFallbacks().GetAwaiter().GetResult();
        UnsupportedRoutesDoNotCountAsNotFound().GetAwaiter().GetResult();
        UnsupportedMetadataWithReliableNotFoundRoutesIsConfirmedNotFound().GetAwaiter().GetResult();
        FatalEndpointFailuresStopImmediately().GetAwaiter().GetResult();
        Console.WriteLine("shared manifest safety tests passed");
    }

    private static async Task MetadataSuccessReturnsFoundWithoutFallbacks()
    {
        int rawCalls = 0;
        int treeCalls = 0;
        int commitCalls = 0;
        SharedManifestProbeResult result = await Probe(
            MetadataFound("valid", "file-commit"),
            () => { rawCalls++; return ContentNotFound(); },
            () => { treeCalls++; return ContentNotFound(); },
            () => { commitCalls++; return Task.FromResult("unexpected"); });

        Assert(result.State == SharedManifestProbeState.Found, nameof(MetadataSuccessReturnsFoundWithoutFallbacks));
        Assert(result.LastCommitId == "file-commit", nameof(MetadataSuccessReturnsFoundWithoutFallbacks));
        Assert(rawCalls == 0 && treeCalls == 0 && commitCalls == 0,
            nameof(MetadataSuccessReturnsFoundWithoutFallbacks) + ": fallback must not run");
    }

    private static async Task MetadataNotFoundRawFoundWithPathCommitReturnsFoundNotCreate()
    {
        SharedManifestProbeResult result = await Probe(
            MetadataNotFound(),
            () => ContentFound("valid"),
            ContentNotFound,
            () => Task.FromResult("raw-file-commit"));

        Assert(result.State == SharedManifestProbeState.Found && result.ContentRoute == "raw",
            nameof(MetadataNotFoundRawFoundWithPathCommitReturnsFoundNotCreate));
        Assert(SharedManifestSafety.CreateCommitPlan(result).Action == "update",
            nameof(MetadataNotFoundRawFoundWithPathCommitReturnsFoundNotCreate) + ": existing file must not create");
    }

    private static async Task MetadataAndRawNotFoundTreeBlobFoundReturnsFoundNotCreate()
    {
        SharedManifestProbeResult result = await Probe(
            MetadataNotFound(),
            ContentNotFound,
            () => ContentFound("valid"),
            () => Task.FromResult("tree-file-commit"));

        Assert(result.State == SharedManifestProbeState.Found && result.ContentRoute == "tree-blob",
            nameof(MetadataAndRawNotFoundTreeBlobFoundReturnsFoundNotCreate));
        Assert(SharedManifestSafety.CreateCommitPlan(result).Action == "update",
            nameof(MetadataAndRawNotFoundTreeBlobFoundReturnsFoundNotCreate) + ": existing file must not create");
    }

    private static async Task MetadataMissingContentOrEncodingFallsBack()
    {
        int rawCalls = 0;
        SharedManifestMetadataProbe metadata = MetadataFound("valid", "metadata-commit");
        metadata.Encoding = null;
        SharedManifestProbeResult result = await Probe(
            metadata,
            () => { rawCalls++; return ContentFound("valid"); },
            ContentNotFound,
            () => Task.FromResult("unexpected"));

        Assert(result.State == SharedManifestProbeState.Found && rawCalls == 1,
            nameof(MetadataMissingContentOrEncodingFallsBack));
        Assert(result.LastCommitId == "metadata-commit",
            nameof(MetadataMissingContentOrEncodingFallsBack) + ": metadata last_commit_id remains usable");
    }

    private static async Task ContentFoundWithoutFileLastCommitIsIndeterminateAndNotCreate()
    {
        SharedManifestProbeResult result = await Probe(
            MetadataNotFound(),
            () => ContentFound("valid"),
            ContentNotFound,
            () => Task.FromResult<string>(null));

        AssertIndeterminateAndCreateRejected(result,
            nameof(ContentFoundWithoutFileLastCommitIsIndeterminateAndNotCreate));
    }

    private static async Task AllRoutesConfirmedMissingAllowsCreateOnlyAfterValidatedProjectAndRef()
    {
        SharedManifestProbeResult validated = await Probe(
            MetadataNotFound(), ContentNotFound, ContentNotFound,
            () => Task.FromResult<string>(null), true);
        Assert(validated.State == SharedManifestProbeState.ConfirmedNotFound,
            nameof(AllRoutesConfirmedMissingAllowsCreateOnlyAfterValidatedProjectAndRef));
        Assert(SharedManifestSafety.CreateCommitPlan(validated).Action == "create",
            nameof(AllRoutesConfirmedMissingAllowsCreateOnlyAfterValidatedProjectAndRef));

        SharedManifestProbeResult unvalidated = await Probe(
            MetadataNotFound(), ContentNotFound, ContentNotFound,
            () => Task.FromResult<string>(null), false);
        AssertIndeterminateAndCreateRejected(unvalidated,
            nameof(AllRoutesConfirmedMissingAllowsCreateOnlyAfterValidatedProjectAndRef) + ": unvalidated ref");
    }

    private static async Task InvalidBase64FallsBackAndNeverBecomesNotFound()
    {
        SharedManifestMetadataProbe invalid = MetadataFound("valid", "file-commit");
        invalid.Content = "%%%not-base64%%%";
        SharedManifestProbeResult recovered = await Probe(
            invalid,
            () => ContentFound("valid"),
            ContentNotFound,
            () => Task.FromResult("unexpected"));
        Assert(recovered.State == SharedManifestProbeState.Found && recovered.ContentRoute == "raw",
            nameof(InvalidBase64FallsBackAndNeverBecomesNotFound) + ": fallback should recover");

        SharedManifestProbeResult failed = await Probe(
            invalid,
            ContentNotFound,
            ContentNotFound,
            () => Task.FromResult("unexpected"));
        AssertIndeterminateAndCreateRejected(failed,
            nameof(InvalidBase64FallsBackAndNeverBecomesNotFound) + ": malformed content is not missing");
    }

    private static async Task InvalidManifestJsonIsIndeterminateAndNotCreate()
    {
        SharedManifestProbeResult result = await Probe(
            MetadataFound("invalid-json", "file-commit"),
            ContentNotFound,
            ContentNotFound,
            () => Task.FromResult("unexpected"));
        AssertIndeterminateAndCreateRejected(result, nameof(InvalidManifestJsonIsIndeterminateAndNotCreate));
    }

    private static async Task AuthenticationServerAndNetworkFailuresAreIndeterminate()
    {
        Exception[] errors =
        {
            ApiError(401, "unauthorized"),
            ApiError(403, "forbidden"),
            ApiError(500, "server error"),
            new IOException("network failed")
        };

        foreach (Exception error in errors)
        {
            SharedManifestProbeResult result = await Probe(
                new SharedManifestMetadataProbe
                {
                    State = SharedManifestEndpointState.Failed,
                    Error = error
                },
                ContentNotFound,
                ContentNotFound,
                () => Task.FromResult<string>(null));
            AssertIndeterminateAndCreateRejected(result,
                nameof(AuthenticationServerAndNetworkFailuresAreIndeterminate) + ": " + error.GetType().Name);
        }
    }

    private static async Task TreeBlobReadFailureIsIndeterminateAndNotCreate()
    {
        GitLabApiException blobFailure = new GitLabApiException(
            "GitLab resource not found.",
            404,
            "https://gitlab.example/api/v4/projects/1/repository/blobs/blob-id/raw",
            "not found");
        var pagedBlobFailure = new GitLabTreeBlobDownloadException(
            blobFailure.Message,
            2,
            2,
            blobFailure);
        SharedManifestContentProbe classified =
            SharedManifestSafety.ClassifyTreeBlobException(pagedBlobFailure);
        Assert(classified.State == SharedManifestEndpointState.Failed &&
            classified.PagesChecked == 2 &&
            classified.FoundPage == 2,
            nameof(TreeBlobReadFailureIsIndeterminateAndNotCreate) + ": blob 404 must remain a failure");

        SharedManifestProbeResult result = await Probe(
            MetadataNotFound(),
            ContentNotFound,
            () => Task.FromResult(classified),
            () => Task.FromResult<string>(null));
        AssertIndeterminateAndCreateRejected(result, nameof(TreeBlobReadFailureIsIndeterminateAndNotCreate));

        SharedManifestContentProbe absentFromTree = SharedManifestSafety.ClassifyTreeBlobException(
            new GitLabTreeFileNotFoundException(
                "File not found in tree. folder=project file=_manifest.json ref=main",
                2));
        Assert(absentFromTree.State == SharedManifestEndpointState.NotFound,
            nameof(TreeBlobReadFailureIsIndeterminateAndNotCreate) + ": absent tree entry is a confirmed route miss");
    }

    private static void Tree404ClassificationRequiresTypedValidation()
    {
        var rawTree404 = new GitLabApiException(
            "tree path missing",
            404,
            "https://gitlab.example/api/v4/projects/1/repository/tree?path=project&ref=main",
            "not found");
        SharedManifestContentProbe rawClassification =
            SharedManifestSafety.ClassifyTreeBlobException(rawTree404);
        Assert(rawClassification.State == SharedManifestEndpointState.Failed,
            nameof(Tree404ClassificationRequiresTypedValidation) + ": raw tree 404");

        var validatedPathMiss = new GitLabTreeFileNotFoundException(
            "validated tree path missing",
            1,
            GitLabTreeNotFoundReason.PathNotFoundAfterValidated404,
            rawTree404);
        SharedManifestContentProbe validatedClassification =
            SharedManifestSafety.ClassifyTreeBlobException(validatedPathMiss);
        Assert(validatedClassification.State == SharedManifestEndpointState.NotFound &&
            validatedClassification.TreeNotFoundReason ==
                GitLabTreeNotFoundReason.PathNotFoundAfterValidated404.ToString(),
            nameof(Tree404ClassificationRequiresTypedValidation) + ": validated path miss");

        var pagedFileMiss = new GitLabTreeFileNotFoundException(
            "file missing after paging",
            2);
        SharedManifestContentProbe pagedClassification =
            SharedManifestSafety.ClassifyTreeBlobException(pagedFileMiss);
        Assert(pagedClassification.State == SharedManifestEndpointState.NotFound &&
            pagedClassification.TreeNotFoundReason ==
                GitLabTreeNotFoundReason.FileNotFoundAfterPaging.ToString(),
            nameof(Tree404ClassificationRequiresTypedValidation) + ": paged file miss");
    }

    private static async Task MetadataInvalidContentUsesMetadataCommitAsPinnedRef()
    {
        SharedManifestMetadataProbe metadata = MetadataFound("valid", "commit-C");
        metadata.Content = "invalid-base64";
        var refs = new List<string>();
        int commitCalls = 0;

        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => Task.FromResult(metadata),
            () =>
            {
                commitCalls++;
                return Task.FromResult(CommitFound("unexpected"));
            },
            contentRef =>
            {
                refs.Add(contentRef);
                return Task.FromResult(PinnedContentFound("valid", contentRef, "raw"));
            },
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")),
            ParseManifest,
            null);

        Assert(result.State == SharedManifestProbeState.Found &&
            result.LastCommitId == "commit-C" &&
            result.ContentRef == "commit-C",
            nameof(MetadataInvalidContentUsesMetadataCommitAsPinnedRef));
        Assert(refs.Count == 1 && refs[0] == "commit-C" && commitCalls == 0,
            nameof(MetadataInvalidContentUsesMetadataCommitAsPinnedRef) +
            ": fallback must use metadata commit C, never the branch ref");
    }

    private static async Task MetadataNotFoundPinsCommitBeforeContent()
    {
        var sequence = new List<string>();
        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => Task.FromResult(MetadataNotFound()),
            () =>
            {
                sequence.Add("commit:commit-C");
                return Task.FromResult(CommitFound("commit-C"));
            },
            contentRef =>
            {
                sequence.Add("raw:" + contentRef);
                return Task.FromResult(PinnedContentFound("valid", contentRef, "raw"));
            },
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")),
            ParseManifest,
            null);

        Assert(result.State == SharedManifestProbeState.Found && result.LastCommitId == "commit-C",
            nameof(MetadataNotFoundPinsCommitBeforeContent));
        Assert(sequence.Count == 2 &&
            sequence[0] == "commit:commit-C" &&
            sequence[1] == "raw:commit-C",
            nameof(MetadataNotFoundPinsCommitBeforeContent) +
            ": commit C must be resolved before content is read at ref C");
    }

    private static async Task BranchAdvanceAfterPinKeepsContentAndLastCommitAtPinnedCommit()
    {
        string branchTip = "commit-C";
        string contentRefUsed = null;
        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => Task.FromResult(MetadataNotFound()),
            () =>
            {
                string pinned = branchTip;
                branchTip = "commit-D";
                return Task.FromResult(CommitFound(pinned));
            },
            contentRef =>
            {
                contentRefUsed = contentRef;
                return Task.FromResult(PinnedContentFound("valid", contentRef, "raw"));
            },
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")),
            ParseManifest,
            null);

        SharedManifestCommitPlan plan = SharedManifestSafety.CreateCommitPlan(result);
        Assert(branchTip == "commit-D" && contentRefUsed == "commit-C",
            nameof(BranchAdvanceAfterPinKeepsContentAndLastCommitAtPinnedCommit));
        Assert(result.LastCommitId == "commit-C" && plan.LastCommitId == "commit-C",
            nameof(BranchAdvanceAfterPinKeepsContentAndLastCommitAtPinnedCommit) +
            ": later update against D can be rejected by last_commit_id C");
    }

    private static async Task MismatchedContentAndCommitSnapshotIsRejected()
    {
        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => Task.FromResult(MetadataNotFound()),
            () => Task.FromResult(CommitFound("commit-C")),
            contentRef => Task.FromResult(new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.Found,
                Content = Encoding.UTF8.GetBytes("valid"),
                ContentRef = contentRef,
                LastCommitId = "commit-B",
                ContentRoute = "raw"
            }),
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")),
            ParseManifest,
            null);

        AssertIndeterminateAndCreateRejected(result, nameof(MismatchedContentAndCommitSnapshotIsRejected));
        Assert(result.DecisionReason == "pinned-snapshot-mismatch",
            nameof(MismatchedContentAndCommitSnapshotIsRejected) +
            ": old content A plus newer commit B must be structurally rejected");
    }

    private static async Task PinnedCommitWithoutPinnedContentIsIndeterminate()
    {
        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => Task.FromResult(MetadataNotFound()),
            () => Task.FromResult(CommitFound("commit-C")),
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "raw")),
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")),
            ParseManifest,
            null);

        AssertIndeterminateAndCreateRejected(result, nameof(PinnedCommitWithoutPinnedContentIsIndeterminate));
        Assert(result.DecisionReason == "pinned-content-unavailable",
            nameof(PinnedCommitWithoutPinnedContentIsIndeterminate));
    }

    private static async Task EmptyCommitHistoryWithBranchContentIsIndeterminate()
    {
        string contentRefUsed = null;
        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => Task.FromResult(MetadataNotFound()),
            () => Task.FromResult(CommitNotFound()),
            contentRef =>
            {
                contentRefUsed = contentRef;
                return Task.FromResult(PinnedContentFound("valid", contentRef, "raw"));
            },
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")),
            ParseManifest,
            null);

        Assert(contentRefUsed == "main", nameof(EmptyCommitHistoryWithBranchContentIsIndeterminate));
        AssertIndeterminateAndCreateRejected(result, nameof(EmptyCommitHistoryWithBranchContentIsIndeterminate));
        Assert(result.State != SharedManifestProbeState.ConfirmedNotFound,
            nameof(EmptyCommitHistoryWithBranchContentIsIndeterminate));
    }

    private static async Task TreePagingFindsManifestOnSecondPage()
    {
        var requestedPages = new List<int>();
        GitLabTreeSearchResult result = await GitLabTreePaging.FindBlobAsync(
            "_manifest.json",
            100,
            page =>
            {
                requestedPages.Add(page);
                if (page == 1)
                {
                    return Task.FromResult(CreateTreeItems(100, "other-"));
                }

                return Task.FromResult(new List<GitLabTreeItem>
                {
                    new GitLabTreeItem { Id = "manifest-blob", Name = "_manifest.json", Type = "blob" }
                });
            });

        Assert(result.Target != null && result.Target.Id == "manifest-blob",
            nameof(TreePagingFindsManifestOnSecondPage));
        Assert(result.PagesChecked == 2 && result.FoundPage == 2 &&
            requestedPages.Count == 2 && requestedPages[0] == 1 && requestedPages[1] == 2,
            nameof(TreePagingFindsManifestOnSecondPage) + ": page one alone must not produce NotFound");
    }

    private static async Task TreePagingChecksEmptySecondPageBeforeNotFound()
    {
        var requestedPages = new List<int>();
        GitLabTreeSearchResult result = await GitLabTreePaging.FindBlobAsync(
            "_manifest.json",
            100,
            page =>
            {
                requestedPages.Add(page);
                return Task.FromResult(page == 1
                    ? CreateTreeItems(100, "other-")
                    : new List<GitLabTreeItem>());
            });

        Assert(result.Target == null && result.PagesChecked == 2,
            nameof(TreePagingChecksEmptySecondPageBeforeNotFound));
        Assert(requestedPages.Count == 2,
            nameof(TreePagingChecksEmptySecondPageBeforeNotFound) +
            ": a full first page requires checking page two");
    }

    private static async Task UnsupportedEndpointClassificationAndFallbacks()
    {
        int[] unsupportedStatuses = { 405, 501 };
        foreach (int status in unsupportedStatuses)
        {
            int rawCalls = 0;
            SharedManifestProbeResult result = await ProbeWithMetadataException(
                ApiError(status, "unsupported"),
                CommitFound("commit-C"),
                contentRef =>
                {
                    rawCalls++;
                    return Task.FromResult(PinnedContentFound("valid", contentRef, "raw"));
                },
                contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")));
            Assert(result.State == SharedManifestProbeState.Found && rawCalls == 1,
                nameof(UnsupportedEndpointClassificationAndFallbacks) + ": status " + status);
        }

        SharedManifestProbeResult explicit400 = await ProbeWithMetadataException(
            ApiError(400, "unknown parameter: ref"),
            CommitFound("commit-C"),
            contentRef => Task.FromResult(PinnedContentFound("valid", contentRef, "raw")),
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")));
        Assert(explicit400.State == SharedManifestProbeState.Found,
            nameof(UnsupportedEndpointClassificationAndFallbacks) + ": explicit unsupported 400");

        int unrelatedRawCalls = 0;
        SharedManifestProbeResult unrelated400 = await ProbeWithMetadataException(
            ApiError(400, "branch is protected"),
            CommitFound("commit-C"),
            contentRef =>
            {
                unrelatedRawCalls++;
                return Task.FromResult(PinnedContentFound("valid", contentRef, "raw"));
            },
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")));
        AssertIndeterminateAndCreateRejected(unrelated400,
            nameof(UnsupportedEndpointClassificationAndFallbacks) + ": unrelated 400");
        Assert(unrelatedRawCalls == 0,
            nameof(UnsupportedEndpointClassificationAndFallbacks) + ": unrelated 400 must stop immediately");

        SharedManifestProbeResult raw405 = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => Task.FromResult(MetadataNotFound()),
            () => Task.FromResult(CommitFound("commit-C")),
            contentRef => throw ApiError(405, "raw unsupported"),
            contentRef => Task.FromResult(PinnedContentFound("valid", contentRef, "tree-blob")),
            ParseManifest,
            null);
        Assert(raw405.State == SharedManifestProbeState.Found && raw405.ContentRoute == "tree-blob",
            nameof(UnsupportedEndpointClassificationAndFallbacks) + ": raw 405 must continue to tree");

        SharedManifestProbeResult metadataAndRawUnsupported = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => throw ApiError(501, "metadata unsupported"),
            () => Task.FromResult(CommitFound("commit-C")),
            contentRef => throw ApiError(405, "raw unsupported"),
            contentRef => Task.FromResult(PinnedContentFound("valid", contentRef, "tree-blob")),
            ParseManifest,
            null);
        Assert(metadataAndRawUnsupported.State == SharedManifestProbeState.Found,
            nameof(UnsupportedEndpointClassificationAndFallbacks) +
            ": metadata and raw unsupported must allow pinned tree success");
    }

    private static async Task UnsupportedRoutesDoNotCountAsNotFound()
    {
        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => throw ApiError(405, "metadata unsupported"),
            () => throw ApiError(501, "commits not implemented"),
            contentRef => throw ApiError(405, "raw unsupported"),
            contentRef => throw ApiError(501, "tree not implemented"),
            ParseManifest,
            null);

        AssertIndeterminateAndCreateRejected(result, nameof(UnsupportedRoutesDoNotCountAsNotFound));
        Assert(result.State != SharedManifestProbeState.ConfirmedNotFound,
            nameof(UnsupportedRoutesDoNotCountAsNotFound));
    }

    private static async Task UnsupportedMetadataWithReliableNotFoundRoutesIsConfirmedNotFound()
    {
        SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => throw ApiError(405, "metadata unsupported"),
            () => Task.FromResult(CommitNotFound()),
            contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "raw")),
            contentRef =>
            {
                SharedManifestContentProbe missing = PinnedContentNotFound(contentRef, "tree-blob");
                missing.PagesChecked = 2;
                return Task.FromResult(missing);
            },
            ParseManifest,
            null);

        Assert(result.State == SharedManifestProbeState.ConfirmedNotFound,
            nameof(UnsupportedMetadataWithReliableNotFoundRoutesIsConfirmedNotFound));
        Assert(SharedManifestSafety.CreateCommitPlan(result).Action == "create",
            nameof(UnsupportedMetadataWithReliableNotFoundRoutesIsConfirmedNotFound));
    }

    private static async Task FatalEndpointFailuresStopImmediately()
    {
        Exception[] errors =
        {
            ApiError(401, "unauthorized"),
            ApiError(403, "forbidden"),
            ApiError(408, "request timeout"),
            ApiError(429, "rate limited"),
            ApiError(500, "server error"),
            new GitLabResponseFormatException(
                200,
                "https://gitlab.example/api/v4/projects/1/repository/files/manifest?ref=main",
                "a non-null repository file metadata object",
                Encoding.UTF8.GetBytes("null")),
            new IOException("network failed")
        };

        foreach (Exception error in errors)
        {
            int commitCalls = 0;
            int rawCalls = 0;
            SharedManifestProbeResult result = await SharedManifestSafety.ProbeAsync(
                true,
                "main",
                () => throw error,
                () =>
                {
                    commitCalls++;
                    return Task.FromResult(CommitFound("unexpected"));
                },
                contentRef =>
                {
                    rawCalls++;
                    return Task.FromResult(PinnedContentFound("valid", contentRef, "raw"));
                },
                contentRef => Task.FromResult(PinnedContentNotFound(contentRef, "tree-blob")),
                ParseManifest,
                null);
            AssertIndeterminateAndCreateRejected(result,
                nameof(FatalEndpointFailuresStopImmediately) + ": " + error.Message);
            Assert(commitCalls == 0 && rawCalls == 0,
                nameof(FatalEndpointFailuresStopImmediately) + ": fatal errors must not fall back");
        }
    }

    private static void UpdateLastCommitMismatchIsClassifiedAsConcurrentUpdate()
    {
        GitLabApiException error = ApiError(
            400,
            "The file has changed since you started editing it: project/_manifest.json");
        Assert(SharedManifestSafety.ClassifyCommitConflict(error, "update") ==
            SharedManifestCommitConflictKind.UpdatedAfterRead,
            nameof(UpdateLastCommitMismatchIsClassifiedAsConcurrentUpdate));
    }

    private static void InitialCreateExistingFileIsClassifiedAsConcurrentCreate()
    {
        GitLabApiException error = ApiError(400, "A file with this name already exists");
        Assert(SharedManifestSafety.ClassifyCommitConflict(error, "create") ==
            SharedManifestCommitConflictKind.CreatedAfterProbe,
            nameof(InitialCreateExistingFileIsClassifiedAsConcurrentCreate));
    }

    private static void UnrelatedHttp400IsNotClassifiedAsConflict()
    {
        GitLabApiException error = ApiError(400, "branch is protected");
        Assert(SharedManifestSafety.ClassifyCommitConflict(error, "update") ==
            SharedManifestCommitConflictKind.None,
            nameof(UnrelatedHttp400IsNotClassifiedAsConflict));
    }

    private static void ExistingManifestCommitPlanUsesUpdateAndExactLastCommitId()
    {
        SharedManifestCommitPlan plan = SharedManifestSafety.CreateCommitPlan(new SharedManifestProbeResult
        {
            State = SharedManifestProbeState.Found,
            Manifest = ValidManifest(),
            LastCommitId = "exact-file-commit"
        });
        Assert(plan.Action == "update" && plan.LastCommitId == "exact-file-commit",
            nameof(ExistingManifestCommitPlanUsesUpdateAndExactLastCommitId));
    }

    private static void ConfirmedMissingManifestCommitPlanUsesCreateWithoutLastCommitId()
    {
        SharedManifestCommitPlan plan = SharedManifestSafety.CreateCommitPlan(new SharedManifestProbeResult
        {
            State = SharedManifestProbeState.ConfirmedNotFound
        });
        Assert(plan.Action == "create" && string.IsNullOrWhiteSpace(plan.LastCommitId),
            nameof(ConfirmedMissingManifestCommitPlanUsesCreateWithoutLastCommitId));
    }

    private static Task<SharedManifestProbeResult> Probe(
        SharedManifestMetadataProbe metadata,
        Func<Task<SharedManifestContentProbe>> raw,
        Func<Task<SharedManifestContentProbe>> tree,
        Func<Task<string>> commit,
        bool validated = true)
    {
        return SharedManifestSafety.ProbeAsync(
            validated,
            "main",
            () => Task.FromResult(metadata),
            async () =>
            {
                string lastCommitId = await commit().ConfigureAwait(false);
                return new SharedManifestCommitProbe
                {
                    State = string.IsNullOrWhiteSpace(lastCommitId)
                        ? SharedManifestEndpointState.NotFound
                        : SharedManifestEndpointState.Found,
                    StatusCode = 200,
                    LastCommitId = lastCommitId
                };
            },
            async contentRef => AttachSnapshot(
                await raw().ConfigureAwait(false),
                contentRef,
                "raw"),
            async contentRef => AttachSnapshot(
                await tree().ConfigureAwait(false),
                contentRef,
                "tree-blob"),
            ParseManifest,
            null);
    }

    private static SharedManifestContentProbe AttachSnapshot(
        SharedManifestContentProbe probe,
        string contentRef,
        string route)
    {
        if (probe == null)
        {
            return null;
        }

        probe.ContentRef = contentRef;
        probe.LastCommitId = contentRef;
        probe.ContentRoute = route;
        return probe;
    }

    private static Task<SharedManifestProbeResult> ProbeWithMetadataException(
        Exception metadataError,
        SharedManifestCommitProbe commitProbe,
        Func<string, Task<SharedManifestContentProbe>> raw,
        Func<string, Task<SharedManifestContentProbe>> tree)
    {
        return SharedManifestSafety.ProbeAsync(
            true,
            "main",
            () => throw metadataError,
            () => Task.FromResult(commitProbe),
            raw,
            tree,
            ParseManifest,
            null);
    }

    private static SharedManifestCommitProbe CommitFound(string lastCommitId)
    {
        return new SharedManifestCommitProbe
        {
            State = SharedManifestEndpointState.Found,
            StatusCode = 200,
            LastCommitId = lastCommitId
        };
    }

    private static SharedManifestCommitProbe CommitNotFound()
    {
        return new SharedManifestCommitProbe
        {
            State = SharedManifestEndpointState.NotFound,
            StatusCode = 200
        };
    }

    private static SharedManifestContentProbe PinnedContentFound(
        string content,
        string contentRef,
        string route)
    {
        return new SharedManifestContentProbe
        {
            State = SharedManifestEndpointState.Found,
            StatusCode = 200,
            Content = Encoding.UTF8.GetBytes(content),
            LastCommitId = contentRef,
            ContentRoute = route,
            ContentRef = contentRef
        };
    }

    private static SharedManifestContentProbe PinnedContentNotFound(
        string contentRef,
        string route)
    {
        return new SharedManifestContentProbe
        {
            State = SharedManifestEndpointState.NotFound,
            StatusCode = 404,
            LastCommitId = contentRef,
            ContentRoute = route,
            ContentRef = contentRef
        };
    }

    private static List<GitLabTreeItem> CreateTreeItems(int count, string prefix)
    {
        var items = new List<GitLabTreeItem>();
        for (int index = 0; index < count; index++)
        {
            items.Add(new GitLabTreeItem
            {
                Id = prefix + index,
                Name = prefix + index + ".json",
                Type = "blob"
            });
        }

        return items;
    }

    private static SharedProjectManifest ParseManifest(byte[] bytes)
    {
        string text = Encoding.UTF8.GetString(bytes ?? new byte[0]);
        if (text != "valid")
        {
            throw new InvalidOperationException("invalid manifest json");
        }

        return ValidManifest();
    }

    private static SharedProjectManifest ValidManifest()
    {
        return new SharedProjectManifest { Project = "project" };
    }

    private static SharedManifestMetadataProbe MetadataFound(string content, string lastCommitId)
    {
        return new SharedManifestMetadataProbe
        {
            State = SharedManifestEndpointState.Found,
            Content = Convert.ToBase64String(Encoding.UTF8.GetBytes(content)),
            Encoding = "base64",
            LastCommitId = lastCommitId
        };
    }

    private static SharedManifestMetadataProbe MetadataNotFound()
    {
        return new SharedManifestMetadataProbe { State = SharedManifestEndpointState.NotFound };
    }

    private static Task<SharedManifestContentProbe> ContentFound(string content)
    {
        return Task.FromResult(new SharedManifestContentProbe
        {
            State = SharedManifestEndpointState.Found,
            Content = Encoding.UTF8.GetBytes(content)
        });
    }

    private static Task<SharedManifestContentProbe> ContentNotFound()
    {
        return Task.FromResult(new SharedManifestContentProbe
        {
            State = SharedManifestEndpointState.NotFound
        });
    }

    private static GitLabApiException ApiError(int statusCode, string body)
    {
        return new GitLabApiException(
            "GitLab API error " + statusCode,
            statusCode,
            "https://gitlab.example/api/v4/projects/1/repository/commits",
            body);
    }

    private static void AssertIndeterminateAndCreateRejected(
        SharedManifestProbeResult result,
        string scenario)
    {
        Assert(result.State == SharedManifestProbeState.Indeterminate, scenario + ": expected Indeterminate");
        bool rejected = false;
        try
        {
            SharedManifestSafety.CreateCommitPlan(result);
        }
        catch (InvalidOperationException)
        {
            rejected = true;
        }

        Assert(rejected, scenario + ": create must be rejected");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}
