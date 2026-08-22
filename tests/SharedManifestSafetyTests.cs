using System;
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
        UpdateLastCommitMismatchIsClassifiedAsConcurrentUpdate();
        InitialCreateExistingFileIsClassifiedAsConcurrentCreate();
        UnrelatedHttp400IsNotClassifiedAsConflict();
        ExistingManifestCommitPlanUsesUpdateAndExactLastCommitId();
        ConfirmedMissingManifestCommitPlanUsesCreateWithoutLastCommitId();
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
        SharedManifestContentProbe classified =
            SharedManifestSafety.ClassifyTreeBlobException(blobFailure);
        Assert(classified.State == SharedManifestEndpointState.Failed,
            nameof(TreeBlobReadFailureIsIndeterminateAndNotCreate) + ": blob 404 must remain a failure");

        SharedManifestProbeResult result = await Probe(
            MetadataNotFound(),
            ContentNotFound,
            () => Task.FromResult(classified),
            () => Task.FromResult<string>(null));
        AssertIndeterminateAndCreateRejected(result, nameof(TreeBlobReadFailureIsIndeterminateAndNotCreate));

        SharedManifestContentProbe absentFromTree = SharedManifestSafety.ClassifyTreeBlobException(
            new InvalidOperationException("File not found in tree. folder=project file=_manifest.json ref=main"));
        Assert(absentFromTree.State == SharedManifestEndpointState.NotFound,
            nameof(TreeBlobReadFailureIsIndeterminateAndNotCreate) + ": absent tree entry is a confirmed route miss");
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
            () => Task.FromResult(metadata),
            raw,
            tree,
            commit,
            ParseManifest,
            null);
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
