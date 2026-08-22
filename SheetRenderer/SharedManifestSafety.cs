using System;
using System.Threading.Tasks;

internal enum SharedManifestProbeState
{
    Found,
    ConfirmedNotFound,
    Indeterminate
}

internal enum SharedManifestEndpointState
{
    Found,
    NotFound,
    Unsupported,
    Failed
}

internal sealed class SharedManifestMetadataProbe
{
    public SharedManifestEndpointState State { get; set; }
    public int? StatusCode { get; set; }
    public string Content { get; set; }
    public string Encoding { get; set; }
    public string LastCommitId { get; set; }
    public Exception Error { get; set; }
}

internal sealed class SharedManifestContentProbe
{
    public SharedManifestEndpointState State { get; set; }
    public int? StatusCode { get; set; }
    public byte[] Content { get; set; }
    public string LastCommitId { get; set; }
    public string ContentRoute { get; set; }
    public string ContentRef { get; set; }
    public int? PagesChecked { get; set; }
    public int? FoundPage { get; set; }
    public Exception Error { get; set; }
}

internal sealed class SharedManifestCommitProbe
{
    public SharedManifestEndpointState State { get; set; }
    public int? StatusCode { get; set; }
    public string LastCommitId { get; set; }
    public Exception Error { get; set; }
}

internal sealed class SharedManifestEndpointClassification
{
    public SharedManifestEndpointState State { get; set; }
    public int? StatusCode { get; set; }
    public Exception Error { get; set; }
}

internal sealed class SharedManifestProbeResult
{
    public SharedManifestProbeState State { get; set; }
    public SharedProjectManifest Manifest { get; set; }
    public string LastCommitId { get; set; }
    public string ContentRoute { get; set; }
    public string ContentRef { get; set; }
    public string DecisionReason { get; set; }
    public Exception Error { get; set; }
}

internal sealed class SharedManifestCommitPlan
{
    public string Action { get; set; }
    public string LastCommitId { get; set; }
}

internal enum SharedManifestCommitConflictKind
{
    None,
    UpdatedAfterRead,
    CreatedAfterProbe
}

internal sealed class SharedManifestCommitConflictException : InvalidOperationException
{
    internal const string UserMessage =
        "共有処理中に共有先が更新されました。ローカルの変更は失われていません。" +
        "もう一度「シート共有」を実行してください。最新状態との競合がある場合は差分確認画面が表示されます。";

    public SharedManifestCommitConflictException(
        SharedManifestCommitConflictKind conflictKind,
        GitLabApiException apiException)
        : base(UserMessage, apiException)
    {
        ConflictKind = conflictKind;
        ApiException = apiException;
    }

    public SharedManifestCommitConflictKind ConflictKind { get; private set; }

    public GitLabApiException ApiException { get; private set; }
}

internal static class SharedManifestSafety
{
    public static async Task<SharedManifestProbeResult> ProbeAsync(
        bool projectAndRefValidated,
        string branchRef,
        Func<Task<SharedManifestMetadataProbe>> metadataProbe,
        Func<Task<SharedManifestCommitProbe>> lastCommitIdProbe,
        Func<string, Task<SharedManifestContentProbe>> rawProbe,
        Func<string, Task<SharedManifestContentProbe>> treeBlobProbe,
        Func<byte[], SharedProjectManifest> manifestParser,
        Action<string> log)
    {
        if (string.IsNullOrWhiteSpace(branchRef)) throw new ArgumentException("Branch ref is required.", nameof(branchRef));
        if (metadataProbe == null) throw new ArgumentNullException(nameof(metadataProbe));
        if (lastCommitIdProbe == null) throw new ArgumentNullException(nameof(lastCommitIdProbe));
        if (rawProbe == null) throw new ArgumentNullException(nameof(rawProbe));
        if (treeBlobProbe == null) throw new ArgumentNullException(nameof(treeBlobProbe));
        if (manifestParser == null) throw new ArgumentNullException(nameof(manifestParser));

        SharedManifestMetadataProbe metadata = await InvokeMetadataProbeAsync(metadataProbe).ConfigureAwait(false);
        LogEndpoint(log, "metadata", metadata.State, metadata.StatusCode, metadata.Error);
        if (metadata.State == SharedManifestEndpointState.Failed)
        {
            return Indeterminate(metadata.Error, "metadata-failed", log);
        }

        bool metadataContentWasUnusable = false;
        if (metadata.State == SharedManifestEndpointState.Found)
        {
            byte[] metadataBytes;
            if (TryDecodeMetadataContent(metadata, out metadataBytes) &&
                !string.IsNullOrWhiteSpace(metadata.LastCommitId))
            {
                LogSkipped(log, "raw");
                LogSkipped(log, "tree-blob");
                LogSkipped(log, "path-commits");
                return BuildFoundResult(
                    metadataBytes,
                    metadata.LastCommitId,
                    "metadata",
                    metadata.LastCommitId,
                    manifestParser,
                    log);
            }

            metadataContentWasUnusable = true;
            Log(log,
                "[SharedManifestProbe] route=metadata snapshot=unusable fallback=true" +
                " hasContent=" + (metadataBytes != null && metadataBytes.Length > 0) +
                " hasLastCommitId=" + (!string.IsNullOrWhiteSpace(metadata.LastCommitId)));
        }

        string pinnedCommitId = metadata.State == SharedManifestEndpointState.Found
            ? metadata.LastCommitId
            : null;
        SharedManifestCommitProbe commitProbe = null;
        if (string.IsNullOrWhiteSpace(pinnedCommitId))
        {
            commitProbe = await InvokeCommitProbeAsync(lastCommitIdProbe).ConfigureAwait(false);
            LogCommitEndpoint(log, commitProbe);
            if (commitProbe.State == SharedManifestEndpointState.Failed)
            {
                return Indeterminate(commitProbe.Error, "path-commits-failed", log);
            }

            if (commitProbe.State == SharedManifestEndpointState.Found &&
                !string.IsNullOrWhiteSpace(commitProbe.LastCommitId))
            {
                pinnedCommitId = commitProbe.LastCommitId;
            }
        }
        else
        {
            LogSkipped(log, "path-commits");
        }

        if (!string.IsNullOrWhiteSpace(pinnedCommitId))
        {
            Log(log,
                "[SharedManifestProbe] pinnedCommitId=" + pinnedCommitId +
                " contentRefType=commit");

            SharedManifestContentProbe pinnedRaw = await InvokeContentProbeAsync(
                rawProbe,
                pinnedCommitId).ConfigureAwait(false);
            LogContentEndpoint(log, "raw", pinnedRaw);
            if (pinnedRaw.State == SharedManifestEndpointState.Failed)
            {
                return Indeterminate(pinnedRaw.Error, "pinned-raw-failed", log);
            }

            if (pinnedRaw.State == SharedManifestEndpointState.Found)
            {
                return BuildPinnedFoundResult(
                    pinnedRaw,
                    pinnedCommitId,
                    manifestParser,
                    log);
            }

            SharedManifestContentProbe pinnedTree = await InvokeContentProbeAsync(
                treeBlobProbe,
                pinnedCommitId).ConfigureAwait(false);
            LogContentEndpoint(log, "tree-blob", pinnedTree);
            if (pinnedTree.State == SharedManifestEndpointState.Failed)
            {
                return Indeterminate(pinnedTree.Error, "pinned-tree-blob-failed", log);
            }

            if (pinnedTree.State == SharedManifestEndpointState.Found)
            {
                return BuildPinnedFoundResult(
                    pinnedTree,
                    pinnedCommitId,
                    manifestParser,
                    log);
            }

            return Indeterminate(
                new InvalidOperationException(
                    "共有マニフェストのコミットIDは取得できましたが、同じコミット時点の内容を取得できませんでした。"),
                "pinned-content-unavailable",
                log);
        }

        Log(log,
            "[SharedManifestProbe] pinnedCommitId=none contentRef=" + branchRef +
            " contentRefType=branch existenceCheck=true");

        SharedManifestContentProbe branchRaw = await InvokeContentProbeAsync(
            rawProbe,
            branchRef).ConfigureAwait(false);
        LogContentEndpoint(log, "raw", branchRaw);
        if (branchRaw.State == SharedManifestEndpointState.Failed)
        {
            return Indeterminate(branchRaw.Error, "branch-raw-failed", log);
        }

        if (branchRaw.State == SharedManifestEndpointState.Found)
        {
            return Indeterminate(
                new InvalidOperationException(
                    "共有マニフェストの内容は存在しますが、ファイル固有のコミットIDを取得できませんでした。"),
                "content-found-without-pinned-commit",
                log);
        }

        SharedManifestContentProbe branchTree = await InvokeContentProbeAsync(
            treeBlobProbe,
            branchRef).ConfigureAwait(false);
        LogContentEndpoint(log, "tree-blob", branchTree);
        if (branchTree.State == SharedManifestEndpointState.Failed)
        {
            return Indeterminate(branchTree.Error, "branch-tree-blob-failed", log);
        }

        if (branchTree.State == SharedManifestEndpointState.Found)
        {
            return Indeterminate(
                new InvalidOperationException(
                    "共有マニフェストの内容は存在しますが、ファイル固有のコミットIDを取得できませんでした。"),
                "content-found-without-pinned-commit",
                log);
        }

        bool commitHistoryUnavailableOrEmpty = commitProbe != null &&
            (commitProbe.State == SharedManifestEndpointState.NotFound ||
             commitProbe.State == SharedManifestEndpointState.Unsupported);
        bool metadataAllowsAbsenceCheck = !metadataContentWasUnusable &&
            (metadata.State == SharedManifestEndpointState.NotFound ||
             metadata.State == SharedManifestEndpointState.Unsupported);
        bool rawAllowsAbsenceCheck =
            branchRaw.State == SharedManifestEndpointState.NotFound ||
            branchRaw.State == SharedManifestEndpointState.Unsupported;
        bool completeTreeConfirmedNotFound = branchTree.State == SharedManifestEndpointState.NotFound;

        if (projectAndRefValidated &&
            commitHistoryUnavailableOrEmpty &&
            metadataAllowsAbsenceCheck &&
            rawAllowsAbsenceCheck &&
            completeTreeConfirmedNotFound)
        {
            var notFound = new SharedManifestProbeResult
            {
                State = SharedManifestProbeState.ConfirmedNotFound,
                DecisionReason = "validated-project-ref; commit-history-empty-or-unsupported; complete-tree-not-found"
            };
            LogFinal(log, notFound);
            return notFound;
        }

        return Indeterminate(
            new InvalidOperationException("共有マニフェストの存在または内容を安全に確認できませんでした。"),
            "absence-not-confirmed",
            log);
    }

    public static SharedManifestCommitPlan CreateCommitPlan(SharedManifestProbeResult probeResult)
    {
        if (probeResult == null)
        {
            throw new InvalidOperationException("共有マニフェストの取得結果がありません。");
        }

        if (probeResult.State == SharedManifestProbeState.ConfirmedNotFound)
        {
            return new SharedManifestCommitPlan { Action = "create" };
        }

        if (probeResult.State == SharedManifestProbeState.Found &&
            probeResult.Manifest != null &&
            !string.IsNullOrWhiteSpace(probeResult.LastCommitId))
        {
            return new SharedManifestCommitPlan
            {
                Action = "update",
                LastCommitId = probeResult.LastCommitId
            };
        }

        throw new InvalidOperationException(
            "共有マニフェストの状態を安全に確認できないため、共有処理を中止しました。");
    }

    public static SharedManifestCommitConflictKind ClassifyCommitConflict(
        GitLabApiException exception,
        string manifestAction)
    {
        if (exception == null || (exception.StatusCode != 400 && exception.StatusCode != 409))
        {
            return SharedManifestCommitConflictKind.None;
        }

        string body = exception.ResponseBody ?? string.Empty;
        if (string.Equals(manifestAction, "update", StringComparison.OrdinalIgnoreCase) &&
            (Contains(body, "has changed since you started editing") ||
             Contains(body, "last_commit_id does not match") ||
             Contains(body, "last commit id does not match")))
        {
            return SharedManifestCommitConflictKind.UpdatedAfterRead;
        }

        if (string.Equals(manifestAction, "create", StringComparison.OrdinalIgnoreCase) &&
            (Contains(body, "a file with this name already exists") ||
             Contains(body, "file already exists")))
        {
            return SharedManifestCommitConflictKind.CreatedAfterProbe;
        }

        return SharedManifestCommitConflictKind.None;
    }

    public static SharedManifestContentProbe ClassifyTreeBlobException(Exception exception)
    {
        GitLabTreeFileNotFoundException treeNotFound = exception as GitLabTreeFileNotFoundException;
        if (treeNotFound != null)
        {
            return new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.NotFound,
                StatusCode = 404,
                ContentRoute = "tree-blob",
                PagesChecked = treeNotFound.PagesChecked
            };
        }

        GitLabApiException apiException = exception as GitLabApiException;
        if (apiException != null &&
            apiException.StatusCode == 404 &&
            Contains(apiException.Url ?? string.Empty, "/repository/tree"))
        {
            return new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.NotFound,
                StatusCode = 404,
                ContentRoute = "tree-blob"
            };
        }

        InvalidOperationException invalidOperationException = exception as InvalidOperationException;
        if (invalidOperationException != null &&
            invalidOperationException.Message != null &&
            invalidOperationException.Message.StartsWith(
                "File not found in tree.",
                StringComparison.Ordinal))
        {
            return new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.NotFound,
                StatusCode = 404,
                ContentRoute = "tree-blob"
            };
        }

        GitLabTreeBlobDownloadException blobFailure = exception as GitLabTreeBlobDownloadException;
        SharedManifestEndpointClassification classification = ClassifyEndpointException(
            blobFailure == null ? exception : blobFailure.InnerException);
        return new SharedManifestContentProbe
        {
            State = classification.State,
            StatusCode = classification.StatusCode,
            ContentRoute = "tree-blob",
            PagesChecked = blobFailure == null ? (int?)null : blobFailure.PagesChecked,
            FoundPage = blobFailure == null ? (int?)null : blobFailure.FoundPage,
            Error = exception
        };
    }

    public static SharedManifestEndpointClassification ClassifyEndpointException(Exception exception)
    {
        GitLabApiException apiException = FindGitLabApiException(exception);
        if (apiException != null)
        {
            if (apiException.StatusCode == 405 || apiException.StatusCode == 501)
            {
                return new SharedManifestEndpointClassification
                {
                    State = SharedManifestEndpointState.Unsupported,
                    StatusCode = apiException.StatusCode,
                    Error = exception
                };
            }

            if (apiException.StatusCode == 400 &&
                IsExplicitUnsupportedResponse(apiException.ResponseBody))
            {
                return new SharedManifestEndpointClassification
                {
                    State = SharedManifestEndpointState.Unsupported,
                    StatusCode = apiException.StatusCode,
                    Error = exception
                };
            }

            return new SharedManifestEndpointClassification
            {
                State = SharedManifestEndpointState.Failed,
                StatusCode = apiException.StatusCode,
                Error = exception
            };
        }

        return new SharedManifestEndpointClassification
        {
            State = SharedManifestEndpointState.Failed,
            Error = exception
        };
    }

    private static bool IsExplicitUnsupportedResponse(string responseBody)
    {
        string body = responseBody ?? string.Empty;
        return Contains(body, "unsupported parameter") ||
            Contains(body, "unknown parameter") ||
            Contains(body, "parameter is not supported") ||
            Contains(body, "parameter is unsupported") ||
            Contains(body, "endpoint is not supported") ||
            Contains(body, "unsupported endpoint") ||
            Contains(body, "feature is not supported") ||
            Contains(body, "unsupported feature") ||
            Contains(body, "not implemented");
    }

    private static SharedManifestProbeResult BuildPinnedFoundResult(
        SharedManifestContentProbe contentProbe,
        string pinnedCommitId,
        Func<byte[], SharedProjectManifest> manifestParser,
        Action<string> log)
    {
        if (contentProbe == null ||
            contentProbe.Content == null ||
            contentProbe.Content.Length == 0 ||
            !string.Equals(contentProbe.LastCommitId, pinnedCommitId, StringComparison.Ordinal) ||
            !string.Equals(contentProbe.ContentRef, pinnedCommitId, StringComparison.Ordinal))
        {
            return Indeterminate(
                new InvalidOperationException(
                    "共有マニフェストの内容とファイル固有コミットIDが同じスナップショットではありません。"),
                "pinned-snapshot-mismatch",
                log);
        }

        return BuildFoundResult(
            contentProbe.Content,
            pinnedCommitId,
            contentProbe.ContentRoute,
            contentProbe.ContentRef,
            manifestParser,
            log);
    }

    private static SharedManifestProbeResult BuildFoundResult(
        byte[] content,
        string lastCommitId,
        string contentRoute,
        string contentRef,
        Func<byte[], SharedProjectManifest> manifestParser,
        Action<string> log)
    {
        SharedProjectManifest manifest;
        try
        {
            manifest = manifestParser(content);
        }
        catch (Exception ex)
        {
            return Indeterminate(ex, "manifest-json-invalid", log);
        }

        if (manifest == null)
        {
            return Indeterminate(
                new InvalidOperationException("共有マニフェストのJSON形式が正しくありません。"),
                "manifest-json-invalid",
                log);
        }

        if (string.IsNullOrWhiteSpace(lastCommitId))
        {
            return Indeterminate(
                new InvalidOperationException("共有マニフェストのファイル固有の最終コミットIDを取得できませんでした。"),
                "last-commit-id-missing",
                log);
        }

        var found = new SharedManifestProbeResult
        {
            State = SharedManifestProbeState.Found,
            Manifest = manifest,
            LastCommitId = lastCommitId,
            ContentRoute = contentRoute,
            ContentRef = contentRef,
            DecisionReason = "content-and-last-commit-from-same-snapshot"
        };
        LogFinal(log, found);
        return found;
    }

    private static bool TryDecodeMetadataContent(
        SharedManifestMetadataProbe metadata,
        out byte[] content)
    {
        content = null;
        if (metadata == null ||
            !string.Equals(metadata.Encoding, "base64", StringComparison.OrdinalIgnoreCase) ||
            string.IsNullOrWhiteSpace(metadata.Content))
        {
            return false;
        }

        try
        {
            content = Convert.FromBase64String(metadata.Content);
            return content.Length > 0;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    private static async Task<SharedManifestMetadataProbe> InvokeMetadataProbeAsync(
        Func<Task<SharedManifestMetadataProbe>> probe)
    {
        try
        {
            return await probe().ConfigureAwait(false) ?? new SharedManifestMetadataProbe
            {
                State = SharedManifestEndpointState.Failed,
                Error = new InvalidOperationException("Metadata probe returned no result.")
            };
        }
        catch (Exception ex)
        {
            SharedManifestEndpointClassification classification = ClassifyEndpointException(ex);
            return new SharedManifestMetadataProbe
            {
                State = classification.State,
                StatusCode = classification.StatusCode,
                Error = ex
            };
        }
    }

    private static async Task<SharedManifestCommitProbe> InvokeCommitProbeAsync(
        Func<Task<SharedManifestCommitProbe>> probe)
    {
        try
        {
            return await probe().ConfigureAwait(false) ?? new SharedManifestCommitProbe
            {
                State = SharedManifestEndpointState.Failed,
                Error = new InvalidOperationException("Commit probe returned no result.")
            };
        }
        catch (Exception ex)
        {
            SharedManifestEndpointClassification classification = ClassifyEndpointException(ex);
            return new SharedManifestCommitProbe
            {
                State = classification.State,
                StatusCode = classification.StatusCode,
                Error = ex
            };
        }
    }

    private static async Task<SharedManifestContentProbe> InvokeContentProbeAsync(
        Func<string, Task<SharedManifestContentProbe>> probe,
        string contentRef)
    {
        try
        {
            return await probe(contentRef).ConfigureAwait(false) ?? new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.Failed,
                ContentRef = contentRef,
                Error = new InvalidOperationException("Content probe returned no result.")
            };
        }
        catch (Exception ex)
        {
            SharedManifestEndpointClassification classification = ClassifyEndpointException(ex);
            return new SharedManifestContentProbe
            {
                State = classification.State,
                StatusCode = classification.StatusCode,
                ContentRef = contentRef,
                Error = ex
            };
        }
    }

    private static SharedManifestProbeResult Indeterminate(
        Exception error,
        string reason,
        Action<string> log)
    {
        var result = new SharedManifestProbeResult
        {
            State = SharedManifestProbeState.Indeterminate,
            DecisionReason = reason,
            Error = error
        };
        LogFinal(log, result);
        return result;
    }

    private static void LogEndpoint(
        Action<string> log,
        string route,
        SharedManifestEndpointState state,
        int? statusCode,
        Exception error)
    {
        string detail = string.Empty;
        GitLabApiException apiError = FindGitLabApiException(error);
        if (apiError != null)
        {
            detail =
                " status=" + apiError.StatusCode +
                " url=" + apiError.Url +
                " body=" + NormalizeForLog(apiError.ResponseBody, 1000);
        }
        else if (error != null)
        {
            detail = " error=" + error.GetType().Name +
                " message=" + NormalizeForLog(error.Message, 500);
        }
        else if (statusCode.HasValue)
        {
            detail = " status=" + statusCode.Value;
        }

        Log(log, "[SharedManifestProbe] route=" + route + " result=" + state + detail);
    }

    private static void LogCommitEndpoint(Action<string> log, SharedManifestCommitProbe probe)
    {
        LogEndpoint(log, "path-commits", probe.State, probe.StatusCode, probe.Error);
        Log(log,
            "[SharedManifestProbe] route=path-commits hasPinnedCommitId=" +
            (!string.IsNullOrWhiteSpace(probe.LastCommitId)) +
            " pinnedCommitId=" + (probe.LastCommitId ?? "none"));
    }

    private static void LogContentEndpoint(
        Action<string> log,
        string route,
        SharedManifestContentProbe probe)
    {
        LogEndpoint(log, route, probe.State, probe.StatusCode, probe.Error);
        Log(log,
            "[SharedManifestProbe] route=" + route +
            " contentRef=" + (probe.ContentRef ?? "none") +
            " snapshotLastCommitId=" + (probe.LastCommitId ?? "none") +
            " pagesChecked=" + (probe.PagesChecked.HasValue ? probe.PagesChecked.Value.ToString() : "n/a") +
            " foundPage=" + (probe.FoundPage.HasValue ? probe.FoundPage.Value.ToString() : "n/a"));
    }

    private static void LogSkipped(Action<string> log, string route)
    {
        Log(log, "[SharedManifestProbe] route=" + route + " result=Skipped");
    }

    private static void LogFinal(Action<string> log, SharedManifestProbeResult result)
    {
        Log(log,
            "[SharedManifestSnapshot] state=" + result.State +
            " contentRoute=" + (result.ContentRoute ?? "none") +
            " contentRef=" + (result.ContentRef ?? "none") +
            " hasLastCommitId=" + (!string.IsNullOrWhiteSpace(result.LastCommitId)) +
            " lastCommitId=" + (result.LastCommitId ?? "none") +
            " reason=" + (result.DecisionReason ?? "none"));
    }

    private static GitLabApiException FindGitLabApiException(Exception exception)
    {
        Exception current = exception;
        while (current != null)
        {
            GitLabApiException apiException = current as GitLabApiException;
            if (apiException != null)
            {
                return apiException;
            }

            current = current.InnerException;
        }

        return null;
    }

    internal static string NormalizeForLog(string value, int maxLength)
    {
        string normalized = (value ?? string.Empty)
            .Replace('\r', ' ')
            .Replace('\n', ' ');
        return normalized.Length <= maxLength
            ? normalized
            : normalized.Substring(0, maxLength) + "...";
    }

    private static bool Contains(string value, string expected)
    {
        return value.IndexOf(expected, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void Log(Action<string> log, string message)
    {
        if (log != null)
        {
            log(message);
        }
    }
}
