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
    public Exception Error { get; set; }
}

internal sealed class SharedManifestProbeResult
{
    public SharedManifestProbeState State { get; set; }
    public SharedProjectManifest Manifest { get; set; }
    public string LastCommitId { get; set; }
    public string ContentRoute { get; set; }
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
        Func<Task<SharedManifestMetadataProbe>> metadataProbe,
        Func<Task<SharedManifestContentProbe>> rawProbe,
        Func<Task<SharedManifestContentProbe>> treeBlobProbe,
        Func<Task<string>> lastCommitIdProbe,
        Func<byte[], SharedProjectManifest> manifestParser,
        Action<string> log)
    {
        if (metadataProbe == null) throw new ArgumentNullException(nameof(metadataProbe));
        if (rawProbe == null) throw new ArgumentNullException(nameof(rawProbe));
        if (treeBlobProbe == null) throw new ArgumentNullException(nameof(treeBlobProbe));
        if (lastCommitIdProbe == null) throw new ArgumentNullException(nameof(lastCommitIdProbe));
        if (manifestParser == null) throw new ArgumentNullException(nameof(manifestParser));

        SharedManifestMetadataProbe metadata = await InvokeMetadataProbeAsync(metadataProbe).ConfigureAwait(false);
        LogEndpoint(log, "metadata", metadata.State, metadata.StatusCode, metadata.Error);
        if (metadata.State == SharedManifestEndpointState.Failed)
        {
            return Indeterminate(metadata.Error, log);
        }

        bool metadataContentWasUnusable = false;
        if (metadata.State == SharedManifestEndpointState.Found)
        {
            byte[] metadataBytes;
            if (TryDecodeMetadataContent(metadata, out metadataBytes))
            {
                LogSkipped(log, "raw");
                LogSkipped(log, "tree-blob");
                return await BuildFoundResultAsync(
                    metadataBytes,
                    metadata.LastCommitId,
                    "metadata",
                    lastCommitIdProbe,
                    manifestParser,
                    log).ConfigureAwait(false);
            }

            metadataContentWasUnusable = true;
            Log(log, "[SharedManifestProbe] route=metadata content=unusable fallback=true");
        }

        SharedManifestContentProbe raw = await InvokeContentProbeAsync(rawProbe).ConfigureAwait(false);
        LogEndpoint(log, "raw", raw.State, raw.StatusCode, raw.Error);
        if (raw.State == SharedManifestEndpointState.Failed)
        {
            return Indeterminate(raw.Error, log);
        }

        bool rawContentWasUnusable = raw.State == SharedManifestEndpointState.Found &&
            (raw.Content == null || raw.Content.Length == 0);
        if (raw.State == SharedManifestEndpointState.Found && !rawContentWasUnusable)
        {
            LogSkipped(log, "tree-blob");
            return await BuildFoundResultAsync(
                raw.Content,
                metadata == null ? null : metadata.LastCommitId,
                "raw",
                lastCommitIdProbe,
                manifestParser,
                log).ConfigureAwait(false);
        }

        if (rawContentWasUnusable)
        {
            Log(log, "[SharedManifestProbe] route=raw content=unusable fallback=true");
        }

        SharedManifestContentProbe treeBlob = await InvokeContentProbeAsync(treeBlobProbe).ConfigureAwait(false);
        LogEndpoint(log, "tree-blob", treeBlob.State, treeBlob.StatusCode, treeBlob.Error);
        if (treeBlob.State == SharedManifestEndpointState.Failed)
        {
            return Indeterminate(treeBlob.Error, log);
        }

        bool treeContentWasUnusable = treeBlob.State == SharedManifestEndpointState.Found &&
            (treeBlob.Content == null || treeBlob.Content.Length == 0);
        if (treeBlob.State == SharedManifestEndpointState.Found && !treeContentWasUnusable)
        {
            return await BuildFoundResultAsync(
                treeBlob.Content,
                metadata == null ? null : metadata.LastCommitId,
                "tree-blob",
                lastCommitIdProbe,
                manifestParser,
                log).ConfigureAwait(false);
        }

        if (!metadataContentWasUnusable &&
            !rawContentWasUnusable &&
            !treeContentWasUnusable &&
            metadata.State == SharedManifestEndpointState.NotFound &&
            raw.State == SharedManifestEndpointState.NotFound &&
            treeBlob.State == SharedManifestEndpointState.NotFound &&
            projectAndRefValidated)
        {
            var notFound = new SharedManifestProbeResult
            {
                State = SharedManifestProbeState.ConfirmedNotFound
            };
            LogFinal(log, notFound);
            return notFound;
        }

        return Indeterminate(
            new InvalidOperationException("共有マニフェストの存在または内容を安全に確認できませんでした。"),
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
        GitLabApiException apiException = exception as GitLabApiException;
        if (apiException != null &&
            apiException.StatusCode == 404 &&
            Contains(apiException.Url ?? string.Empty, "/repository/tree"))
        {
            return new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.NotFound,
                StatusCode = 404
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
                StatusCode = 404
            };
        }

        return new SharedManifestContentProbe
        {
            State = SharedManifestEndpointState.Failed,
            Error = exception
        };
    }

    private static async Task<SharedManifestProbeResult> BuildFoundResultAsync(
        byte[] content,
        string lastCommitId,
        string contentRoute,
        Func<Task<string>> lastCommitIdProbe,
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
            return Indeterminate(ex, log);
        }

        if (manifest == null)
        {
            return Indeterminate(
                new InvalidOperationException("共有マニフェストのJSON形式が正しくありません。"),
                log);
        }

        if (string.IsNullOrWhiteSpace(lastCommitId))
        {
            try
            {
                lastCommitId = await lastCommitIdProbe().ConfigureAwait(false);
                Log(log,
                    "[SharedManifestProbe] route=path-commits result=" +
                    (string.IsNullOrWhiteSpace(lastCommitId) ? "Empty" : "Found") +
                    " status=200");
            }
            catch (Exception ex)
            {
                LogEndpoint(log, "path-commits", SharedManifestEndpointState.Failed, null, ex);
                return Indeterminate(ex, log);
            }
        }

        if (string.IsNullOrWhiteSpace(lastCommitId))
        {
            return Indeterminate(
                new InvalidOperationException("共有マニフェストのファイル固有の最終コミットIDを取得できませんでした。"),
                log);
        }

        var found = new SharedManifestProbeResult
        {
            State = SharedManifestProbeState.Found,
            Manifest = manifest,
            LastCommitId = lastCommitId,
            ContentRoute = contentRoute
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
            return new SharedManifestMetadataProbe
            {
                State = SharedManifestEndpointState.Failed,
                Error = ex
            };
        }
    }

    private static async Task<SharedManifestContentProbe> InvokeContentProbeAsync(
        Func<Task<SharedManifestContentProbe>> probe)
    {
        try
        {
            return await probe().ConfigureAwait(false) ?? new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.Failed,
                Error = new InvalidOperationException("Content probe returned no result.")
            };
        }
        catch (Exception ex)
        {
            return new SharedManifestContentProbe
            {
                State = SharedManifestEndpointState.Failed,
                Error = ex
            };
        }
    }

    private static SharedManifestProbeResult Indeterminate(Exception error, Action<string> log)
    {
        var result = new SharedManifestProbeResult
        {
            State = SharedManifestProbeState.Indeterminate,
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
        GitLabApiException apiError = error as GitLabApiException;
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

    private static void LogSkipped(Action<string> log, string route)
    {
        Log(log, "[SharedManifestProbe] route=" + route + " result=Skipped");
    }

    private static void LogFinal(Action<string> log, SharedManifestProbeResult result)
    {
        Log(log,
            "[SharedManifestSnapshot] state=" + result.State +
            " contentRoute=" + (result.ContentRoute ?? "none") +
            " hasLastCommitId=" + (!string.IsNullOrWhiteSpace(result.LastCommitId)));
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
