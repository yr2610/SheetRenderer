using System;
using System.Collections.Generic;
using System.Threading.Tasks;

internal sealed class GitLabTreeSearchResult
{
    public GitLabTreeItem Target { get; set; }
    public int PagesChecked { get; set; }
    public int? FoundPage { get; set; }
}

internal sealed class GitLabTreeBlobDownloadResult
{
    public byte[] Content { get; set; }
    public int PagesChecked { get; set; }
    public int FoundPage { get; set; }
}

internal enum GitLabTreeNotFoundReason
{
    FileNotFoundAfterPaging,
    PathNotFoundAfterValidated404
}

internal sealed class GitLabTreePath404Validation
{
    public string ProjectId { get; set; }
    public string RefName { get; set; }
    public string RequestedPath { get; set; }
    public Func<Task> ValidateProjectAsync { get; set; }
    public Func<Task> ValidateRefAsync { get; set; }
    public Action<string> Log { get; set; }
}

internal sealed class GitLabTreeFileNotFoundException : InvalidOperationException
{
    public GitLabTreeFileNotFoundException(string message, int pagesChecked)
        : this(
            message,
            pagesChecked,
            GitLabTreeNotFoundReason.FileNotFoundAfterPaging,
            null)
    {
    }

    public GitLabTreeFileNotFoundException(
        string message,
        int pagesChecked,
        GitLabTreeNotFoundReason reason,
        Exception innerException)
        : base(message, innerException)
    {
        PagesChecked = pagesChecked;
        Reason = reason;
    }

    public int PagesChecked { get; private set; }

    public GitLabTreeNotFoundReason Reason { get; private set; }
}

internal sealed class GitLabTreeBlobDownloadException : InvalidOperationException
{
    public GitLabTreeBlobDownloadException(
        string message,
        int pagesChecked,
        int foundPage,
        Exception innerException)
        : base(message, innerException)
    {
        PagesChecked = pagesChecked;
        FoundPage = foundPage;
    }

    public int PagesChecked { get; private set; }
    public int FoundPage { get; private set; }
}

internal static class GitLabTreePaging
{
    public static async Task<byte[]> TryDownloadAsync(
        Func<Task<byte[]>> download,
        Action<string> log = null)
    {
        if (download == null) throw new ArgumentNullException(nameof(download));

        try
        {
            return await download().ConfigureAwait(false);
        }
        catch (GitLabTreeFileNotFoundException ex)
        {
            if (ex.Reason == GitLabTreeNotFoundReason.PathNotFoundAfterValidated404)
            {
                Log(log,
                    "[GitLabTreePath404] convertedToTypedNotFound=true finalAction=null");
            }

            return null;
        }
    }

    public static async Task<GitLabTreeSearchResult> FindBlobOrThrowAsync(
        string fileName,
        int pageSize,
        Func<int, Task<List<GitLabTreeItem>>> pageLoader,
        GitLabTreePath404Validation path404Validation,
        string notFoundMessage)
    {
        GitLabTreeSearchResult search = await FindBlobWithValidatedPath404Async(
            fileName,
            pageSize,
            pageLoader,
            path404Validation).ConfigureAwait(false);
        if (search.Target == null || string.IsNullOrEmpty(search.Target.Id))
        {
            throw new GitLabTreeFileNotFoundException(
                notFoundMessage,
                search.PagesChecked);
        }

        return search;
    }

    public static async Task<GitLabTreeSearchResult> FindBlobWithValidatedPath404Async(
        string fileName,
        int pageSize,
        Func<int, Task<List<GitLabTreeItem>>> pageLoader,
        GitLabTreePath404Validation path404Validation)
    {
        if (pageLoader == null) throw new ArgumentNullException(nameof(pageLoader));

        int requestedPage = 0;
        try
        {
            return await FindBlobAsync(
                fileName,
                pageSize,
                page =>
                {
                    requestedPage = page;
                    return pageLoader(page);
                }).ConfigureAwait(false);
        }
        catch (GitLabApiException tree404)
        {
            if (!IsRepositoryTree404(tree404))
            {
                throw;
            }

            bool validationAttempted = path404Validation != null &&
                path404Validation.ValidateProjectAsync != null &&
                path404Validation.ValidateRefAsync != null;
            LogPath404(
                path404Validation,
                validationAttempted,
                false,
                false,
                false,
                "validation",
                null);
            if (!validationAttempted)
            {
                LogPath404(
                    path404Validation,
                    false,
                    false,
                    false,
                    false,
                    "rethrow",
                    null);
                throw;
            }

            bool projectValidated = false;
            bool refValidated = false;
            try
            {
                await path404Validation.ValidateProjectAsync().ConfigureAwait(false);
                projectValidated = true;
            }
            catch (Exception validationError)
            {
                LogPath404(
                    path404Validation,
                    true,
                    false,
                    false,
                    false,
                    "rethrow",
                    validationError);
                throw;
            }

            try
            {
                await path404Validation.ValidateRefAsync().ConfigureAwait(false);
                refValidated = true;
            }
            catch (Exception validationError)
            {
                LogPath404(
                    path404Validation,
                    true,
                    projectValidated,
                    false,
                    false,
                    "rethrow",
                    validationError);
                throw;
            }

            LogPath404(
                path404Validation,
                true,
                projectValidated,
                refValidated,
                true,
                "typed-not-found",
                null);
            throw new GitLabTreeFileNotFoundException(
                "GitLab repository tree path was not found after validating the project and ref.",
                requestedPage,
                GitLabTreeNotFoundReason.PathNotFoundAfterValidated404,
                tree404);
        }
    }

    public static async Task<GitLabTreeSearchResult> FindBlobAsync(
        string fileName,
        int pageSize,
        Func<int, Task<List<GitLabTreeItem>>> pageLoader)
    {
        if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("File name is required.", nameof(fileName));
        if (pageSize <= 0) throw new ArgumentOutOfRangeException(nameof(pageSize));
        if (pageLoader == null) throw new ArgumentNullException(nameof(pageLoader));

        int page = 1;
        while (true)
        {
            List<GitLabTreeItem> items = await pageLoader(page).ConfigureAwait(false);
            if (items == null)
            {
                throw new InvalidOperationException(
                    "GitLab repository tree page loader returned null for page " + page + ".");
            }

            foreach (GitLabTreeItem item in items)
            {
                if (item != null &&
                    string.Equals(item.Type, "blob", StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.Name, fileName, StringComparison.Ordinal))
                {
                    return new GitLabTreeSearchResult
                    {
                        Target = item,
                        PagesChecked = page,
                        FoundPage = page
                    };
                }
            }

            if (items.Count < pageSize)
            {
                return new GitLabTreeSearchResult
                {
                    PagesChecked = page
                };
            }

            page++;
        }
    }

    private static bool IsRepositoryTree404(GitLabApiException exception)
    {
        Uri uri;
        return exception != null &&
            exception.StatusCode == 404 &&
            Uri.TryCreate(exception.Url, UriKind.Absolute, out uri) &&
            uri.AbsolutePath.TrimEnd('/').EndsWith(
                "/repository/tree",
                StringComparison.OrdinalIgnoreCase);
    }

    private static void LogPath404(
        GitLabTreePath404Validation validation,
        bool validationAttempted,
        bool projectValidated,
        bool refValidated,
        bool convertedToTypedNotFound,
        string finalAction,
        Exception validationError)
    {
        if (validation == null || validation.Log == null)
        {
            return;
        }

        string errorDetail = string.Empty;
        GitLabApiException apiError = validationError as GitLabApiException;
        if (apiError != null)
        {
            errorDetail =
                " validationError=" + apiError.GetType().Name +
                " validationStatus=" + apiError.StatusCode;
        }
        else if (validationError is GitLabResponseFormatException)
        {
            GitLabResponseFormatException formatError =
                (GitLabResponseFormatException)validationError;
            errorDetail =
                " validationError=" + formatError.GetType().Name +
                " validationStatus=" + formatError.StatusCode;
        }
        else if (validationError != null)
        {
            errorDetail = " validationError=" + validationError.GetType().Name;
        }

        Log(validation.Log,
            "[GitLabTreePath404] route=tree status=404" +
            " projectId=" + NormalizeForLog(validation.ProjectId, 300) +
            " ref=" + NormalizeForLog(validation.RefName, 300) +
            " requestedPath=" + NormalizeForLog(validation.RequestedPath, 500) +
            " 404Source=tree-path" +
            " validationAttempted=" + validationAttempted +
            " projectValidated=" + projectValidated +
            " refValidated=" + refValidated +
            " convertedToTypedNotFound=" + convertedToTypedNotFound +
            " finalAction=" + finalAction +
            errorDetail);
    }

    private static string NormalizeForLog(string value, int maxLength)
    {
        string normalized = (value ?? string.Empty)
            .Replace('\r', ' ')
            .Replace('\n', ' ');
        return normalized.Length <= maxLength
            ? normalized
            : normalized.Substring(0, maxLength) + "...";
    }

    private static void Log(Action<string> log, string message)
    {
        if (log != null)
        {
            log(message);
        }
    }
}
