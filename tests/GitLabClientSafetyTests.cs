using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

internal static class GitLabClientSafetyTests
{
    private const string MetadataUrl =
        "https://gitlab.example/api/v4/projects/1/repository/files/project%2F_manifest.json?ref=main";
    private const string CommitsUrl =
        "https://gitlab.example/api/v4/projects/1/repository/commits?ref_name=main&path=project%2F_manifest.json";
    private const string TreeUrl =
        "https://gitlab.example/api/v4/projects/1/repository/tree?path=project&ref=main&page=1";

    public static void Run()
    {
        TryTreeDownloadReturnsContent().GetAwaiter().GetResult();
        TryTreeDownloadReturnsNullOnlyForTypedTreeMiss().GetAwaiter().GetResult();
        TryTreeDownloadRethrowsBlobFailures().GetAwaiter().GetResult();
        TryTreeDownloadDoesNotTrustExceptionMessages().GetAwaiter().GetResult();
        OldGitLabMissingTreePathReturnsNullAfterPaging().GetAwaiter().GetResult();
        ValidatedTreePath404ReturnsTypedNotFoundAndNull().GetAwaiter().GetResult();
        TreePath404WithoutValidationRethrows().GetAwaiter().GetResult();
        TreePath404ValidationFailuresRethrow().GetAwaiter().GetResult();
        OnlyDirectTreePage404UsesValidation().GetAwaiter().GetResult();
        RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess();
        PathCommitResponseValidatesNullAndEntries();
        TreeResponseValidatesNullAndEntries();
        TreePagingRejectsNullPageLoaderResult().GetAwaiter().GetResult();
        InvalidSuccessfulResponseIsFailed();
        Console.WriteLine("GitLab client safety tests passed");
    }

    private static async Task TryTreeDownloadReturnsContent()
    {
        byte[] expected = Encoding.UTF8.GetBytes("content");
        byte[] actual = await GitLabTreePaging.TryDownloadAsync(
            () => Task.FromResult(expected));
        Assert(object.ReferenceEquals(expected, actual), nameof(TryTreeDownloadReturnsContent));
    }

    private static async Task TryTreeDownloadReturnsNullOnlyForTypedTreeMiss()
    {
        byte[] actual = await GitLabTreePaging.TryDownloadAsync(
            () => throw new GitLabTreeFileNotFoundException(
                "typed tree miss",
                2));
        Assert(actual == null, nameof(TryTreeDownloadReturnsNullOnlyForTypedTreeMiss));
    }

    private static async Task TryTreeDownloadRethrowsBlobFailures()
    {
        int[] statusCodes = { 404, 403, 500 };
        foreach (int statusCode in statusCodes)
        {
            int validationCalls = 0;
            var apiFailure = new GitLabApiException(
                "GitLab resource not found.",
                statusCode,
                "https://gitlab.example/api/v4/projects/1/repository/blobs/blob-id/raw",
                "failure");
            var blobFailure = new GitLabTreeBlobDownloadException(
                apiFailure.Message,
                2,
                2,
                apiFailure);

            GitLabTreeBlobDownloadException actual =
                await AssertThrowsAsync<GitLabTreeBlobDownloadException>(
                    async () =>
                    {
                        await GitLabTreePaging.TryDownloadAsync(async () =>
                        {
                            await GitLabTreePaging.FindBlobOrThrowAsync(
                                "_manifest.json",
                                100,
                                page => Task.FromResult(new List<GitLabTreeItem>
                                {
                                    new GitLabTreeItem
                                    {
                                        Id = "blob-id",
                                        Name = "_manifest.json",
                                        Type = "blob"
                                    }
                                }),
                                Validation(
                                    () => { validationCalls++; return Task.CompletedTask; },
                                    () => { validationCalls++; return Task.CompletedTask; }),
                                "missing");
                            throw blobFailure;
                        });
                    },
                    nameof(TryTreeDownloadRethrowsBlobFailures) + ": " + statusCode);
            Assert(object.ReferenceEquals(blobFailure, actual),
                nameof(TryTreeDownloadRethrowsBlobFailures) + ": " + statusCode);
            Assert(validationCalls == 0,
                nameof(TryTreeDownloadRethrowsBlobFailures) + ": validation " + statusCode);
        }
    }

    private static async Task OldGitLabMissingTreePathReturnsNullAfterPaging()
    {
        int validationCalls = 0;
        GitLabTreeFileNotFoundException firstMiss = null;
        byte[] firstResult = await GitLabTreePaging.TryDownloadAsync(async () =>
        {
            try
            {
                await GitLabTreePaging.FindBlobOrThrowAsync(
                    "_manifest.json",
                    100,
                    page => Task.FromResult(new List<GitLabTreeItem>()),
                    Validation(
                        () => { validationCalls++; return Task.CompletedTask; },
                        () => { validationCalls++; return Task.CompletedTask; }),
                    "missing after paging");
            }
            catch (GitLabTreeFileNotFoundException ex)
            {
                firstMiss = ex;
                throw;
            }

            return new byte[0];
        });
        Assert(firstResult == null &&
            firstMiss != null &&
            firstMiss.Reason == GitLabTreeNotFoundReason.FileNotFoundAfterPaging &&
            firstMiss.PagesChecked == 1 &&
            validationCalls == 0,
            nameof(OldGitLabMissingTreePathReturnsNullAfterPaging) + ": first empty page");

        int pages = 0;
        GitLabTreeFileNotFoundException secondMiss = null;
        byte[] secondResult = await GitLabTreePaging.TryDownloadAsync(async () =>
        {
            try
            {
                await GitLabTreePaging.FindBlobOrThrowAsync(
                    "_manifest.json",
                    100,
                    page =>
                    {
                        pages++;
                        return Task.FromResult(page == 1
                            ? CreateTreeItems(100)
                            : new List<GitLabTreeItem>());
                    },
                    null,
                    "missing after two pages");
            }
            catch (GitLabTreeFileNotFoundException ex)
            {
                secondMiss = ex;
                throw;
            }

            return new byte[0];
        });
        Assert(secondResult == null &&
            secondMiss != null &&
            secondMiss.PagesChecked == 2 &&
            pages == 2,
            nameof(OldGitLabMissingTreePathReturnsNullAfterPaging) + ": full page then empty page");
    }

    private static async Task ValidatedTreePath404ReturnsTypedNotFoundAndNull()
    {
        int projectValidationCalls = 0;
        int refValidationCalls = 0;
        int pageCalls = 0;
        var logs = new List<string>();
        GitLabTreeFileNotFoundException observed = null;
        GitLabTreePath404Validation validation = Validation(
            () => { projectValidationCalls++; return Task.CompletedTask; },
            () => { refValidationCalls++; return Task.CompletedTask; },
            logs.Add);

        byte[] result = await GitLabTreePaging.TryDownloadAsync(
            async () =>
            {
                try
                {
                    await GitLabTreePaging.FindBlobOrThrowAsync(
                        "_manifest.json",
                        100,
                        page =>
                        {
                            pageCalls++;
                            throw Tree404();
                        },
                        validation,
                        "missing");
                }
                catch (GitLabTreeFileNotFoundException ex)
                {
                    observed = ex;
                    throw;
                }

                return new byte[0];
            },
            logs.Add);

        Assert(result == null &&
            observed != null &&
            observed.Reason == GitLabTreeNotFoundReason.PathNotFoundAfterValidated404 &&
            observed.PagesChecked == 1,
            nameof(ValidatedTreePath404ReturnsTypedNotFoundAndNull) + ": typed path miss");
        Assert(pageCalls == 1 && projectValidationCalls == 1 && refValidationCalls == 1,
            nameof(ValidatedTreePath404ReturnsTypedNotFoundAndNull) + ": validate once");
        Assert(ContainsLog(logs, "validationAttempted=True") &&
            ContainsLog(logs, "projectValidated=True") &&
            ContainsLog(logs, "refValidated=True") &&
            ContainsLog(logs, "convertedToTypedNotFound=True") &&
            ContainsLog(logs, "finalAction=null") &&
            ContainsLog(logs, "projectId=1") &&
            ContainsLog(logs, "ref=main") &&
            ContainsLog(logs, "requestedPath=project"),
            nameof(ValidatedTreePath404ReturnsTypedNotFoundAndNull) + ": diagnostics");
    }

    private static async Task TreePath404WithoutValidationRethrows()
    {
        GitLabApiException expected = Tree404();
        GitLabApiException actual = await AssertThrowsAsync<GitLabApiException>(
            async () =>
            {
                await GitLabTreePaging.FindBlobOrThrowAsync(
                    "_manifest.json",
                    100,
                    page => throw expected,
                    null,
                    "missing");
            },
            nameof(TreePath404WithoutValidationRethrows));
        Assert(object.ReferenceEquals(expected, actual),
            nameof(TreePath404WithoutValidationRethrows));
    }

    private static async Task TreePath404ValidationFailuresRethrow()
    {
        int refCalls = 0;
        var projectMissing = new GitLabApiException(
            "project missing",
            404,
            "https://gitlab.example/api/v4/projects/1",
            "not found");
        Exception projectResult = await RunValidationFailure(
            () => throw projectMissing,
            () => { refCalls++; return Task.CompletedTask; },
            nameof(TreePath404ValidationFailuresRethrow) + ": project missing");
        Assert(object.ReferenceEquals(projectMissing, projectResult) && refCalls == 0,
            nameof(TreePath404ValidationFailuresRethrow) + ": project missing");

        var refMissing = new GitLabApiException(
            "ref missing",
            404,
            CommitsUrl,
            "not found");
        Exception refResult = await RunValidationFailure(
            () => Task.CompletedTask,
            () => throw refMissing,
            nameof(TreePath404ValidationFailuresRethrow) + ": ref missing");
        Assert(object.ReferenceEquals(refMissing, refResult),
            nameof(TreePath404ValidationFailuresRethrow) + ": ref missing");

        Exception[] failures =
        {
            new GitLabApiException("unauthorized", 401, MetadataUrl, "unauthorized"),
            new GitLabApiException("forbidden", 403, MetadataUrl, "forbidden"),
            new GitLabApiException("rate limited", 429, MetadataUrl, "rate limited"),
            new GitLabApiException("server error", 500, MetadataUrl, "server error"),
            new GitLabResponseFormatException(
                200,
                MetadataUrl,
                "a non-null project object",
                Json("null")),
            new System.IO.IOException("network failed")
        };
        foreach (Exception failure in failures)
        {
            Exception actual = await RunValidationFailure(
                () => throw failure,
                () => Task.CompletedTask,
                nameof(TreePath404ValidationFailuresRethrow) + ": " + failure.GetType().Name);
            Assert(object.ReferenceEquals(failure, actual),
                nameof(TreePath404ValidationFailuresRethrow) + ": " + failure.Message);
        }
    }

    private static async Task OnlyDirectTreePage404UsesValidation()
    {
        int validationCalls = 0;
        GitLabTreePath404Validation validation = Validation(
            () => { validationCalls++; return Task.CompletedTask; },
            () => { validationCalls++; return Task.CompletedTask; });
        var raw404 = new GitLabApiException(
            "raw missing",
            404,
            MetadataUrl + "/raw",
            "not found");
        GitLabApiException rawResult = await AssertThrowsAsync<GitLabApiException>(
            async () =>
            {
                await GitLabTreePaging.FindBlobOrThrowAsync(
                    "_manifest.json",
                    100,
                    page => throw raw404,
                    validation,
                    "missing");
            },
            nameof(OnlyDirectTreePage404UsesValidation));
        Assert(object.ReferenceEquals(raw404, rawResult) && validationCalls == 0,
            nameof(OnlyDirectTreePage404UsesValidation));
    }

    private static async Task TryTreeDownloadDoesNotTrustExceptionMessages()
    {
        string[] messages =
        {
            "File not found in tree. but this is not a typed tree miss",
            "An unrelated failure contains GitLab resource not found. text"
        };

        foreach (string message in messages)
        {
            var expected = new InvalidOperationException(message);
            InvalidOperationException actual = await AssertThrowsAsync<InvalidOperationException>(
                async () =>
                {
                    await GitLabTreePaging.TryDownloadAsync(() => throw expected);
                },
                nameof(TryTreeDownloadDoesNotTrustExceptionMessages));
            Assert(object.ReferenceEquals(expected, actual),
                nameof(TryTreeDownloadDoesNotTrustExceptionMessages) + ": " + message);
        }

        var ambiguousTree404 = new GitLabApiException(
            "GitLab resource not found.",
            404,
            TreeUrl,
            "not found");
        GitLabApiException propagated = await AssertThrowsAsync<GitLabApiException>(
            async () =>
            {
                await GitLabTreePaging.TryDownloadAsync(() => throw ambiguousTree404);
            },
            nameof(TryTreeDownloadDoesNotTrustExceptionMessages) + ": ambiguous tree API 404");
        Assert(object.ReferenceEquals(ambiguousTree404, propagated),
            nameof(TryTreeDownloadDoesNotTrustExceptionMessages) + ": ambiguous tree API 404");
    }

    private static void RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess()
    {
        GitLabRepositoryFileInfo notFound = GitLabResponseValidation.ReadRepositoryFileInfo(
            404,
            Json("not found"),
            MetadataUrl);
        Assert(notFound == null,
            nameof(RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess) + ": 404");

        GitLabRepositoryFileInfo normal = GitLabResponseValidation.ReadRepositoryFileInfo(
            200,
            Json("{\"file_path\":\"project/_manifest.json\"}"),
            MetadataUrl);
        Assert(normal != null && normal.FilePath == "project/_manifest.json",
            nameof(RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess) + ": normal object");

        AssertFormatFailure(
            () => GitLabResponseValidation.ReadRepositoryFileInfo(200, Json("null"), MetadataUrl),
            MetadataUrl,
            nameof(RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess) + ": JSON null");
        AssertFormatFailure(
            () => GitLabResponseValidation.ReadRepositoryFileInfo(200, new byte[0], MetadataUrl),
            MetadataUrl,
            nameof(RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess) + ": empty body");
        GitLabResponseFormatException malformed = AssertFormatFailure(
            () => GitLabResponseValidation.ReadRepositoryFileInfo(200, Json("{"), MetadataUrl),
            MetadataUrl,
            nameof(RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess) + ": malformed JSON");
        Assert(malformed.InnerException != null,
            nameof(RepositoryFileResponseDistinguishesNotFoundAndMalformedSuccess) + ": inner parse error");
    }

    private static void PathCommitResponseValidatesNullAndEntries()
    {
        string empty = GitLabResponseValidation.ReadPathLastCommitId(
            200,
            Json("[]"),
            CommitsUrl);
        Assert(empty == null, nameof(PathCommitResponseValidatesNullAndEntries) + ": empty array");

        string commitId = GitLabResponseValidation.ReadPathLastCommitId(
            200,
            Json("[{\"id\":\"commit-C\"}]"),
            CommitsUrl);
        Assert(commitId == "commit-C", nameof(PathCommitResponseValidatesNullAndEntries) + ": normal commit");

        string[] invalidResponses =
        {
            "",
            "null",
            "[null]",
            "[{}]",
            "[{\"id\":\"   \"}]",
            "[{\"id\":\"commit-C\"},null]",
            "{"
        };
        foreach (string response in invalidResponses)
        {
            AssertFormatFailure(
                () => GitLabResponseValidation.ReadPathLastCommitId(200, Json(response), CommitsUrl),
                CommitsUrl,
                nameof(PathCommitResponseValidatesNullAndEntries) + ": " + response);
        }
    }

    private static void TreeResponseValidatesNullAndEntries()
    {
        List<GitLabTreeItem> empty = GitLabResponseValidation.ReadTreePage(
            200,
            Json("[]"),
            TreeUrl);
        Assert(empty.Count == 0, nameof(TreeResponseValidatesNullAndEntries) + ": empty array");

        List<GitLabTreeItem> normal = GitLabResponseValidation.ReadTreePage(
            200,
            Json("[{\"id\":\"blob-id\",\"name\":\"_manifest.json\",\"type\":\"blob\"}]"),
            TreeUrl);
        Assert(normal.Count == 1 && normal[0].Path == null,
            nameof(TreeResponseValidatesNullAndEntries) + ": path remains optional");

        string[] invalidResponses =
        {
            "",
            "null",
            "[null]",
            "[{\"name\":\"file\",\"type\":\"blob\"}]",
            "[{\"id\":\"id\",\"type\":\"blob\"}]",
            "[{\"id\":\"id\",\"name\":\"file\"}]",
            "[{\"id\":\"blob-id\",\"name\":\"_manifest.json\",\"type\":\"blob\"},{}]",
            "{"
        };
        foreach (string response in invalidResponses)
        {
            AssertFormatFailure(
                () => GitLabResponseValidation.ReadTreePage(200, Json(response), TreeUrl),
                TreeUrl,
                nameof(TreeResponseValidatesNullAndEntries) + ": " + response);
        }
    }

    private static async Task TreePagingRejectsNullPageLoaderResult()
    {
        await AssertThrowsAsync<InvalidOperationException>(
            async () =>
            {
                await GitLabTreePaging.FindBlobAsync(
                    "_manifest.json",
                    100,
                    page => Task.FromResult<List<GitLabTreeItem>>(null));
            },
            nameof(TreePagingRejectsNullPageLoaderResult) + ": null page");

        GitLabTreeSearchResult empty = await GitLabTreePaging.FindBlobAsync(
            "_manifest.json",
            100,
            page => Task.FromResult(new List<GitLabTreeItem>()));
        Assert(empty.Target == null && empty.PagesChecked == 1,
            nameof(TreePagingRejectsNullPageLoaderResult) + ": normal empty page");
    }

    private static void InvalidSuccessfulResponseIsFailed()
    {
        GitLabResponseFormatException error = AssertFormatFailure(
            () => GitLabResponseValidation.ReadRepositoryFileInfo(200, Json("null"), MetadataUrl),
            MetadataUrl,
            nameof(InvalidSuccessfulResponseIsFailed));
        SharedManifestEndpointClassification classification =
            SharedManifestSafety.ClassifyEndpointException(error);
        Assert(classification.State == SharedManifestEndpointState.Failed &&
            classification.StatusCode == 200,
            nameof(InvalidSuccessfulResponseIsFailed));
    }

    private static GitLabResponseFormatException AssertFormatFailure(
        Action action,
        string expectedUrl,
        string scenario)
    {
        try
        {
            action();
        }
        catch (GitLabResponseFormatException ex)
        {
            Assert(ex.StatusCode == 200 && ex.Url == expectedUrl,
                scenario + ": structured status and URL");
            Assert(!string.IsNullOrWhiteSpace(ex.ExpectedFormat),
                scenario + ": expected format detail");
            Assert(ex.ResponseBody != null && ex.ResponseBody.Length <= 2003,
                scenario + ": bounded response body");
            return ex;
        }

        throw new InvalidOperationException(scenario + ": expected GitLabResponseFormatException");
    }

    private static async Task<TException> AssertThrowsAsync<TException>(
        Func<Task> action,
        string scenario)
        where TException : Exception
    {
        try
        {
            await action().ConfigureAwait(false);
        }
        catch (TException ex)
        {
            return ex;
        }

        throw new InvalidOperationException(
            scenario + ": expected " + typeof(TException).Name);
    }

    private static async Task<Exception> RunValidationFailure(
        Func<Task> validateProjectAsync,
        Func<Task> validateRefAsync,
        string scenario)
    {
        var logs = new List<string>();
        Exception actual = await AssertThrowsAsync<Exception>(
            async () =>
            {
                await GitLabTreePaging.FindBlobOrThrowAsync(
                    "_manifest.json",
                    100,
                    page => throw Tree404(),
                    Validation(validateProjectAsync, validateRefAsync, logs.Add),
                    "missing");
            },
            scenario);
        Assert(ContainsLog(logs, "validationAttempted=True") &&
            ContainsLog(logs, "finalAction=rethrow") &&
            !ContainsLog(logs, "convertedToTypedNotFound=True"),
            scenario + ": validation diagnostics");
        return actual;
    }

    private static GitLabTreePath404Validation Validation(
        Func<Task> validateProjectAsync,
        Func<Task> validateRefAsync,
        Action<string> log = null)
    {
        return new GitLabTreePath404Validation
        {
            ProjectId = "1",
            RefName = "main",
            RequestedPath = "project",
            ValidateProjectAsync = validateProjectAsync,
            ValidateRefAsync = validateRefAsync,
            Log = log
        };
    }

    private static GitLabApiException Tree404()
    {
        return new GitLabApiException(
            "tree path missing",
            404,
            TreeUrl,
            "not found");
    }

    private static List<GitLabTreeItem> CreateTreeItems(int count)
    {
        var items = new List<GitLabTreeItem>();
        for (int index = 0; index < count; index++)
        {
            items.Add(new GitLabTreeItem
            {
                Id = "blob-" + index,
                Name = "other-" + index + ".json",
                Type = "blob"
            });
        }

        return items;
    }

    private static bool ContainsLog(List<string> logs, string value)
    {
        foreach (string log in logs)
        {
            if (log != null && log.IndexOf(value, StringComparison.Ordinal) >= 0)
            {
                return true;
            }
        }

        return false;
    }

    private static byte[] Json(string value)
    {
        return Encoding.UTF8.GetBytes(value);
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition)
        {
            throw new InvalidOperationException(message);
        }
    }
}
