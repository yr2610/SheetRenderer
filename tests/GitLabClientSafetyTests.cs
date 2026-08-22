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
                        await GitLabTreePaging.TryDownloadAsync(() => throw blobFailure);
                    },
                    nameof(TryTreeDownloadRethrowsBlobFailures) + ": " + statusCode);
            Assert(object.ReferenceEquals(blobFailure, actual),
                nameof(TryTreeDownloadRethrowsBlobFailures) + ": " + statusCode);
        }
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
