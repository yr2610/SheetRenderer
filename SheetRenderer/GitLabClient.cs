using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.Json;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

public static class GitLabClient
{
    // 既に入れてるやつでOK
    private static void EnsureTls12()
    {
        ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
    }

    private static readonly HttpClient _httpClient = new HttpClient() { Timeout = TimeSpan.FromSeconds(60) };

    private static T DeserializeJson<T>(byte[] jsonBytes)
    {
        var ser = new DataContractJsonSerializer(typeof(T));
        using (var ms = new MemoryStream(jsonBytes))
        {
            return (T)ser.ReadObject(ms);
        }
    }

    public static async Task<byte[]> DownloadFileViaTreeAsync(
        string baseUrl,
        string projectId,
        string folderPath,   // 例: "foo/2025-10-22"
        string fileName,     // 例: "index_rpa8.txt"
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        GitLabTreeBlobDownloadResult result = await DownloadFileViaTreeWithDiagnosticsAsync(
            baseUrl,
            projectId,
            folderPath,
            fileName,
            refName,
            privateToken,
            cancellationToken).ConfigureAwait(false);
        return result.Content;
    }

    internal static async Task<GitLabTreeBlobDownloadResult> DownloadFileViaTreeWithDiagnosticsAsync(
        string baseUrl,
        string projectId,
        string folderPath,
        string fileName,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();
        const int perPage = 100;

        GitLabTreeSearchResult search = await GitLabTreePaging.FindBlobAsync(
            fileName,
            perPage,
            page => DownloadTreePageAsync(
                baseUrl,
                projectId,
                folderPath,
                refName,
                page,
                perPage,
                privateToken,
                cancellationToken)).ConfigureAwait(false);

        if (search.Target == null || string.IsNullOrEmpty(search.Target.Id))
        {
            throw new GitLabTreeFileNotFoundException(
                $"File not found in tree. folder={folderPath} file={fileName} ref={refName} pages={search.PagesChecked}",
                search.PagesChecked);
        }

        try
        {
            byte[] content = await DownloadBlobRawAsync(
                baseUrl,
                projectId,
                search.Target.Id,
                privateToken,
                cancellationToken).ConfigureAwait(false);
            return new GitLabTreeBlobDownloadResult
            {
                Content = content,
                PagesChecked = search.PagesChecked,
                FoundPage = search.FoundPage.Value
            };
        }
        catch (GitLabApiException ex)
        {
            throw new GitLabTreeBlobDownloadException(
                ex.Message,
                search.PagesChecked,
                search.FoundPage.Value,
                ex);
        }
    }

    public static async Task<byte[]> TryDownloadFileViaTreeAsync(
        string baseUrl,
        string projectId,
        string folderPath,
        string fileName,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        return await GitLabTreePaging.TryDownloadAsync(
            () => DownloadFileViaTreeAsync(
                baseUrl,
                projectId,
                folderPath,
                fileName,
                refName,
                privateToken,
                cancellationToken)).ConfigureAwait(false);
    }

    internal static async Task<List<GitLabTreeItem>> ListTreeItemsAsync(
        string baseUrl,
        string projectId,
        string folderPath,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        const int perPage = 100;
        int page = 1;

        var allItems = new List<GitLabTreeItem>();

        while (true)
        {
            List<GitLabTreeItem> pageItems = await DownloadTreePageAsync(
                baseUrl,
                projectId,
                folderPath,
                refName,
                page,
                perPage,
                privateToken,
                cancellationToken).ConfigureAwait(false);
            allItems.AddRange(pageItems);

            if (pageItems.Count < perPage)
            {
                break;
            }

            page++;
        }

        return allItems;
    }

    private static async Task<List<GitLabTreeItem>> DownloadTreePageAsync(
        string baseUrl,
        string projectId,
        string folderPath,
        string refName,
        int page,
        int perPage,
        string privateToken,
        CancellationToken cancellationToken)
    {
        string encodedFolder = Uri.EscapeDataString(folderPath ?? string.Empty);
        string encodedRef = Uri.EscapeDataString(refName);
        string treeUrl =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/tree" +
            $"?path={encodedFolder}&ref={encodedRef}&per_page={perPage}&page={page}";

        using (var req = new HttpRequestMessage(HttpMethod.Get, treeUrl))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bodyBytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, treeUrl, bodyBytes);
                }

                return GitLabResponseValidation.ReadTreePage(
                    (int)res.StatusCode,
                    bodyBytes,
                    treeUrl);
            }
        }
    }

    public static async Task<byte[]> DownloadBlobRawAsync(
        string baseUrl,
        string projectId,
        string blobId,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string blobUrl =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/blobs/{Uri.EscapeDataString(blobId)}/raw";

        using (var req = new HttpRequestMessage(HttpMethod.Get, blobUrl))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, blobUrl, bytes);
                }

                return bytes;
            }
        }
    }

    public static async Task<byte[]> DownloadArchiveZipAsync(
        string baseUrl,
        string projectId,
        string refName,
        string archivePath,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/archive.zip?sha={Uri.EscapeDataString(refName)}&path={Uri.EscapeDataString(archivePath ?? string.Empty)}";

        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, url, bytes);
                }

                return bytes;
            }
        }
    }

    public static async Task<byte[]> DownloadFileRawByPathAsync(
        string baseUrl,
        string projectId,
        string filePath,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/files/{Uri.EscapeDataString(filePath)}/raw?ref={Uri.EscapeDataString(refName)}";

        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, url, bytes);
                }

                return bytes;
            }
        }
    }

    public static async Task<byte[]> TryDownloadFileRawByPathAsync(
        string baseUrl,
        string projectId,
        string filePath,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/files/{Uri.EscapeDataString(filePath)}/raw?ref={Uri.EscapeDataString(refName)}";

        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (res.StatusCode == HttpStatusCode.NotFound)
                {
                    return null;
                }

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, url, bytes);
                }

                return bytes;
            }
        }
    }

    internal static async Task<GitLabRepositoryFileInfo> TryGetRepositoryFileInfoAsync(
        string baseUrl,
        string projectId,
        string filePath,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/files/{Uri.EscapeDataString(filePath)}?ref={Uri.EscapeDataString(refName)}";

        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (res.StatusCode == HttpStatusCode.NotFound)
                {
                    return GitLabResponseValidation.ReadRepositoryFileInfo(
                        (int)res.StatusCode,
                        bytes,
                        url);
                }

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, url, bytes);
                }

                return GitLabResponseValidation.ReadRepositoryFileInfo(
                    (int)res.StatusCode,
                    bytes,
                    url);
            }
        }
    }

    public static async Task UpsertTextFileAsync(
        string baseUrl,
        string projectId,
        string filePath,
        string refName,
        string privateToken,
        string content,
        string commitMessage,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        GitLabRepositoryFileInfo existingFile = await TryGetRepositoryFileInfoAsync(
            baseUrl,
            projectId,
            filePath,
            refName,
            privateToken,
            cancellationToken).ConfigureAwait(false);

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/files/{Uri.EscapeDataString(filePath)}";

        var body = new Dictionary<string, object>
        {
            { "branch", refName },
            { "commit_message", commitMessage },
            { "content", content ?? string.Empty },
        };

        HttpMethod method;
        if (existingFile == null)
        {
            method = HttpMethod.Post;
        }
        else
        {
            method = HttpMethod.Put;

            if (!string.IsNullOrWhiteSpace(existingFile.LastCommitId))
            {
                body["last_commit_id"] = existingFile.LastCommitId;
            }
        }

        string jsonBody = JsonSerializer.Serialize(body);
        using (var req = new HttpRequestMessage(method, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);
            req.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
            {
                if (!res.IsSuccessStatusCode)
                {
                    byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    ThrowGitLabApiException(res, url, bytes);
                }
            }
        }
    }

    public static async Task CreateCommitAsync(
        string baseUrl,
        string projectId,
        string branchName,
        string privateToken,
        string commitMessage,
        IEnumerable<object> actions,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        if (actions == null)
        {
            throw new ArgumentNullException(nameof(actions));
        }

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/commits";

        var body = new Dictionary<string, object>
        {
            { "branch", branchName },
            { "commit_message", commitMessage },
            { "actions", actions.ToList() }
        };

        string jsonBody = JsonSerializer.Serialize(body);
        using (var req = new HttpRequestMessage(HttpMethod.Post, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);
            req.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
            {
                if (!res.IsSuccessStatusCode)
                {
                    byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    ThrowGitLabApiException(res, url, bytes);
                }
            }
        }
    }

    public static async Task<string> GetCommitIdAsync(
        string baseUrl,
        string projectId,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/commits?ref_name={Uri.EscapeDataString(refName)}&per_page=1";

        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, url, bytes);
                }

                var commits = DeserializeJson<List<GitLabCommitInfo>>(bytes);
                GitLabCommitInfo commit = commits == null ? null : commits.FirstOrDefault();
                if (commit == null || string.IsNullOrWhiteSpace(commit.Id))
                {
                    throw new InvalidOperationException(
                        "GitLab branch or tag could not be resolved.\n" +
                        "Project ID: " + projectId + "\n" +
                        "Ref: " + refName);
                }

                return commit.Id;
            }
        }
    }

    internal static async Task<string> GetLastCommitIdForPathAsync(
        string baseUrl,
        string projectId,
        string refName,
        string filePath,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/commits" +
            $"?ref_name={Uri.EscapeDataString(refName)}&path={Uri.EscapeDataString(filePath)}&per_page=1";

        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, url, bytes);
                }

                return GitLabResponseValidation.ReadPathLastCommitId(
                    (int)res.StatusCode,
                    bytes,
                    url);
            }
        }
    }

    internal static async Task<GitLabProjectInfo> GetProjectInfoAsync(
        string baseUrl,
        string projectId,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}";

        using (var req = new HttpRequestMessage(HttpMethod.Get, url))
        {
            req.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res = await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (!res.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res, url, bytes);
                }

                return DeserializeJson<GitLabProjectInfo>(bytes);
            }
        }
    }

    private static string SafeUtf8(byte[] bytes)
    {
        try { return Encoding.UTF8.GetString(bytes); } catch { return ""; }
    }

    private static void ThrowGitLabApiException(HttpResponseMessage res, string url, byte[] bodyBytes)
    {
        string bodyText = SafeUtf8(bodyBytes);
        if (res.StatusCode == HttpStatusCode.Unauthorized)
        {
            throw new GitLabApiException(
                "GitLab authentication failed.\n\nThe access token is missing, invalid, or expired.\nPlease check the configured GitLab access token.",
                (int)res.StatusCode,
                url,
                bodyText);
        }

        if (res.StatusCode == HttpStatusCode.Forbidden)
        {
            string forbiddenHint = "The access token is valid but does not have permission to access this repository or resource.";
            if (url.IndexOf("/repository/commits", StringComparison.OrdinalIgnoreCase) >= 0 ||
                url.IndexOf("/repository/files/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                forbiddenHint =
                    "The access token is valid, but it does not have write permission for this repository or branch. " +
                    "For shared-sheet upload, verify that the token has GitLab write access and that the target branch allows commits.";
            }

            throw new GitLabApiException(
                "GitLab access forbidden.\n\n" + forbiddenHint + "\nEndpoint: " + url,
                (int)res.StatusCode,
                url,
                bodyText);
        }

        if (res.StatusCode == HttpStatusCode.NotFound)
        {
            string notFoundHint;
            if (url.IndexOf("/repository/tree", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                notFoundHint = "The specified project, branch (ref), or path may not exist in the repository.";
            }
            else if (url.IndexOf("/repository/blobs/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                notFoundHint = "The specified project or blob ID may not exist, or the GitLab base URL may be incorrect.";
            }
            else if (url.IndexOf("/repository/files/", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                notFoundHint = "The specified project, branch (ref), or file path may not exist in the repository.";
            }
            else if (url.IndexOf("/repository/commits", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                notFoundHint = "The specified project or branch (ref) may not exist in the repository.";
            }
            else
            {
                notFoundHint = "The requested GitLab API endpoint or resource could not be found.";
            }

            throw new GitLabApiException(
                $"GitLab resource not found.\n\n{notFoundHint}\nEndpoint: {url}",
                (int)res.StatusCode,
                url,
                bodyText);
        }

        throw new GitLabApiException(
            $"GitLab API error {(int)res.StatusCode} {res.ReasonPhrase}\nURL: {url}\nBody: {bodyText}",
            (int)res.StatusCode,
            url,
            bodyText);
    }
}
