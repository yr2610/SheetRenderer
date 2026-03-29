using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
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
        EnsureTls12();

        // 1) tree 一覧
        string encodedFolder = Uri.EscapeDataString(folderPath);
        string encodedRef = Uri.EscapeDataString(refName);

        string treeUrl =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/tree?path={encodedFolder}&ref={encodedRef}&per_page=100";

        GitLabTreeItem target = null;

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

                var items = DeserializeJson<List<GitLabTreeItem>>(bodyBytes);
                foreach (var it in items)
                {
                    if (string.Equals(it.Type, "blob", StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(it.Name, fileName, StringComparison.Ordinal))
                    {
                        target = it;
                        break;
                    }
                }
            }
        }

        if (target == null || string.IsNullOrEmpty(target.Id))
        {
            throw new InvalidOperationException($"File not found in tree. folder={folderPath} file={fileName} ref={refName}");
        }

        // 2) blob raw
        string blobUrl =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/blobs/{Uri.EscapeDataString(target.Id)}/raw";

        using (var req2 = new HttpRequestMessage(HttpMethod.Get, blobUrl))
        {
            req2.Headers.Add("PRIVATE-TOKEN", privateToken);

            using (var res2 = await _httpClient.SendAsync(req2, HttpCompletionOption.ResponseContentRead, cancellationToken).ConfigureAwait(false))
            {
                byte[] bytes = await res2.Content.ReadAsByteArrayAsync().ConfigureAwait(false);

                if (!res2.IsSuccessStatusCode)
                {
                    ThrowGitLabApiException(res2, blobUrl, bytes);
                }

                return bytes; // バイナリOK（画像もOK）
            }
        }
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

        string encodedFolder = Uri.EscapeDataString(folderPath ?? string.Empty);
        string encodedRef = Uri.EscapeDataString(refName);

        var allItems = new List<GitLabTreeItem>();

        while (true)
        {
            string treeUrl =
                $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/tree?path={encodedFolder}&ref={encodedRef}&per_page={perPage}&page={page}";

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

                    var pageItems = DeserializeJson<List<GitLabTreeItem>>(bodyBytes) ?? new List<GitLabTreeItem>();
                    allItems.AddRange(pageItems);

                    if (pageItems.Count < perPage)
                    {
                        break;
                    }

                    page++;
                }
            }
        }

        return allItems;
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

    public static async Task<string> GetCommitIdAsync(
        string baseUrl,
        string projectId,
        string refName,
        string privateToken,
        CancellationToken cancellationToken = default(CancellationToken))
    {
        EnsureTls12();

        string url =
            $"{baseUrl.TrimEnd('/')}/api/v4/projects/{Uri.EscapeDataString(projectId)}/repository/commits/{Uri.EscapeDataString(refName)}";

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

                var commit = DeserializeJson<GitLabCommitInfo>(bytes);
                return commit == null ? null : commit.Id;
            }
        }
    }

    private static string SafeUtf8(byte[] bytes)
    {
        try { return Encoding.UTF8.GetString(bytes); } catch { return ""; }
    }

    private static void ThrowGitLabApiException(HttpResponseMessage res, string url, byte[] bodyBytes)
    {
        if (res.StatusCode == HttpStatusCode.Unauthorized)
        {
            throw new InvalidOperationException(
                "GitLab authentication failed.\n\nThe access token is missing, invalid, or expired.\nPlease check the configured GitLab access token.");
        }

        if (res.StatusCode == HttpStatusCode.Forbidden)
        {
            throw new InvalidOperationException(
                "GitLab access forbidden.\n\nThe access token is valid but does not have permission to access this repository or resource.");
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
            else
            {
                notFoundHint = "The requested GitLab API endpoint or resource could not be found.";
            }

            throw new InvalidOperationException(
                $"GitLab resource not found.\n\n{notFoundHint}\nEndpoint: {url}");
        }

        string bodyText = SafeUtf8(bodyBytes);
        throw new InvalidOperationException(
            $"GitLab API error {(int)res.StatusCode} {res.ReasonPhrase}\nURL: {url}\nBody: {bodyText}");
    }
}
