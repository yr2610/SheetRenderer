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
                    string bodyText = SafeUtf8(bodyBytes);
                    throw new InvalidOperationException(
                        $"GitLab tree failed: {(int)res.StatusCode} {res.ReasonPhrase}\nURL: {treeUrl}\nBody: {bodyText}");
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
                    string bodyText = SafeUtf8(bytes);
                    throw new InvalidOperationException(
                        $"GitLab blob raw failed: {(int)res2.StatusCode} {res2.ReasonPhrase}\nURL: {blobUrl}\nBody: {bodyText}");
                }

                return bytes; // バイナリOK（画像もOK）
            }
        }
    }

    private static string SafeUtf8(byte[] bytes)
    {
        try { return Encoding.UTF8.GetString(bytes); } catch { return ""; }
    }
}
