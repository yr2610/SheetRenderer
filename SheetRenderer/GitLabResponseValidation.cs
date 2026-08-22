using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

internal sealed class GitLabResponseFormatException : InvalidOperationException
{
    public GitLabResponseFormatException(
        int statusCode,
        string url,
        string expectedFormat,
        byte[] responseBytes,
        Exception innerException = null)
        : base(
            "GitLab returned an invalid successful response. Status: " + statusCode +
            ", URL: " + url +
            ", Expected: " + expectedFormat,
            innerException)
    {
        StatusCode = statusCode;
        Url = url;
        ExpectedFormat = expectedFormat;
        ResponseBody = GetLimitedResponseBody(responseBytes);
    }

    public int StatusCode { get; private set; }

    public string Url { get; private set; }

    public string ExpectedFormat { get; private set; }

    public string ResponseBody { get; private set; }

    private static string GetLimitedResponseBody(byte[] responseBytes)
    {
        const int maxLength = 2000;
        string body = responseBytes == null
            ? string.Empty
            : Encoding.UTF8.GetString(responseBytes);
        return body.Length <= maxLength
            ? body
            : body.Substring(0, maxLength) + "...";
    }
}

internal static class GitLabResponseValidation
{
    public static GitLabRepositoryFileInfo ReadRepositoryFileInfo(
        int statusCode,
        byte[] responseBytes,
        string url)
    {
        if (statusCode == 404)
        {
            return null;
        }

        return DeserializeRequired<GitLabRepositoryFileInfo>(
            statusCode,
            responseBytes,
            url,
            "a non-null repository file metadata object");
    }

    public static string ReadPathLastCommitId(
        int statusCode,
        byte[] responseBytes,
        string url)
    {
        List<GitLabCommitInfo> commits = DeserializeRequired<List<GitLabCommitInfo>>(
            statusCode,
            responseBytes,
            url,
            "a non-null path commit array");
        if (commits.Count == 0)
        {
            return null;
        }

        for (int index = 0; index < commits.Count; index++)
        {
            GitLabCommitInfo commit = commits[index];
            if (commit == null || string.IsNullOrWhiteSpace(commit.Id))
            {
                throw InvalidResponse(
                    statusCode,
                    responseBytes,
                    url,
                    "path commit item " + index + " with a non-empty id");
            }
        }

        return commits[0].Id;
    }

    public static List<GitLabTreeItem> ReadTreePage(
        int statusCode,
        byte[] responseBytes,
        string url)
    {
        List<GitLabTreeItem> items = DeserializeRequired<List<GitLabTreeItem>>(
            statusCode,
            responseBytes,
            url,
            "a non-null repository tree array");
        for (int index = 0; index < items.Count; index++)
        {
            GitLabTreeItem item = items[index];
            if (item == null ||
                string.IsNullOrWhiteSpace(item.Id) ||
                string.IsNullOrWhiteSpace(item.Name) ||
                string.IsNullOrWhiteSpace(item.Type))
            {
                throw InvalidResponse(
                    statusCode,
                    responseBytes,
                    url,
                    "repository tree item " + index + " with non-empty id, name, and type");
            }
        }

        return items;
    }

    private static T DeserializeRequired<T>(
        int statusCode,
        byte[] responseBytes,
        string url,
        string expectedFormat)
        where T : class
    {
        T result;
        try
        {
            var serializer = new DataContractJsonSerializer(typeof(T));
            using (var stream = new MemoryStream(responseBytes ?? new byte[0]))
            {
                result = (T)serializer.ReadObject(stream);
            }
        }
        catch (Exception ex)
        {
            throw new GitLabResponseFormatException(
                statusCode,
                url,
                expectedFormat,
                responseBytes,
                ex);
        }

        if (result == null)
        {
            throw InvalidResponse(statusCode, responseBytes, url, expectedFormat);
        }

        return result;
    }

    private static GitLabResponseFormatException InvalidResponse(
        int statusCode,
        byte[] responseBytes,
        string url,
        string expectedFormat)
    {
        return new GitLabResponseFormatException(
            statusCode,
            url,
            expectedFormat,
            responseBytes);
    }
}
