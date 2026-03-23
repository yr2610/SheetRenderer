using System;
using System.Collections.Generic;

internal static class GitLabPathResolver
{
    public static string ResolveGitLabRelativePath(string baseFileRelativePath, string requestedPath)
    {
        string normalizedBase = CanonicalizeGitLabRelativePath(baseFileRelativePath, "baseFileRelativePath");
        string normalizedRequested = CanonicalizeGitLabRelativePath(requestedPath, "requestedPath");

        if (string.IsNullOrEmpty(normalizedRequested))
        {
            throw new ArgumentException("requestedPath is empty.", "requestedPath");
        }

        string baseFolder = GetParentFolder(normalizedBase);
        string combined = string.IsNullOrEmpty(baseFolder)
            ? normalizedRequested
            : baseFolder + "/" + normalizedRequested;

        return CanonicalizeGitLabRelativePath(combined, "requestedPath", requestedPath, baseFileRelativePath);
    }

    public static string CanonicalizeGitLabRelativePath(string path)
    {
        return CanonicalizeGitLabRelativePath(path, null, null, null);
    }

    public static string CanonicalizeGitLabRelativePath(string path, string parameterName)
    {
        return CanonicalizeGitLabRelativePath(path, parameterName, null, null);
    }

    public static string CanonicalizeGitLabRelativePath(string path, string parameterName, string originalPath, string basePath)
    {
        string rawPath = path ?? string.Empty;
        string normalized = rawPath.Replace('\\', '/').Trim();
        string[] parts = normalized.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var stack = new List<string>();

        for (int i = 0; i < parts.Length; i++)
        {
            string part = parts[i];
            if (part == ".")
            {
                continue;
            }

            if (part == "..")
            {
                if (stack.Count == 0)
                {
                    throw CreatePathEscapeException(rawPath, parameterName, originalPath, basePath);
                }

                stack.RemoveAt(stack.Count - 1);
                continue;
            }

            stack.Add(part);
        }

        return string.Join("/", stack.ToArray());
    }

    public static string NormalizeGitLabRelativePath(string path)
    {
        return CanonicalizeGitLabRelativePath(path);
    }

    public static string NormalizeGitLabFilePathStrict(string path)
    {
        string normalized = CanonicalizeGitLabRelativePath(path, "path");
        if (string.IsNullOrEmpty(normalized))
        {
            throw new ArgumentException("path is empty.", "path");
        }

        return normalized;
    }

    private static string GetParentFolder(string path)
    {
        int idx = path.LastIndexOf('/');
        if (idx < 0)
        {
            return string.Empty;
        }

        return path.Substring(0, idx);
    }

    private static InvalidOperationException CreatePathEscapeException(
        string path,
        string parameterName,
        string originalPath,
        string basePath)
    {
        string message = "Path escapes repository root. " +
            "path='" + (path ?? string.Empty) + "'";

        if (!string.IsNullOrEmpty(parameterName))
        {
            message += ", parameter='" + parameterName + "'";
        }

        if (!string.IsNullOrEmpty(originalPath))
        {
            message += ", originalPath='" + originalPath + "'";
        }

        if (!string.IsNullOrEmpty(basePath))
        {
            message += ", basePath='" + basePath + "'";
        }

        return new InvalidOperationException(message + ".");
    }
}
