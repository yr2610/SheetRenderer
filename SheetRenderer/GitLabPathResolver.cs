using System;
internal static class GitLabPathResolver
{
    public static string ResolveGitLabRelativePath(string baseFileRelativePath, string requestedPath)
    {
        string normalizedBase = NormalizeGitLabRelativePath(baseFileRelativePath);
        string normalizedRequested = NormalizeGitLabRelativePath(requestedPath);

        if (string.IsNullOrEmpty(normalizedRequested))
        {
            throw new ArgumentException("requestedPath is empty.", "requestedPath");
        }

        string baseFolder = GetParentFolder(normalizedBase);
        string combined = string.IsNullOrEmpty(baseFolder)
            ? normalizedRequested
            : baseFolder + "/" + normalizedRequested;

        return NormalizeCombinedPath(combined);
    }

    public static string NormalizeGitLabRelativePath(string path)
    {
        string normalized = (path ?? string.Empty).Replace('\\', '/').Trim();
        normalized = normalized.Trim('/');
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

    private static string NormalizeCombinedPath(string path)
    {
        string[] parts = (path ?? string.Empty).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var stack = new System.Collections.Generic.List<string>();

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
                    throw new InvalidOperationException("Path escapes repository root: " + path);
                }

                stack.RemoveAt(stack.Count - 1);
                continue;
            }

            stack.Add(part);
        }

        return string.Join("/", stack.ToArray());
    }
}
