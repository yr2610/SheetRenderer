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

internal sealed class GitLabTreeFileNotFoundException : InvalidOperationException
{
    public GitLabTreeFileNotFoundException(string message, int pagesChecked)
        : base(message)
    {
        PagesChecked = pagesChecked;
    }

    public int PagesChecked { get; private set; }
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
            List<GitLabTreeItem> items = await pageLoader(page).ConfigureAwait(false) ??
                new List<GitLabTreeItem>();

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
}
