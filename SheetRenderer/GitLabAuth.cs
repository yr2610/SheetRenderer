using System;
using System.Windows.Forms;

internal static class GitLabAuth
{
    public static string GetOrPromptToken(string baseUrl, string projectId)
    {
        string token = TokenStore.Get(baseUrl, projectId);
        if (!string.IsNullOrEmpty(token))
        {
            return token;
        }

        IWin32Window owner = GetExcelOwnerWindow();

        string input;
        bool remember;
        var result = TokenInputDialog.ShowDialog(owner, baseUrl, projectId, out input, out remember);

        if (result != DialogResult.OK)
        {
            return null;
        }

        try
        {
            GitLabProjectInfo projectInfo = GitLabClient.GetProjectInfoAsync(baseUrl, projectId, input)
                .GetAwaiter()
                .GetResult();

            if (projectInfo == null)
            {
                throw new InvalidOperationException("GitLab project could not be resolved.");
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                owner,
                "GitLab への接続確認に失敗しました。\n" + ex.Message,
                "GitLab Token",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return null;
        }

        if (remember)
        {
            TokenStore.Set(baseUrl, projectId, input);
        }

        return input;
    }

    public static bool DeleteToken(string baseUrl, string projectId)
    {
        return TokenStore.Delete(baseUrl, projectId);
    }

    private static IWin32Window GetExcelOwnerWindow()
    {
        // owner を渡すと、Excelの前面でモーダル表示になりやすい
        // 取れない環境もあるので、その場合は null でOK
        try
        {
            IntPtr hwnd = ExcelDna.Integration.ExcelDnaUtil.WindowHandle;
            if (hwnd != IntPtr.Zero)
            {
                return new WindowWrapper(hwnd);
            }
        }
        catch
        {
        }

        return null;
    }

    private sealed class WindowWrapper : IWin32Window
    {
        private readonly IntPtr _hwnd;

        public WindowWrapper(IntPtr hwnd)
        {
            _hwnd = hwnd;
        }

        public IntPtr Handle
        {
            get
            {
                return _hwnd;
            }
        }
    }
}
