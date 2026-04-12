using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Principal;
using System.Windows.Forms;

internal static class AuthorizationHelper
{
    internal static bool ALLOW_ALL_USERS = false;

    private sealed class AuthorizedUserEntry
    {
        public string UserName { get; set; }
        public DateTime? ExpireDate { get; set; }
    }

    private static readonly IReadOnlyList<AuthorizedUserEntry> AllowedUsers = new[]
    {
        // CreateAuthorizedUser(@"DOMAIN\UserName"),
        // CreateAuthorizedUser(@"DOMAIN\UserName", "2026-12-31"),
    };

    private static AuthorizedUserEntry CreateAuthorizedUser(string userName, string expireDateText = null)
    {
        if (string.IsNullOrWhiteSpace(userName))
        {
            throw new ArgumentException("userName is required.", nameof(userName));
        }

        DateTime? expireDate = null;
        if (!string.IsNullOrWhiteSpace(expireDateText))
        {
            expireDate = DateTime.ParseExact(
                expireDateText,
                "yyyy-MM-dd",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None);
        }

        return new AuthorizedUserEntry
        {
            UserName = userName,
            ExpireDate = expireDate
        };
    }

    internal static string GetCurrentUserName()
    {
        try
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            if (identity != null && !string.IsNullOrWhiteSpace(identity.Name))
            {
                return identity.Name;
            }
        }
        catch
        {
        }

        return Environment.UserName ?? string.Empty;
    }

    internal static bool IsAuthorizedUser()
    {
        if (ALLOW_ALL_USERS)
        {
            return true;
        }

        string currentUserName = GetCurrentUserName();
        DateTime today = DateTime.Today;

        return AllowedUsers.Any(entry =>
            string.Equals(entry.UserName, currentUserName, StringComparison.OrdinalIgnoreCase) &&
            (!entry.ExpireDate.HasValue || today <= entry.ExpireDate.Value.Date));
    }

    internal static bool EnsureAuthorizedUser()
    {
        if (IsAuthorizedUser())
        {
            return true;
        }

        string currentUserName = GetCurrentUserName();
        AuthorizedUserEntry matchedEntry = AllowedUsers.FirstOrDefault(entry =>
            string.Equals(entry.UserName, currentUserName, StringComparison.OrdinalIgnoreCase));
        bool isExpired = matchedEntry != null &&
            matchedEntry.ExpireDate.HasValue &&
            DateTime.Today > matchedEntry.ExpireDate.Value.Date;

        MessageBox.Show(
            (isExpired
                ? "このユーザーの利用期限は終了しています。"
                : "このユーザーには本アドインの利用許可がありません。") + Environment.NewLine +
            "Current user: " + currentUserName,
            "認証エラー",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);

        return false;
    }
}
