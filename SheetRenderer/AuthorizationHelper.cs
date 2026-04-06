using System;
using System.Collections.Generic;
using System.Security.Principal;
using System.Windows.Forms;

internal static class AuthorizationHelper
{
    internal static bool ALLOW_ALL_USERS = false;

    private static readonly HashSet<string> AllowedUsers = new HashSet<string>(StringComparer.Ordinal)
    {
        // 例: @"DOMAIN\UserName",
        @"LAPTOP-9S8RJR29\shinn",
    };

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
        return AllowedUsers.Contains(currentUserName);
    }

    internal static bool EnsureAuthorizedUser()
    {
        if (IsAuthorizedUser())
        {
            return true;
        }

        string currentUserName = GetCurrentUserName();
        MessageBox.Show(
            "このユーザーには本アドインの使用権限がありません。" + Environment.NewLine +
            "Current user: " + currentUserName,
            "認可エラー",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);

        return false;
    }
}
