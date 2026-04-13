using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Windows.Forms;

internal static class AuthorizationHelper
{
    internal static bool ALLOW_ALL_USERS = false;

    private const string PolicyFileName = "sr_policy.dat";
    private const string PolicyPepper = "SheetRendererAuth:v1";

    private sealed class AuthorizedUserEntry
    {
        public string Hash { get; set; }
        public string ExpireDateText { get; set; }
        public DateTime? ExpireDate { get; set; }
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
        AuthorizedUserEntry matchedEntry;
        return TryFindMatchingEntry(currentUserName, out matchedEntry) &&
            (!matchedEntry.ExpireDate.HasValue || DateTime.Today <= matchedEntry.ExpireDate.Value.Date);
    }

    internal static bool EnsureAuthorizedUser()
    {
        if (IsAuthorizedUser())
        {
            return true;
        }

        string currentUserName = GetCurrentUserName();
        AuthorizedUserEntry matchedEntry;
        bool hasMatchedEntry = TryFindMatchingEntry(currentUserName, out matchedEntry);
        bool isExpired = hasMatchedEntry &&
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

    private static bool TryFindMatchingEntry(string currentUserName, out AuthorizedUserEntry matchedEntry)
    {
        matchedEntry = null;

        foreach (AuthorizedUserEntry entry in LoadAuthorizedUsers())
        {
            string expectedHash = ComputeUserHash(currentUserName, entry.ExpireDateText);
            if (string.Equals(entry.Hash, expectedHash, StringComparison.OrdinalIgnoreCase))
            {
                matchedEntry = entry;
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<AuthorizedUserEntry> LoadAuthorizedUsers()
    {
        try
        {
            string policyPath = GetPolicyPath();
            if (!File.Exists(policyPath))
            {
                return Array.Empty<AuthorizedUserEntry>();
            }

            string encoded = File.ReadAllText(policyPath, Encoding.UTF8).Trim();
            if (string.IsNullOrWhiteSpace(encoded))
            {
                return Array.Empty<AuthorizedUserEntry>();
            }

            byte[] compressedBytes = Convert.FromBase64String(encoded);
            string payloadText = DecompressPolicyPayload(compressedBytes);

            var entries = new List<AuthorizedUserEntry>();
            string[] lines = payloadText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            foreach (string rawLine in lines)
            {
                string line = (rawLine ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                string[] parts = line.Split(new[] { '|' }, 2);
                string hash = parts.Length > 0 ? parts[0].Trim() : string.Empty;
                string expireDateText = parts.Length > 1 ? parts[1].Trim() : string.Empty;

                if (string.IsNullOrWhiteSpace(hash))
                {
                    continue;
                }

                DateTime? expireDate = null;
                if (!string.IsNullOrWhiteSpace(expireDateText))
                {
                    DateTime parsedDate;
                    if (!DateTime.TryParseExact(
                        expireDateText,
                        "yyyy-MM-dd",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out parsedDate))
                    {
                        continue;
                    }

                    expireDate = parsedDate.Date;
                }

                entries.Add(new AuthorizedUserEntry
                {
                    Hash = hash,
                    ExpireDateText = expireDateText,
                    ExpireDate = expireDate
                });
            }

            return entries;
        }
        catch
        {
            return Array.Empty<AuthorizedUserEntry>();
        }
    }

    private static string GetPolicyPath()
    {
        return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, PolicyFileName);
    }

    private static string DecompressPolicyPayload(byte[] compressedBytes)
    {
        using (var input = new MemoryStream(compressedBytes))
        using (var gzip = new GZipStream(input, CompressionMode.Decompress))
        using (var output = new MemoryStream())
        {
            gzip.CopyTo(output);
            return Encoding.UTF8.GetString(output.ToArray());
        }
    }

    private static string ComputeUserHash(string userName, string expireDateText)
    {
        string normalizedExpireDate = expireDateText ?? string.Empty;
        string payload = PolicyPepper + "|" + userName + "|" + normalizedExpireDate;
        byte[] bytes = Encoding.UTF8.GetBytes(payload);

        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(bytes);
            var builder = new StringBuilder(hashBytes.Length * 2);
            foreach (byte b in hashBytes)
            {
                builder.Append(b.ToString("x2"));
            }

            return builder.ToString();
        }
    }
}
