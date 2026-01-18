using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

internal static class TokenStore
{
    private static readonly object _gate = new object();

    private const string AppFolderName = "SheetRenderer";
    private const string FileName = "tokenstore.dat";

    // Optional entropy. Keep stable, otherwise old tokens cannot be decrypted.
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("SheetRenderer.TokenStore.v1");

    public static string Get(string baseUrl, string projectId)
    {
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            throw new ArgumentException("baseUrl is required.", nameof(baseUrl));
        }

        if (string.IsNullOrWhiteSpace(projectId))
        {
            throw new ArgumentException("projectId is required.", nameof(projectId));
        }

        string key = MakeKey(baseUrl, projectId);

        lock (_gate)
        {
            Dictionary<string, string> map = LoadNoThrow();
            if (!map.TryGetValue(key, out string protectedBase64) || string.IsNullOrWhiteSpace(protectedBase64))
            {
                return null;
            }

            try
            {
                byte[] protectedBytes = Convert.FromBase64String(protectedBase64);
                byte[] plainBytes = ProtectedData.Unprotect(protectedBytes, Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plainBytes);
            }
            catch
            {
                // Corrupted entry or changed environment: drop it to avoid repeated failures.
                map.Remove(key);
                SaveNoThrow(map);
                return null;
            }
        }
    }

    public static bool TryGet(string baseUrl, string projectId, out string token)
    {
        token = Get(baseUrl, projectId);
        return !string.IsNullOrEmpty(token);
    }

    public static void Set(string baseUrl, string projectId, string token)
    {
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            throw new ArgumentException("baseUrl is required.", nameof(baseUrl));
        }

        if (string.IsNullOrWhiteSpace(projectId))
        {
            throw new ArgumentException("projectId is required.", nameof(projectId));
        }

        if (string.IsNullOrEmpty(token))
        {
            throw new ArgumentException("token is required.", nameof(token));
        }

        string key = MakeKey(baseUrl, projectId);

        lock (_gate)
        {
            Dictionary<string, string> map = LoadNoThrow();

            byte[] plainBytes = Encoding.UTF8.GetBytes(token);
            byte[] protectedBytes = ProtectedData.Protect(plainBytes, Entropy, DataProtectionScope.CurrentUser);
            string protectedBase64 = Convert.ToBase64String(protectedBytes);

            map[key] = protectedBase64;
            SaveNoThrow(map);
        }
    }

    public static bool Delete(string baseUrl, string projectId)
    {
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            throw new ArgumentException("baseUrl is required.", nameof(baseUrl));
        }

        if (string.IsNullOrWhiteSpace(projectId))
        {
            throw new ArgumentException("projectId is required.", nameof(projectId));
        }

        string key = MakeKey(baseUrl, projectId);

        lock (_gate)
        {
            Dictionary<string, string> map = LoadNoThrow();
            bool removed = map.Remove(key);
            if (removed)
            {
                SaveNoThrow(map);
            }
            return removed;
        }
    }

    public static void ClearAll()
    {
        lock (_gate)
        {
            SaveNoThrow(new Dictionary<string, string>(StringComparer.Ordinal));
        }
    }

    public static int DeleteByBaseUrl(string baseUrl)
    {
        if (string.IsNullOrWhiteSpace(baseUrl))
        {
            throw new ArgumentException("baseUrl is required.", nameof(baseUrl));
        }

        string normalizedBaseUrl = NormalizeBaseUrl(baseUrl);

        lock (_gate)
        {
            Dictionary<string, string> map = LoadNoThrow();
            List<string> toRemove = new List<string>();

            foreach (var kv in map)
            {
                if (kv.Key.StartsWith(normalizedBaseUrl + "|", StringComparison.Ordinal))
                {
                    toRemove.Add(kv.Key);
                }
            }

            foreach (string k in toRemove)
            {
                map.Remove(k);
            }

            if (toRemove.Count > 0)
            {
                SaveNoThrow(map);
            }

            return toRemove.Count;
        }
    }

    public static TokenKeyInfo[] GetAllTokenKeys()
    {
        lock (_gate)
        {
            Dictionary<string, string> map = LoadNoThrow();
            var list = new List<TokenKeyInfo>();

            foreach (var kv in map)
            {
                TokenKeyInfo info;
                if (TryParseKey(kv.Key, out info))
                {
                    list.Add(info);
                }
            }

            list.Sort((a, b) =>
            {
                int c = string.Compare(a.BaseUrl, b.BaseUrl, StringComparison.OrdinalIgnoreCase);
                if (c != 0) return c;
                return string.Compare(a.ProjectId, b.ProjectId, StringComparison.OrdinalIgnoreCase);
            });

            return list.ToArray();
        }
    }

    private static bool TryParseKey(string key, out TokenKeyInfo info)
    {
        info = null;

        if (string.IsNullOrWhiteSpace(key))
        {
            return false;
        }

        int idx = key.IndexOf('|');
        if (idx <= 0 || idx >= key.Length - 1)
        {
            return false;
        }

        string baseUrl = key.Substring(0, idx).Trim();
        string projectId = key.Substring(idx + 1).Trim();

        if (baseUrl.Length == 0 || projectId.Length == 0)
        {
            return false;
        }

        info = new TokenKeyInfo
        {
            BaseUrl = baseUrl,
            ProjectId = projectId
        };
        return true;
    }

    // ------------------------
    // Internals
    // ------------------------

    private static string MakeKey(string baseUrl, string projectId)
    {
        string normalizedBaseUrl = NormalizeBaseUrl(baseUrl);
        string normalizedProjectId = projectId.Trim();
        return normalizedBaseUrl + "|" + normalizedProjectId;
    }

    private static string NormalizeBaseUrl(string baseUrl)
    {
        string s = baseUrl.Trim();
        while (s.EndsWith("/", StringComparison.Ordinal))
        {
            s = s.Substring(0, s.Length - 1);
        }

        if (Uri.TryCreate(s, UriKind.Absolute, out Uri uri))
        {
            string host = uri.Host.ToLowerInvariant();
            string scheme = uri.Scheme.ToLowerInvariant();

            string portPart = "";
            if (!uri.IsDefaultPort)
            {
                portPart = ":" + uri.Port.ToString();
            }

            string pathPart = uri.AbsolutePath;
            if (string.IsNullOrEmpty(pathPart) || pathPart == "/")
            {
                pathPart = "";
            }
            else
            {
                while (pathPart.EndsWith("/", StringComparison.Ordinal))
                {
                    pathPart = pathPart.Substring(0, pathPart.Length - 1);
                }
            }

            return scheme + "://" + host + portPart + pathPart;
        }

        return s;
    }

    private static string GetStorePath()
    {
        string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string dir = Path.Combine(appData, AppFolderName);
        return Path.Combine(dir, FileName);
    }

    private static Dictionary<string, string> LoadNoThrow()
    {
        try
        {
            string path = GetStorePath();
            if (!File.Exists(path))
            {
                return new Dictionary<string, string>(StringComparer.Ordinal);
            }

            string[] lines = File.ReadAllLines(path, Encoding.UTF8);
            var map = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                int idx = line.IndexOf('=');
                if (idx <= 0)
                {
                    continue;
                }

                string key = line.Substring(0, idx).Trim();
                string val = line.Substring(idx + 1).Trim();

                if (key.Length == 0 || val.Length == 0)
                {
                    continue;
                }

                map[key] = val;
            }

            return map;
        }
        catch
        {
            return new Dictionary<string, string>(StringComparer.Ordinal);
        }
    }

    private static void SaveNoThrow(Dictionary<string, string> map)
    {
        try
        {
            string path = GetStorePath();
            string dir = Path.GetDirectoryName(path);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var sb = new StringBuilder();
            sb.AppendLine("# TokenStore (DPAPI protected). Do not edit by hand unless you know what you're doing.");
            foreach (var kv in map)
            {
                // key=base64
                sb.Append(kv.Key);
                sb.Append('=');
                sb.AppendLine(kv.Value);
            }

            string tmp = path + ".tmp";
            File.WriteAllText(tmp, sb.ToString(), Encoding.UTF8);

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            File.Move(tmp, path);
        }
        catch
        {
            // ignore
        }
    }
}
