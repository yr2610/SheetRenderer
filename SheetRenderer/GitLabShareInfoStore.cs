using System;
using System.IO;
using System.Runtime.Serialization.Json;

internal static class GitLabShareInfoStore
{
    private static string GetPath()
    {
        string dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SheetRenderer");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "gitlab_share_last.json");
    }

    public static GitLabShareInfo Load()
    {
        string path = GetPath();
        if (!File.Exists(path))
        {
            return new GitLabShareInfo();
        }

        try
        {
            byte[] bytes = File.ReadAllBytes(path);
            var ser = new DataContractJsonSerializer(typeof(GitLabShareInfo));
            using (var ms = new MemoryStream(bytes))
            {
                return (GitLabShareInfo)ser.ReadObject(ms) ?? new GitLabShareInfo();
            }
        }
        catch
        {
            return new GitLabShareInfo();
        }
    }

    public static void Save(GitLabShareInfo data)
    {
        if (data == null)
        {
            data = new GitLabShareInfo();
        }

        string path = GetPath();
        var ser = new DataContractJsonSerializer(typeof(GitLabShareInfo));
        using (var ms = new MemoryStream())
        {
            ser.WriteObject(ms, data);
            File.WriteAllBytes(path, ms.ToArray());
        }
    }
}
