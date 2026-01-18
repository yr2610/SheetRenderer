using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

internal static class GitLabLastInputStore
{
    private static string GetPath()
    {
        string dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "SheetRenderer");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "gitlab_last.json");
    }

    public static GitLabLastInput Load()
    {
        string path = GetPath();
        if (!File.Exists(path))
        {
            return new GitLabLastInput();
        }

        try
        {
            byte[] bytes = File.ReadAllBytes(path);
            var ser = new DataContractJsonSerializer(typeof(GitLabLastInput));
            using (var ms = new MemoryStream(bytes))
            {
                return (GitLabLastInput)ser.ReadObject(ms) ?? new GitLabLastInput();
            }
        }
        catch
        {
            return new GitLabLastInput();
        }
    }

    public static void Save(GitLabLastInput data, bool clearFilePath)
    {
        if (data == null) data = new GitLabLastInput();

        // FilePath だけ毎回空欄にしたい場合
        if (clearFilePath)
        {
            data.FilePath = "";
        }

        string path = GetPath();
        var ser = new DataContractJsonSerializer(typeof(GitLabLastInput));
        using (var ms = new MemoryStream())
        {
            ser.WriteObject(ms, data);
            File.WriteAllBytes(path, ms.ToArray());
        }
    }
}
