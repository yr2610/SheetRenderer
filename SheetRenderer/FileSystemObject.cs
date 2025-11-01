using System;
using System.IO;

public class FileSystemObject
{
    public bool FileExists(string path)
    {
        return File.Exists(path);
    }

    public bool FolderExists(string path)
    {
        return Directory.Exists(path);
    }

    public string GetParentFolderName(string path)
    {
        return Path.GetDirectoryName(path);
    }

    // FSO: GetFileName に相当（＝拡張子付きのファイル名）
    public string GetFileName(string path)
    {
        return Path.GetFileName(path);
    }

    // FSO: GetExtensionName("a/b.txt") -> "txt"（先頭ドットなし／拡張子なしなら空文字）
    public string GetExtensionName(string path)
    {
        var ext = Path.GetExtension(path);
        return string.IsNullOrEmpty(ext) ? "" : (ext[0] == '.' ? ext.Substring(1) : ext);
    }

    // FSO: GetBaseName("a/b.txt") -> "b"（拡張子抜きのファイル名）
    public string GetBaseName(string path)
    {
        return Path.GetFileNameWithoutExtension(path);
    }

    public string GetAbsolutePathName(string path)
    {
        return Path.GetFullPath(path);
    }

    public void CreateFolder(string path)
    {
        Directory.CreateDirectory(path);
    }

    public void DeleteFile(string path)
    {
        File.Delete(path);
    }

    public void DeleteFolder(string path)
    {
        Directory.Delete(path, recursive: true);
    }

    public TextFile OpenTextFile(string path, int mode = 1, bool create = false)
    {
        // mode: 1=ForReading, 2=ForWriting, 8=ForAppending
        if (mode == 1) return new TextFile(File.OpenText(path));
        if (mode == 2)
        {
            if (create || File.Exists(path))
                return new TextFile(new StreamWriter(path, false));
            throw new FileNotFoundException("File not found", path);
        }
        if (mode == 8)
            return new TextFile(new StreamWriter(path, true));

        throw new ArgumentException($"Unknown mode {mode}");
    }

    public class TextFile : IDisposable
    {
        private StreamReader _reader;
        private StreamWriter _writer;

        public TextFile(StreamReader reader) { _reader = reader; }
        public TextFile(StreamWriter writer) { _writer = writer; }

        public string ReadAll()
        {
            if (_reader == null) throw new InvalidOperationException("Not open for reading");
            return _reader.ReadToEnd();
        }

        public void Write(string text)
        {
            if (_writer == null) throw new InvalidOperationException("Not open for writing");
            _writer.Write(text);
        }

        public void WriteLine(string text)
        {
            if (_writer == null) throw new InvalidOperationException("Not open for writing");
            _writer.WriteLine(text);
        }

        public void Close()
        {
            Dispose();
        }

        public void Dispose()
        {
            _reader?.Dispose();
            _writer?.Dispose();
        }
    }

    // FSO: BuildPath("a/b", "c.txt") -> "a/b/c.txt"（区切り気にせず結合）
    public string BuildPath(string path, string name)
    {
        return Path.Combine(path ?? "", name ?? "");
    }

    // FSO: CopyFile(source, destination [, overwrite=false])
    // destination がフォルダなら中へコピー、ファイルならその名前でコピー
    public void CopyFile(string source, string destination, bool overwrite = false)
    {
        if (string.IsNullOrEmpty(source)) throw new ArgumentNullException(nameof(source));
        if (string.IsNullOrEmpty(destination)) throw new ArgumentNullException(nameof(destination));
        if (!File.Exists(source)) throw new FileNotFoundException("Source file not found", source);

        string destFile;
        if (Directory.Exists(destination))
        {
            destFile = Path.Combine(destination, Path.GetFileName(source));
        }
        else
        {
            // フォルダっぽい終端（…\ or …/）ならフォルダとして扱う
            if (destination.EndsWith(Path.DirectorySeparatorChar.ToString()) ||
                destination.EndsWith(Path.AltDirectorySeparatorChar.ToString()))
            {
                Directory.CreateDirectory(destination);
                destFile = Path.Combine(destination, Path.GetFileName(source));
            }
            else
            {
                destFile = destination;
                // 親フォルダが無ければ作成（FSOはエラーだが、運用的にこちらが便利）
                var parent = Path.GetDirectoryName(destFile);
                if (!string.IsNullOrEmpty(parent) && !Directory.Exists(parent))
                    Directory.CreateDirectory(parent);
            }
        }

        File.Copy(source, destFile, overwrite);
    }
}
