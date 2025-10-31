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

    public string GetFileName(string path)
    {
        return Path.GetFileName(path);
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
}
