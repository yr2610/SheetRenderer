// FileLogger.cs
using System;
using System.Diagnostics;
using System.IO;
using System.Text;

public static class FileLogger
{
    private static string _logPath;

    // 例: input = C:\work\data\foo.txt
    // 既定: C:\work\data\foo.parse.log  を作成
    public static void InitializeForInput(string inputFilePath, bool timestamped = false)
    {
        if (string.IsNullOrWhiteSpace(inputFilePath) || !File.Exists(inputFilePath))
        {
            throw new FileNotFoundException("入力ファイルが存在しません。", inputFilePath);
        }

        string dir = Path.GetDirectoryName(inputFilePath);
        string baseName = Path.GetFileNameWithoutExtension(inputFilePath);

        if (timestamped)
        {
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            _logPath = Path.Combine(dir, $"{baseName}.{stamp}.parse.log");
        }
        else
        {
            _logPath = Path.Combine(dir, $"{baseName}.parse.log");
        }

        // 先頭にヘッダを一度だけ
        WriteRaw($"===== Parse Log for \"{Path.GetFileName(inputFilePath)}\" ({DateTime.Now:yyyy-MM-dd HH:mm:ss}) =====");
    }

    public static void Info(string msg) { Write("INFO", msg); }
    public static void Warn(string msg) { Write("WARN", msg); }
    public static void Error(string msg) { Write("ERROR", msg); }

    public static void OpenLog()
    {
        EnsureInitialized();
        if (!File.Exists(_logPath))
        {
            File.WriteAllText(_logPath, "");
        }
        Process.Start(new ProcessStartInfo { FileName = _logPath, UseShellExecute = true });
    }

    private static void Write(string level, string msg)
    {
        EnsureInitialized();
        string line = $"{DateTime.Now:HH:mm:ss.fff} [{level}] {msg}";
        File.AppendAllText(_logPath, line + Environment.NewLine, Encoding.UTF8);
    }

    private static void WriteRaw(string line)
    {
        EnsureInitialized(allowUninitialized: true);
        File.AppendAllText(_logPath, line + Environment.NewLine, Encoding.UTF8);
    }

    private static void EnsureInitialized(bool allowUninitialized = false)
    {
        if (_logPath != null) return;
        if (!allowUninitialized) throw new InvalidOperationException("FileLogger.InitializeForInput を先に呼んでください。");
        // allowUninitialized=true の場合はデフォルト場所を仮設定（まず無い想定）
        string fallback = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "parse.log");
        _logPath = fallback;
    }
}
