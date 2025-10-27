using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Microsoft.ClearScript;
using Microsoft.ClearScript.V8;

public static class JsHost
{
    private static V8ScriptEngine _engine;

    public static void Init(string scriptsRoot)
    {
        // V8 エンジン生成（Isolate 名でキャッシュしやすく）
        _engine = new V8ScriptEngine(V8ScriptEngineFlags.None);

        // 1) 便利ユーティリティ（.NET 側の安全ラッパ）を公開
        _engine.AddHostObject("host", new HostFunctions());
        _engine.AddHostType("Console", typeof(Console));

        // 2) 最小 WSH 互換ポリフィルを注入
        _engine.Script.WScript = new WScriptPolyfill();
        _engine.Script.FS = new FsPolyfill(scriptsRoot);

        // 3) 必要ならグローバル定数
        _engine.Script.__SCRIPTS_ROOT__ = scriptsRoot;
    }

    public static object RunFile(string jsPath, params string[] args)
    {
        if (!Path.IsPathRooted(jsPath))
        {
            throw new ArgumentException("Absolute path required.", nameof(jsPath));
        }

        // 引数を WScript.Arguments 風に供給
        (_engine.Script.WScript as WScriptPolyfill)?.SetArguments(args);

        string code = File.ReadAllText(jsPath);
        return _engine.Evaluate($"(function() {{ {code}\n }})()");
    }
}

public class WScriptPolyfill
{
    private string[] _args = Array.Empty<string>();

    public void Echo(string s)
    {
        Console.WriteLine(s);
    }

    public void Quit(int code = 0)
    {
        throw new ScriptEngineException($"WScript.Quit({code}) called");
    }

    public dynamic Arguments
    {
        get { return _args; }
    }

    public void SetArguments(string[] args)
    {
        _args = args ?? Array.Empty<string>();
    }
}

public class FsPolyfill
{
    private readonly string _root;

    public FsPolyfill(string root)
    {
        _root = root;
    }

    public string ReadAllText(string path)
    {
        string full = Resolve(path);
        return File.ReadAllText(full);
    }

    public void WriteAllText(string path, string content)
    {
        string full = Resolve(path);
        Directory.CreateDirectory(Path.GetDirectoryName(full));
        File.WriteAllText(full, content);
    }

    public bool Exists(string path)
    {
        string full = Resolve(path);
        return File.Exists(full) || Directory.Exists(full);
    }

    public string MapPath(string path)
    {
        return Resolve(path);
    }

    private string Resolve(string path)
    {
        if (Path.IsPathRooted(path))
        {
            return path;
        }
        return Path.GetFullPath(Path.Combine(_root, path));
    }
}
