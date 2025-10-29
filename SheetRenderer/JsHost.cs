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
    private static WScriptPolyfill _wscript;
    private static string _baseDir;

    // ① 初期化（アドイン読み込み時 or 最初に使うときに1回だけ呼ぶ）
    public static void Init(string baseDir)
    {
        if (_engine != null) return; // もう初期化済みなら何もしない

        _baseDir = baseDir ?? throw new ArgumentNullException(nameof(baseDir));

        _engine = new V8ScriptEngine();
        _wscript = new WScriptPolyfill();
        _engine.AddHostObject("WScript", _wscript);

        // ここで C# 側の橋渡しオブジェクトを公開したい場合は追加で AddHostObject する
        // 例: _engine.AddHostObject("HostCrypto", new CryptoBridge());
    }

    // ② スクリプト読み込み（WSFの<script src="...">相当）
    //    ここで読み込まれた関数・変数はグローバルに積み上がって残る
    public static void LoadModule(string path)
    {
        EnsureInit();

        // 相対パス対応
        if (!Path.IsPathRooted(path))
        {
            path = Path.Combine(_baseDir, path);
        }

        if (!File.Exists(path))
            throw new FileNotFoundException("JSモジュールが見つかりません", path);

        var code = File.ReadAllText(path);

        // IIFEで包まず、素のままExecuteするのがポイント
        // → これで各ファイルの関数が同じグローバルにたまっていく
        _engine.Execute(path, code);
    }

    // ③ WScript.Arguments っぽいものをセット
    public static void SetArguments(params string[] args)
    {
        EnsureInit();
        _wscript.SetArguments(args);
    }

    // ④ JS側の関数を呼ぶ
    public static object Call(string functionName, params object[] args)
    {
        EnsureInit();

        // C# から JS のグローバル関数を直接呼ぶ
        // _engine.Script.<名前> は dynamic なのでこうやって呼べる
        return _engine.Script[functionName].Invoke(false, args);
    }

    private static void EnsureInit()
    {
        if (_engine == null)
            throw new InvalidOperationException("JsHost.Init() がまだ呼ばれていません。");
    }
}

// 最低限の WScript 代用品
// WScriptもどき（最小）
public class WScriptPolyfill
{
    private string[] _args = new string[0];

    public void SetArguments(string[] args)
    {
        _args = args ?? new string[0];
    }

    public dynamic Arguments
    {
        get { return new ArgumentsView(_args); }
    }

    public void Echo(object msg)
    {
        // 必要ならロギング先を差し替える
        System.Diagnostics.Debug.WriteLine(msg == null ? "" : msg.ToString());
    }

    public void Quit(int code = 0)
    {
        // WSHだとWScript.Quit(code)でホスト終了だけど、
        // ここでは例外投げることでCall側に「失敗」として伝える案もある
        throw new JsQuitException(code);
    }

    private class ArgumentsView
    {
        private readonly string[] _inner;
        public ArgumentsView(string[] args) { _inner = args; }

        public int length { get { return _inner.Length; } }
        public string this[int i] { get { return _inner[i]; } }
        //public string Item(int i) { return _inner[i]; }
    }
}

public class JsQuitException : Exception
{
    public int ExitCode { get; }
    public JsQuitException(int exitCode) : base("Script requested quit: " + exitCode)
    {
        ExitCode = exitCode;
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
