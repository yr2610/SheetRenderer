using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using Microsoft.ClearScript;
using Microsoft.ClearScript.V8;
using YamlDotNet.Serialization;

public sealed class YamlBridge
{
    private readonly IDeserializer _deserializer;

    private readonly V8ScriptEngine _engine;
    private readonly ScriptObject _objCtor;
    private readonly ScriptObject _arrCtor;

    public YamlBridge(V8ScriptEngine engine)
    {
        _deserializer = new DeserializerBuilder().Build();

        _engine = engine ?? throw new ArgumentNullException("engine");
        _objCtor = _engine.Evaluate("Object") as ScriptObject; // = JS の Object 関数
        _arrCtor = _engine.Evaluate("Array") as ScriptObject; // = JS の Array 関数
    }

    private ScriptObject NewJsObject()
    {
        // asConstructor=true で呼ぶと {} が生成される
        var obj = _objCtor != null ? _objCtor.Invoke(true) as ScriptObject
                                    : _engine.Evaluate("({})") as ScriptObject; // フォールバック
        if (obj == null) throw new InvalidOperationException("Failed to create JS object.");
        return obj;
    }

    private ScriptObject NewJsArray()
    {
        // asConstructor=true で呼ぶと [] が生成される
        var arr = _arrCtor != null ? _arrCtor.Invoke(true) as ScriptObject
                                    : _engine.Evaluate("([])") as ScriptObject; // フォールバック
        if (arr == null) throw new InvalidOperationException("Failed to create JS array.");
        return arr;
    }

    // -------------------------------
    // メイン: .NET構造 → JSネイティブ
    // -------------------------------
    private object ToScriptObject(object value)
    {
        if (value == null) return null;

        var mapObj = value as IDictionary<object, object>;
        if (mapObj != null)
        {
            var jsObj = NewJsObject();
            foreach (var kv in mapObj)
            {
                var key = Convert.ToString(kv.Key) ?? string.Empty;
                jsObj[key] = ToScriptObject(kv.Value);
            }
            return jsObj;
        }

        var mapStr = value as IDictionary<string, object>;
        if (mapStr != null)
        {
            var jsObj = NewJsObject();
            foreach (var kv in mapStr)
            {
                jsObj[kv.Key] = ToScriptObject(kv.Value);
            }
            return jsObj;
        }

        var list = value as IList;
        if (list != null)
        {
            var jsArr = NewJsArray();
            for (int i = 0; i < list.Count; i++)
            {
                jsArr[i] = ToScriptObject(list[i]);
            }
            return jsArr;
        }

        return value; // スカラーはそのまま
    }

    // 動作確認用ダミー
    public object MakeSample()
    {
        var dict = new Dictionary<string, object>
        {
            {"a", 1},
            {"b", new List<object> {10, 20}},
            {"c", true}
        };
        return ToScriptObject(dict);
    }

    // JS: const obj = Yaml.Load(text);
    public object Load(string yamlText)
    {
        if (yamlText == null) throw new ArgumentNullException("yamlText");
        var obj = _deserializer.Deserialize<object>(yamlText);
        return ToScriptObject(obj);
    }

    // JS: const obj = Yaml.LoadFile("C:\\path\\file.yml");
    public object LoadFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentNullException("path");
        var full = Path.GetFullPath(path);
        using (var sr = new StreamReader(full))
        {
            var obj = _deserializer.Deserialize<object>(sr);
            return ToScriptObject(obj);
        }
    }
}
