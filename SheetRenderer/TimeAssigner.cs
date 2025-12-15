using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Serialization;

public static class TimeAssigner
{
    // ========= public API =========

    public static void Assign(string jsonFilePath)
    {
        if (string.IsNullOrEmpty(jsonFilePath))
        {
            throw new ArgumentException(nameof(jsonFilePath));
        }

        string json = File.ReadAllText(jsonFilePath);

        dynamic root = JsonConvert.DeserializeObject<ExpandoObject>(
            json,
            new ExpandoObjectConverter()
        );

        Apply(root);

        var settings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented,
            ContractResolver = new DefaultContractResolver
            {
                NamingStrategy = new CamelCaseNamingStrategy()
            }
        };

        string output = JsonConvert.SerializeObject(root, settings);
        File.WriteAllText(jsonFilePath, output);
    }

    public static void Apply(dynamic root)
    {
        foreach (var node in root.children)
        {
            SheetCalcTime(node);
        }
    }

    private static void SheetCalcTime(dynamic node)
    {
        InitializeParent(node, null);

        ForAllNodesRecurse(node, (Action<dynamic>)(n => InitializeTotalTimeNode(n)));
        ForAllNodesRecurse(node, (Action<dynamic>)(n => Pre(n)));
        ForAllNodesRecurse(node, (Action<dynamic>)(n => AffectTime(n)));

        DeletePropertyForAllNodes(node, "affectNodes");
        DeletePropertyForAllNodes(node, "exclusionTime");
        DeletePropertyForAllNodes(node, "parent");

        DeleteVariablesForAllNodes(node, new List<string> { "default_time", "time" });
    }

    private static void InitializeTotalTimeNode(dynamic n)
    {
        if (IsLeaf(n)) return;

        var time = GetTime(n);
        if (time != null)
        {
            n.affectNodes = new List<dynamic>();
            n.exclusionTime = 0;
        }
    }

    private static void Pre(dynamic n)
    {
        if (IsLeaf(n))
        {
            string result = null;
            object resultValue = null;

            if (n.initialValues != null &&
                ((IDictionary<string, object>)n.initialValues)
                    .TryGetValue("result", out resultValue))
            {
                result = resultValue as string;
            }

            if (!string.IsNullOrEmpty(result) && result.StartsWith("-"))
            {
                return;
            }
        }

        int? time = GetTime(n);
        var parent = n.parent;

        while (parent != null)
        {
            int? totalTime = GetTime(parent);
            if (totalTime != null)
            {
                if (time == null)
                {
                    if (IsLeaf(n))
                    {
                        parent.affectNodes.Add(n);
                    }
                }
                else
                {
                    parent.exclusionTime += time.Value;
                }
                return;
            }

            if (IsLeaf(n) && time == null)
            {
                int? defaultTime = GetDefaultTime(parent);
                if (defaultTime != null)
                {
                    SetEstimatedTime(n, defaultTime.Value);
                    time = defaultTime;
                }
            }

            parent = parent.parent;
        }
    }

    private static void AffectTime(dynamic n)
    {
        int? time = GetTime(n);
        if (time == null) return;

        if (IsLeaf(n))
        {
            SetEstimatedTime(n, time.Value);
            return;
        }

        int adjustedTime = Math.Max(0, time.Value - (int)n.exclusionTime);
        int leafTime = adjustedTime / n.affectNodes.Count;
        int remain = adjustedTime % n.affectNodes.Count;

        foreach (var affectNode in n.affectNodes)
        {
            int actualLeafTime = leafTime;
            if (remain-- > 0)
            {
                actualLeafTime++;
            }
            SetEstimatedTime(affectNode, actualLeafTime);
        }
    }

    private static void InitializeParent(dynamic node, dynamic parent)
    {
        node.parent = parent;
        foreach (var child in node.children)
        {
            InitializeParent(child, node);
        }
    }

    private static void DeletePropertyForAllNodes(dynamic node, string name)
    {
        ForAllNodesRecurse(node, (Action<dynamic>)(n =>
        {
            var dict = (IDictionary<string, object>)n;
            if (dict.ContainsKey(name))
            {
                dict.Remove(name);
            }
        }));
    }

    private static void DeleteVariablesForAllNodes(dynamic node, List<string> names)
    {
        ForAllNodesRecurse(node, (Action<dynamic>)(n =>
        {
            if (n.variables == null) return;

            var vars = (IDictionary<string, object>)n.variables;
            foreach (var name in names)
            {
                if (vars.ContainsKey(name))
                {
                    vars.Remove(name);
                }
            }
        }));
    }

    private static void SetEstimatedTime(dynamic node, int time)
    {
        if (node.initialValues == null)
        {
            node.initialValues = new ExpandoObject();
        }
        node.initialValues.estimated_time = time;
    }

    private static int? GetNumber(dynamic node, string name)
    {
        if (node.variables == null) return null;

        object value;
        if (((IDictionary<string, object>)node.variables)
            .TryGetValue(name, out value))
        {
            if (value is string s)
            {
                int n;
                if (int.TryParse(s, out n)) return n;
            }
            return value as int?;
        }

        return null;
    }

    private static int? GetTime(dynamic node)
    {
        return GetNumber(node, "time");
    }

    private static int? GetDefaultTime(dynamic node)
    {
        return GetNumber(node, "default_time");
    }

    private static bool IsLeaf(dynamic node)
    {
        return node.children == null || node.children.Count == 0;
    }

    private static void ForAllNodesRecurse(dynamic node, Action<dynamic> action)
    {
        action(node);
        foreach (var child in node.children)
        {
            ForAllNodesRecurse(child, action);
        }
    }
}
