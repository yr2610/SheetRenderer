using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

public static class TimeAssigner
{
    public static void Assign(string jsonFilePath)
    {
        if (string.IsNullOrWhiteSpace(jsonFilePath))
        {
            throw new ArgumentException("jsonFilePath is required.", nameof(jsonFilePath));
        }

        if (!File.Exists(jsonFilePath))
        {
            throw new FileNotFoundException("JSON ファイルが存在しません。", jsonFilePath);
        }

        JObject root;
        using (var reader = new StreamReader(jsonFilePath, Encoding.UTF8))
        using (var jsonReader = new JsonTextReader(reader))
        {
            root = JObject.Load(jsonReader);
        }

        Apply(root);

        using (var writer = new StreamWriter(jsonFilePath, false, new UTF8Encoding(false)))
        using (var jsonWriter = new JsonTextWriter(writer) { Formatting = Formatting.Indented })
        {
            root.WriteTo(jsonWriter);
        }
    }

    public static void Apply(JObject root)
    {
        if (root == null)
        {
            throw new ArgumentNullException(nameof(root));
        }

        AssignRecursive(
            node: root,
            inheritedDefaultTime: null
        );
    }

    private static void AssignRecursive(
        JObject node,
        int? inheritedDefaultTime
    )
    {
        var variables = node["variables"] as JObject;

        int? ownTime = variables?["time"]?.Value<int?>();
        int? ownDefaultTime = variables?["default_time"]?.Value<int?>();

        int? effectiveDefaultTime = ownDefaultTime ?? inheritedDefaultTime;

        var children = node["children"] as JArray;

        bool hasChildren = children != null && children.Count > 0;

        if (!hasChildren)
        {
            int? estimated =
                ownTime ??
                effectiveDefaultTime;

            if (estimated.HasValue)
            {
                SetEstimatedTime(node, estimated.Value);
            }
            else
            {
                WarnNoEstimatedTime(node);
            }

            CleanupVariables(node);
            return;
        }

        int totalAssigned = 0;
        var leafNodes = new List<JObject>();

        foreach (var child in children)
        {
            if (child is JObject childObj)
            {
                AssignRecursive(childObj, effectiveDefaultTime);

                int? childTime = GetEstimatedTime(childObj);
                if (childTime.HasValue)
                {
                    totalAssigned += childTime.Value;
                }
                else
                {
                    leafNodes.Add(childObj);
                }
            }
        }

        if (ownTime.HasValue && leafNodes.Count > 0)
        {
            int remain = ownTime.Value - totalAssigned;
            if (remain > 0)
            {
                int perNode = remain / leafNodes.Count;

                foreach (var leaf in leafNodes)
                {
                    SetEstimatedTime(leaf, perNode);
                }
            }
            else if (remain < 0)
            {
                FileLogger.Warn($"子ノードの時間合計が上限を超えています (id={GetNodeId(node)}, remain={remain}).");
            }
        }

        CleanupVariables(node);
    }

    private static void SetEstimatedTime(JObject node, int time)
    {
        var initialValues = node["initialValues"] as JObject;
        if (initialValues == null)
        {
            initialValues = new JObject();
            node["initialValues"] = initialValues;
        }

        int? existing = initialValues["estimated_time"]?.Value<int?>();
        if (existing.HasValue && existing.Value != time)
        {
            FileLogger.Warn(
                $"既存の estimated_time を上書きしました (id={GetNodeId(node)}, before={existing.Value}, after={time}).");
        }

        initialValues["estimated_time"] = time;
    }

    private static int? GetEstimatedTime(JObject node)
    {
        return node["initialValues"]?["estimated_time"]?.Value<int?>();
    }

    private static void CleanupVariables(JObject node)
    {
        var variables = node["variables"] as JObject;
        if (variables == null) return;

        variables.Remove("time");
        variables.Remove("default_time");

        if (!variables.HasValues)
        {
            node.Remove("variables");
        }
    }

    private static void WarnNoEstimatedTime(JObject node)
    {
        FileLogger.Warn($"推定時間を割り当てできませんでした (id={GetNodeId(node)}).");
    }

    private static string GetNodeId(JObject node)
    {
        var id = node["id"];
        if (id == null)
        {
            return "(no id)";
        }

        return id.ToString();
    }
}
