using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

public static class TimeAssigner
{
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
        // 1) このノードの variables を読む
        var variables = node["variables"] as JObject;

        int? ownTime = variables?["time"]?.Value<int?>();
        int? ownDefaultTime = variables?["default_time"]?.Value<int?>();

        // default_time は親 → 子へ継承
        int? effectiveDefaultTime = ownDefaultTime ?? inheritedDefaultTime;

        // 2) 子ノードを取得
        var children = node["children"] as JArray;

        bool hasChildren = children != null && children.Count > 0;

        // 3) 葉ノードの場合
        if (!hasChildren)
        {
            int? estimated =
                ownTime ??
                effectiveDefaultTime;

            if (estimated.HasValue)
            {
                SetEstimatedTime(node, estimated.Value);
            }

            CleanupVariables(node);
            return;
        }

        // 4) 非leafノード
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

        // 5) このノードに time が指定されている場合、未割当 leaf に配分
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
}
