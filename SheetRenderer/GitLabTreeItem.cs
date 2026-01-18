using System.Runtime.Serialization;

[DataContract]
internal sealed class GitLabTreeItem
{
    [DataMember(Name = "id")]
    public string Id { get; set; } // blob sha

    [DataMember(Name = "name")]
    public string Name { get; set; }

    [DataMember(Name = "type")]
    public string Type { get; set; } // "blob" or "tree"

    [DataMember(Name = "path")]
    public string Path { get; set; }
}
