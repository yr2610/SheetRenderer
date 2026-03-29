using System.Runtime.Serialization;

[DataContract]
internal sealed class GitLabProjectInfo
{
    [DataMember(Name = "name")]
    public string Name { get; set; }
}
