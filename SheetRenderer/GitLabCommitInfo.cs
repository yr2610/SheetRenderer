using System.Runtime.Serialization;

[DataContract]
internal sealed class GitLabCommitInfo
{
    [DataMember(Name = "id")]
    public string Id { get; set; }
}
