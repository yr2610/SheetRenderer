using System.Runtime.Serialization;

[DataContract]
internal sealed class GitLabRepositoryFileInfo
{
    [DataMember(Name = "file_path")]
    public string FilePath { get; set; }

    [DataMember(Name = "last_commit_id")]
    public string LastCommitId { get; set; }

    [DataMember(Name = "content_sha256")]
    public string ContentSha256 { get; set; }

    [DataMember(Name = "content")]
    public string Content { get; set; }

    [DataMember(Name = "encoding")]
    public string Encoding { get; set; }
}
