using System.Runtime.Serialization;

[DataContract]
internal sealed class GitLabLastInput
{
    [DataMember(Name = "baseUrl")]
    public string BaseUrl { get; set; }

    [DataMember(Name = "projectId")]
    public string ProjectId { get; set; }

    [DataMember(Name = "refName")]
    public string RefName { get; set; }

    [DataMember(Name = "filePath")]
    public string FilePath { get; set; }
}
