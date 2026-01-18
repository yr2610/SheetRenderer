internal sealed class TokenKeyInfo
{
    public string BaseUrl { get; set; }
    public string ProjectId { get; set; }

    public string DisplayText
    {
        get { return BaseUrl + " / " + ProjectId; }
    }
}
