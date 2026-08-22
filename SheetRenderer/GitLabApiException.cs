using System;

internal sealed class GitLabApiException : InvalidOperationException
{
    public GitLabApiException(
        string message,
        int statusCode,
        string url,
        string responseBody)
        : base(message)
    {
        StatusCode = statusCode;
        Url = url;
        ResponseBody = responseBody;
    }

    public int StatusCode { get; private set; }

    public string Url { get; private set; }

    public string ResponseBody { get; private set; }
}
