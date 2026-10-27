namespace SharePoint.Cleaner.Core.Models;

public sealed class ScanContext
{
    public ScanContext(Uri targetUri, HttpClient httpClient)
    {
        TargetUri = targetUri;
        HttpClient = httpClient;
    }

    public Uri TargetUri { get; }
    public HttpClient HttpClient { get; }
}
