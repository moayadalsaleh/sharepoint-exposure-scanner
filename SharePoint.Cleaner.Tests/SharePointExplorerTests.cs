using System.Net;
using System.Net.Http;
using SharePoint.Cleaner.Infrastructure.SharePoint;

namespace SharePoint.Cleaner.Tests;

public class SharePointExplorerTests
{
    [Fact]
    public async Task GetSiteTreeAsync_ShouldKeepSubsitePathInApiRequests()
    {
        var requestedUris = new List<Uri>();
        using var client = new HttpClient(new StubHandler(requestedUris));
        var explorer = new SharePointExplorer(() => client);

        var nodes = await explorer.GetSiteTreeAsync(
            new Uri("https://www.example.com"),
            CancellationToken.None);

        var root = Assert.Single(nodes);
        var arabicSite = Assert.Single(root.Children);

        Assert.Equal("https://www.example.com/ar", arabicSite.Url);
        Assert.Equal("Arabic", arabicSite.Title);
        Assert.DoesNotContain(requestedUris, uri => uri.AbsoluteUri == "https://www.example.com/ar/_api/web?$select=Title");
        Assert.Single(requestedUris, uri => uri.AbsoluteUri == "https://www.example.com/_api/web?$select=Title");
    }

    [Fact]
    public async Task GetSiteTreeAsync_ShouldFallbackToWebInfosWhenWebsIsUnauthorized()
    {
        using var client = new HttpClient(new WebInfosHandler());
        var explorer = new SharePointExplorer(() => client);

        var nodes = await explorer.GetSiteTreeAsync(new Uri("http://www.example.ae"), CancellationToken.None);

        var root = Assert.Single(nodes);
        var ar = Assert.Single(root.Children);
        Assert.Equal("http://www.example.ae/ar", ar.Url);
        var academics = Assert.Single(ar.Children);
        Assert.Equal("http://www.example.ae/ar/Academics", academics.Url);
    }

    private sealed class WebInfosHandler : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var path = request.RequestUri!.AbsolutePath;
            if (path.EndsWith("/web/webs", StringComparison.Ordinal))
            {
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Unauthorized));
            }

            var body = path switch
            {
                "/_api/web/webinfos" => "{\"value\":[{\"ServerRelativeUrl\":\"/ar\",\"Title\":\"AR\"}]}",
                "/ar/_api/web/webinfos" => "{\"value\":[{\"ServerRelativeUrl\":\"/ar/Academics\",\"Title\":\"Acad\"}]}",
                _ when path.EndsWith("/_api/web", StringComparison.Ordinal) => "{\"Title\":\"T\"}",
                _ => "{\"value\":[]}"
            };

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(body) });
        }
    }

    private sealed class StubHandler(List<Uri> requestedUris) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            requestedUris.Add(request.RequestUri!);
            var path = request.RequestUri!.AbsolutePath;

            var body = path.EndsWith("/_api/web/webs", StringComparison.Ordinal)
                ? path == "/_api/web/webs"
                    ? "{\"value\":[{\"Title\":\"Arabic\",\"Url\":\"https://www.example.com/ar\"}]}"
                    : "{\"value\":[]}"
                : path.EndsWith("/_api/web", StringComparison.Ordinal)
                    ? path == "/_api/web"
                        ? "{\"Title\":\"MOE\"}"
                        : "{\"Title\":\"Arabic\"}"
                    : "{\"value\":[]}";

            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(body)
            });
        }
    }
}