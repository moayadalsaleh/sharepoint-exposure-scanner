using System.Net;
using SharePoint.Cleaner.Checks.Checks;
using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Tests;

public class NewChecksTests
{
    private sealed class StubHandler(Func<Uri, (HttpStatusCode, string)> respond) : HttpMessageHandler
    {
        public List<Uri> Requests { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Requests.Add(request.RequestUri!);
            var (code, body) = respond(request.RequestUri!);
            return Task.FromResult(new HttpResponseMessage(code) { Content = new StringContent(body) });
        }
    }

    [Fact]
    public async Task UserEnumeration_ShouldBeExposed_AndKeepSubsitePath()
    {
        var handler = new StubHandler(u => u.AbsolutePath.EndsWith("siteusers")
            ? (HttpStatusCode.OK, "{\"value\":[{\"Title\":\"Alice\"}]}")
            : (HttpStatusCode.Unauthorized, ""));
        var context = new ScanContext(new Uri("https://host/ar"), new HttpClient(handler));

        var result = await new UserEnumerationExposureCheck().ExecuteAsync(context, CancellationToken.None);

        Assert.Equal(CheckOutcome.Exposed, result.Outcome);
        Assert.All(handler.Requests, r => Assert.StartsWith("/ar/_api/", r.AbsolutePath));
    }

    [Fact]
    public async Task UserEnumeration_ShouldParseSharePoint2013VerboseFormat()
    {
        var handler = new StubHandler(u => u.AbsolutePath.EndsWith("siteusers")
            ? (HttpStatusCode.OK, "{\"d\":{\"results\":[{\"Title\":\"Alice\"},{\"Title\":\"Bob\"}]}}")
            : (HttpStatusCode.Unauthorized, ""));
        var context = new ScanContext(new Uri("https://host/"), new HttpClient(handler));

        var result = await new UserEnumerationExposureCheck().ExecuteAsync(context, CancellationToken.None);

        Assert.Equal(CheckOutcome.Exposed, result.Outcome);
        Assert.Contains("2 user record", result.Evidence);
    }

    [Fact]
    public async Task SharePointJson_ShouldRetryWithVerbose_WhenNoMetadataRejected()
    {
        var accepts = new List<string>();
        var handler = new AcceptCapturingHandler(accepts);

        using var response = await SharePointJson.GetAsync(new HttpClient(handler), new Uri("https://host/_api/web"), CancellationToken.None);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        Assert.Equal(2, accepts.Count);
        Assert.Contains("odata=verbose", accepts[1]);
    }

    private sealed class AcceptCapturingHandler(List<string> accepts) : HttpMessageHandler
    {
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var accept = request.Headers.Accept.ToString();
            accepts.Add(accept);
            var code = accept.Contains("nometadata") ? HttpStatusCode.UnsupportedMediaType : HttpStatusCode.OK;
            return Task.FromResult(new HttpResponseMessage(code) { Content = new StringContent("{}") });
        }
    }

    [Theory]
    [InlineData("15.0.0.4569", "SharePoint 2013")]
    [InlineData("16.0.0.4351", "SharePoint 2016")]
    [InlineData("16.0.0.10337", "SharePoint 2019")]
    [InlineData("16.0.0.14326", "Subscription Edition")]
    public void VersionInfo_ShouldMapBuildToProduct(string header, string expected)
    {
        Assert.Contains(expected, SharePointVersionInfo.Describe(header));
    }

    [Fact]
    public async Task SearchRest_ShouldBeSecured_WhenUnauthorized()
    {
        var handler = new StubHandler(_ => (HttpStatusCode.Unauthorized, ""));
        var context = new ScanContext(new Uri("https://host/"), new HttpClient(handler));

        var result = await new SearchRestExposureCheck().ExecuteAsync(context, CancellationToken.None);

        Assert.Equal(CheckOutcome.Secured, result.Outcome);
    }
}
