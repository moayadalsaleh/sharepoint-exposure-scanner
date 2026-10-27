using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class FrontPageExtensionsCheck : HttpCheckBase
{
    private const string Path = "/_vti_inf.html";

    public override string CheckId => "SP-L-001";
    public override string Title => "FrontPage Server Extensions exposure";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var endpoint = Combine(context.TargetUri, Path);

        try
        {
            var (statusCode, body, duration) = await GetAsync(context.HttpClient, endpoint, cancellationToken);
            var exposed = statusCode == 200 &&
                          (body.Contains("FrontPage Configuration Information", StringComparison.OrdinalIgnoreCase) ||
                           body.Contains("FPVersion", StringComparison.OrdinalIgnoreCase));

            return new ScanCheckResult
            {
                CheckId = CheckId,
                Title = Title,
                RiskLevel = RiskLevel.Low,
                Outcome = exposed ? CheckOutcome.Exposed : CheckOutcome.Secured,
                Endpoint = endpoint.ToString(),
                StatusCode = statusCode,
                Duration = duration,
                Evidence = exposed
                    ? "FrontPage signature detected in response."
                    : "FrontPage signature not detected.",
                Recommendation = "Disable or restrict legacy FrontPage-related endpoints when not required."
            };
        }
        catch (Exception ex)
        {
            return new ScanCheckResult
            {
                CheckId = CheckId,
                Title = Title,
                RiskLevel = RiskLevel.Low,
                Outcome = CheckOutcome.Error,
                Endpoint = endpoint.ToString(),
                Evidence = ex.Message,
                Recommendation = "Verify connectivity and authentication, then retry the check.",
                Duration = TimeSpan.Zero
            };
        }
    }
}
