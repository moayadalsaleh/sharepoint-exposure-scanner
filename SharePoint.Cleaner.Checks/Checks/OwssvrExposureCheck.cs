using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class OwssvrExposureCheck : HttpCheckBase
{
    private const string Path = "/_vti_bin/owssvr.dll?dialogview=FileOpen&FileDialogFilterValue=*.*&location=&SortField=Created&SortDir=Desc";

    public override string CheckId => "SP-H-002";
    public override string Title => "Owssvr list content exposure";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var endpoint = Combine(context.TargetUri, Path);

        try
        {
            var (statusCode, body, duration) = await GetAsync(context.HttpClient, endpoint, cancellationToken);
            var exposed = statusCode == 200 && body.Contains("All Site Content", StringComparison.OrdinalIgnoreCase);

            return new ScanCheckResult
            {
                CheckId = CheckId,
                Title = Title,
                RiskLevel = RiskLevel.High,
                Outcome = exposed ? CheckOutcome.Exposed : CheckOutcome.Secured,
                Endpoint = endpoint.ToString(),
                StatusCode = statusCode,
                Duration = duration,
                Evidence = exposed
                    ? "Owssvr endpoint returned browsable site content markers."
                    : "Owssvr endpoint did not expose expected content markers.",
                Recommendation = "Restrict anonymous access to owssvr and list browsing endpoints."
            };
        }
        catch (Exception ex)
        {
            return new ScanCheckResult
            {
                CheckId = CheckId,
                Title = Title,
                RiskLevel = RiskLevel.High,
                Outcome = CheckOutcome.Error,
                Endpoint = endpoint.ToString(),
                Evidence = ex.Message,
                Recommendation = "Validate auth/access policy and rerun this high-risk check.",
                Duration = TimeSpan.Zero
            };
        }
    }
}
