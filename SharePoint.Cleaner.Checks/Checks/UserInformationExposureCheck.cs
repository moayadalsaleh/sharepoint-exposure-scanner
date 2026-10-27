using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class UserInformationExposureCheck : HttpCheckBase
{
    private const string Path = "/_layouts/userdisp.aspx?Force=True&id=1";

    public override string CheckId => "SP-H-001";
    public override string Title => "User information page exposure";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var endpoint = Combine(context.TargetUri, Path);

        try
        {
            var (statusCode, body, duration) = await GetAsync(context.HttpClient, endpoint, cancellationToken);
            var exposed = statusCode == 200 && body.Contains("SPFieldText", StringComparison.OrdinalIgnoreCase);

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
                    ? "User profile field markers are visible to the scanner context."
                    : "User profile markers were not exposed.",
                Recommendation = "Restrict anonymous or broad access to userdisp.aspx and user profile endpoints."
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
                Recommendation = "Check site reachability and authentication configuration.",
                Duration = TimeSpan.Zero
            };
        }
    }
}
