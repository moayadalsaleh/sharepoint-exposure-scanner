using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class ListDataServiceExposureCheck : HttpCheckBase
{
    public override string CheckId => "SP-H-005";
    public override string Title => "Legacy ListData.svc exposure";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var endpoint = Combine(context.TargetUri, "/_vti_bin/listdata.svc");
        try
        {
            var (status, body, duration) = await GetAsync(context.HttpClient, endpoint, cancellationToken);
            if (status == 200 && body.Contains("<collection", StringComparison.OrdinalIgnoreCase))
            {
                var count = System.Text.RegularExpressions.Regex.Matches(body, "<collection", System.Text.RegularExpressions.RegexOptions.IgnoreCase).Count;
                return BuildResult(RiskLevel.High, CheckOutcome.Exposed, endpoint, status, duration,
                    $"The ListData.svc service document is readable and advertises {count} queryable list collection(s).",
                    "Block /_vti_bin/listdata.svc for anonymous users and remove anonymous access from non-public lists.");
            }

            var outcome = status >= 500 ? CheckOutcome.Error : CheckOutcome.Secured;
            return BuildResult(RiskLevel.High, outcome, endpoint, status, duration,
                $"ListData.svc returned HTTP {status} without a service document.",
                "No action required for this check.");
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return BuildResult(RiskLevel.High, CheckOutcome.Error, endpoint, null, TimeSpan.Zero, ex.Message,
                "Verify connectivity and re-run the scan.");
        }
    }
}
