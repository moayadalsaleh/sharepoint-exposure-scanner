using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class VersionDisclosureCheck : HttpCheckBase
{
    private static readonly string[] Headers = ["MicrosoftSharePointTeamServices", "X-AspNet-Version", "X-Powered-By"];

    public override string CheckId => "SP-L-002";
    public override string Title => "Version disclosure in HTTP headers";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var endpoint = Combine(context.TargetUri, "/");
        var sw = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            using var response = await context.HttpClient.GetAsync(endpoint, cancellationToken);
            var found = new List<string>();
            foreach (var name in Headers)
            {
                if (response.Headers.TryGetValues(name, out var values))
                {
                    found.Add($"{name}: {string.Join(",", values)}");
                }
            }

            var exposed = found.Count > 0;
            return BuildResult(RiskLevel.Low, exposed ? CheckOutcome.Exposed : CheckOutcome.Secured, endpoint,
                (int)response.StatusCode, sw.Elapsed,
                exposed ? "Response headers disclose: " + string.Join("; ", found) + "." + DescribeVersion(found) : "No version-disclosing headers were returned.",
                "Remove version headers at IIS or the reverse proxy and keep the farm patched.");
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return BuildResult(RiskLevel.Low, CheckOutcome.Error, endpoint, null, sw.Elapsed, ex.Message,
                "Verify connectivity and re-run the scan.");
        }
    }

    private static string DescribeVersion(IEnumerable<string> found)
    {
        var header = found.FirstOrDefault(f => f.StartsWith("MicrosoftSharePointTeamServices:", StringComparison.Ordinal));
        var described = SharePointVersionInfo.Describe(header?["MicrosoftSharePointTeamServices:".Length..].Trim());
        return described is null ? string.Empty : $" Detected: {described}.";
    }
}
