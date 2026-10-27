using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class WebServicesExposureCheck : HttpCheckBase
{
    private static readonly string[] ServicePaths =
    [
        "/_vti_bin/lists.asmx",
        "/_vti_bin/usergroup.asmx",
        "/_vti_bin/webs.asmx",
        "/_vti_bin/client.svc"
    ];

    public override string CheckId => "SP-M-001";
    public override string Title => "SharePoint service endpoints exposure";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var exposedEndpoints = new List<string>();
        var errors = new List<string>();
        var totalDuration = TimeSpan.Zero;
        int? firstStatusCode = null;

        foreach (var relativePath in ServicePaths)
        {
            var endpoint = Combine(context.TargetUri, relativePath);
            try
            {
                var (statusCode, body, duration) = await GetAsync(context.HttpClient, endpoint, cancellationToken);
                totalDuration += duration;
                firstStatusCode ??= statusCode;

                var exposed = statusCode == 200 &&
                              (body.Contains("Service Description", StringComparison.OrdinalIgnoreCase) ||
                               body.Contains(".asmx", StringComparison.OrdinalIgnoreCase) ||
                               body.Contains("Endpoint not found", StringComparison.OrdinalIgnoreCase));

                if (exposed)
                {
                    exposedEndpoints.Add(endpoint.ToString());
                }
            }
            catch (Exception ex)
            {
                errors.Add($"{endpoint} -> {ex.Message}");
            }
        }

        var anyExposed = exposedEndpoints.Count > 0;
        var anyError = errors.Count > 0 && !anyExposed;
        var outcome = anyExposed ? CheckOutcome.Exposed : (anyError ? CheckOutcome.Error : CheckOutcome.Secured);
        var evidence = anyExposed
            ? "Exposed endpoints: " + string.Join("; ", exposedEndpoints)
            : errors.Count > 0
                ? "Errors: " + string.Join("; ", errors)
                : "No exposed service endpoints detected by MVP probe set.";

        return new ScanCheckResult
        {
            CheckId = CheckId,
            Title = Title,
            RiskLevel = RiskLevel.Medium,
            Outcome = outcome,
            Endpoint = context.TargetUri.ToString(),
            StatusCode = firstStatusCode,
            Duration = totalDuration,
            Evidence = evidence,
            Recommendation = "Limit anonymous access and enforce least-privilege access on SharePoint web services."
        };
    }
}
