using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class SiteStructureExposureCheck : HttpCheckBase
{
    public override string CheckId => "SP-M-002";
    public override string Title => "Subsite enumeration via REST";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var webs = await ProbeJsonAsync(context.HttpClient,
            Combine(context.TargetUri, "/_api/web/webs?$select=Title,ServerRelativeUrl&$top=100"), cancellationToken);
        var infos = await ProbeJsonAsync(context.HttpClient,
            Combine(context.TargetUri, "/_api/web/webinfos?$select=Title,ServerRelativeUrl&$top=100"), cancellationToken);

        var findings = new List<string>();
        foreach (var (name, probe) in new[] { ("webs", webs), ("webinfos", infos) })
        {
            if (!probe.Succeeded)
            {
                continue;
            }

            var urls = probe.Values.Select(v => GetString(v, "ServerRelativeUrl")).Where(u => u is not null).Take(5);
            findings.Add($"{name} returned {probe.Values.Count} subsite(s) (e.g. {string.Join(", ", urls)})");
        }

        var exposed = findings.Count > 0;
        var outcome = exposed ? CheckOutcome.Exposed
            : webs.Failed && infos.Failed ? CheckOutcome.Error
            : CheckOutcome.Secured;

        var evidence = exposed
            ? string.Join("; ", findings) + "."
            : outcome == CheckOutcome.Error
                ? "Both probes failed (network error or HTTP 5xx)."
                : $"webs returned HTTP {webs.StatusCode}, webinfos returned HTTP {infos.StatusCode}; no subsites were disclosed.";

        return BuildResult(RiskLevel.Medium, outcome, infos.Endpoint, infos.StatusCode ?? webs.StatusCode,
            webs.Duration + infos.Duration, evidence,
            "Restrict anonymous access and block site-structure REST endpoints for anonymous users.");
    }
}
