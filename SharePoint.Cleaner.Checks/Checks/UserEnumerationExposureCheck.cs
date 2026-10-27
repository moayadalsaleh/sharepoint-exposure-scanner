using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class UserEnumerationExposureCheck : HttpCheckBase
{
    public override string CheckId => "SP-H-004";
    public override string Title => "User and group enumeration via REST";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var users = await ProbeJsonAsync(context.HttpClient,
            Combine(context.TargetUri, "/_api/web/siteusers?$select=Id,Title&$top=100"), cancellationToken);
        var groups = await ProbeJsonAsync(context.HttpClient,
            Combine(context.TargetUri, "/_api/web/sitegroups?$select=Id,Title&$top=100"), cancellationToken);

        var duration = users.Duration + groups.Duration;
        var findings = new List<string>();
        if (users.Succeeded)
        {
            var names = users.Values.Select(u => GetString(u, "Title")).Where(t => !string.IsNullOrWhiteSpace(t)).Take(5);
            findings.Add($"siteusers returned {users.Values.Count} user record(s) (first names: {string.Join(", ", names)})");
        }

        if (groups.Succeeded)
        {
            var names = groups.Values.Select(g => GetString(g, "Title")).Where(t => !string.IsNullOrWhiteSpace(t)).Take(5);
            findings.Add($"sitegroups returned {groups.Values.Count} group(s) (e.g. {string.Join(", ", names)})");
        }

        var exposed = findings.Count > 0;
        var outcome = exposed ? CheckOutcome.Exposed
            : users.Failed && groups.Failed ? CheckOutcome.Error
            : CheckOutcome.Secured;

        var evidence = exposed
            ? string.Join("; ", findings) + "."
            : outcome == CheckOutcome.Error
                ? $"Both probes failed: {users.Error ?? users.StatusCode?.ToString()}, {groups.Error ?? groups.StatusCode?.ToString()}."
                : $"Users endpoint returned HTTP {users.StatusCode}, groups endpoint returned HTTP {groups.StatusCode}; no records were disclosed.";

        return BuildResult(RiskLevel.High, outcome, users.Endpoint, users.StatusCode ?? groups.StatusCode, duration, evidence,
            "Block anonymous access to user and group enumeration endpoints and review Read permissions on the User Information List.");
    }
}
