using System.Text.Json;
using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class SearchRestExposureCheck : HttpCheckBase
{
    public override string CheckId => "SP-M-003";
    public override string Title => "Anonymous Search REST exposure";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var probe = await ProbeJsonAsync(context.HttpClient,
            Combine(context.TargetUri, "/_api/search/query?querytext='*'&rowlimit=3"), cancellationToken);

        if (probe.Succeeded && TryGetTotalRows(probe.Root!.Value, out var total) && total > 0)
        {
            return BuildResult(RiskLevel.Medium, CheckOutcome.Exposed, probe.Endpoint, probe.StatusCode, probe.Duration,
                $"A wildcard Search REST query returned {total} indexed item(s) without authentication.",
                "Disable or restrict anonymous Search REST queries and exclude non-public content from the index.");
        }

        var outcome = probe.Failed ? CheckOutcome.Error : CheckOutcome.Secured;
        var evidence = probe.Error is not null
            ? probe.Error
            : $"Search REST returned HTTP {probe.StatusCode} with no results.";
        return BuildResult(RiskLevel.Medium, outcome, probe.Endpoint, probe.StatusCode, probe.Duration, evidence,
            "No action required for this check.");
    }

    // Handles both odata=nometadata and verbose response shapes.
    private static bool TryGetTotalRows(JsonElement root, out long total)
    {
        total = 0;
        if (root.ValueKind == JsonValueKind.Object &&
            (root.TryGetProperty("PrimaryQueryResult", out var primary) ||
             (root.TryGetProperty("d", out var d) && d.TryGetProperty("query", out var q) && q.TryGetProperty("PrimaryQueryResult", out primary))) &&
            primary.ValueKind == JsonValueKind.Object &&
            primary.TryGetProperty("RelevantResults", out var relevant) &&
            relevant.TryGetProperty("TotalRows", out var rows) &&
            rows.ValueKind == JsonValueKind.Number)
        {
            total = rows.GetInt64();
            return true;
        }

        return false;
    }
}
