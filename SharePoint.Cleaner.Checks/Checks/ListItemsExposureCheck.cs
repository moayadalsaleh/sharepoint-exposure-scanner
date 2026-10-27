using System.Text.Json;
using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public sealed class ListItemsExposureCheck : HttpCheckBase
{
    public override string CheckId => "SP-H-003";
    public override string Title => "SharePoint lists and items data exposure";

    private const string ListsPath = "/_api/web/lists?$select=Title,ItemCount,Hidden&$filter=Hidden eq false";

    public override async Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken)
    {
        var listEndpoint = Combine(context.TargetUri, ListsPath);
        var findings = new List<ListItemFinding>();

        try
        {
            using var listResponse = await SharePointJson.GetAsync(context.HttpClient, listEndpoint, cancellationToken);
            var listBody = await listResponse.Content.ReadAsStringAsync(cancellationToken);

            if ((int)listResponse.StatusCode == 401 || (int)listResponse.StatusCode == 403)
            {
                return new ScanCheckResult
                {
                    CheckId = CheckId,
                    Title = Title,
                    RiskLevel = RiskLevel.Critical,
                    Outcome = CheckOutcome.Secured,
                    Endpoint = listEndpoint.ToString(),
                    StatusCode = (int)listResponse.StatusCode,
                    Evidence = "Lists endpoint is protected.",
                    Recommendation = "Keep list APIs restricted to required authenticated users only.",
                    Duration = TimeSpan.Zero,
                    ListItemFindings = []
                };
            }

            if (!listResponse.IsSuccessStatusCode)
            {
                return new ScanCheckResult
                {
                    CheckId = CheckId,
                    Title = Title,
                    RiskLevel = RiskLevel.Critical,
                    Outcome = CheckOutcome.Error,
                    Endpoint = listEndpoint.ToString(),
                    StatusCode = (int)listResponse.StatusCode,
                    Evidence = $"Unexpected status code from lists API: {(int)listResponse.StatusCode}",
                    Recommendation = "Verify endpoint compatibility and authentication mode.",
                    Duration = TimeSpan.Zero,
                    ListItemFindings = []
                };
            }

            using var listJson = JsonDocument.Parse(listBody);
            if (!SharePointJson.TryGetArray(listJson.RootElement, out var listsNode))
            {
                return new ScanCheckResult
                {
                    CheckId = CheckId,
                    Title = Title,
                    RiskLevel = RiskLevel.High,
                    Outcome = CheckOutcome.Error,
                    Endpoint = listEndpoint.ToString(),
                    StatusCode = (int)listResponse.StatusCode,
                    Evidence = "Could not parse lists response.",
                    Recommendation = "Confirm this target is SharePoint and supports REST list endpoints.",
                    Duration = TimeSpan.Zero,
                    ListItemFindings = []
                };
            }

            // Keep MVP bounded: inspect up to 3 lists, pull up to 3 items from each.
            var candidateTitles = new List<string>();
            foreach (var list in listsNode.EnumerateArray())
            {
                if (candidateTitles.Count >= 3)
                {
                    break;
                }

                var title = list.TryGetProperty("Title", out var t) ? t.GetString() : null;
                var itemCount = list.TryGetProperty("ItemCount", out var c) ? c.GetInt32() : 0;
                if (!string.IsNullOrWhiteSpace(title) && itemCount > 0)
                {
                    candidateTitles.Add(title);
                }
            }

            foreach (var title in candidateTitles)
            {
                var escapedTitle = title.Replace("'", "''", StringComparison.Ordinal);
                var itemEndpoint = Combine(context.TargetUri,
                    $"/_api/web/lists/getbytitle('{Uri.EscapeDataString(escapedTitle)}')/items?$top=3");

                using var itemResponse = await SharePointJson.GetAsync(context.HttpClient, itemEndpoint, cancellationToken);
                var itemBody = await itemResponse.Content.ReadAsStringAsync(cancellationToken);

                if (!itemResponse.IsSuccessStatusCode)
                {
                    continue;
                }

                using var itemJson = JsonDocument.Parse(itemBody);
                if (!SharePointJson.TryGetArray(itemJson.RootElement, out var itemsNode))
                {
                    continue;
                }

                foreach (var item in itemsNode.EnumerateArray())
                {
                    var id = item.TryGetProperty("Id", out var idNode)
                        ? idNode.ToString()
                        : "N/A";
                    findings.Add(new ListItemFinding
                    {
                        ListTitle = title,
                        ItemId = id,
                        Endpoint = itemEndpoint.ToString(),
                        Preview = BuildItemPreview(item)
                    });
                }
            }

            var exposed = findings.Count > 0;
            var visibleLists = listsNode.EnumerateArray().Count();
            var totalItems = listsNode.EnumerateArray().Sum(l => l.TryGetProperty("ItemCount", out var ic) && ic.ValueKind == JsonValueKind.Number ? ic.GetInt64() : 0);
            var readLists = findings.GroupBy(f => f.ListTitle)
                .Select(g => $"{g.Key} ({g.Count()} sample item(s) read)")
                .ToList();
            return new ScanCheckResult
            {
                CheckId = CheckId,
                Title = Title,
                RiskLevel = RiskLevel.Critical,
                Outcome = exposed ? CheckOutcome.Exposed : CheckOutcome.Secured,
                Endpoint = listEndpoint.ToString(),
                StatusCode = (int)listResponse.StatusCode,
                Evidence = exposed
                    ? $"The lists API disclosed {visibleLists} visible list(s) holding about {totalItems:N0} item(s) in total (document-library files included). " +
                      $"Item records were successfully read from: {string.Join("; ", readLists)}."
                    : $"The lists API disclosed {visibleLists} visible list(s) (~{totalItems:N0} items) but no item records could be read.",
                Recommendation = "Restrict list and item APIs, enforce least privilege, and disable anonymous exposure.",
                Duration = TimeSpan.Zero,
                ListItemFindings = findings
            };
        }
        catch (Exception ex)
        {
            return new ScanCheckResult
            {
                CheckId = CheckId,
                Title = Title,
                RiskLevel = RiskLevel.Critical,
                Outcome = CheckOutcome.Error,
                Endpoint = listEndpoint.ToString(),
                Evidence = ex.Message,
                Recommendation = "Validate API availability and authentication context.",
                Duration = TimeSpan.Zero,
                ListItemFindings = findings
            };
        }
    }

    private static string BuildItemPreview(JsonElement item)
    {
        var segments = new List<string>();
        foreach (var prop in item.EnumerateObject())
        {
            if (segments.Count >= 5)
            {
                break;
            }

            var value = prop.Value.ValueKind switch
            {
                JsonValueKind.String => prop.Value.GetString(),
                JsonValueKind.Number => prop.Value.ToString(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                _ => null
            };

            if (!string.IsNullOrWhiteSpace(value))
            {
                segments.Add($"{prop.Name}: {value}");
            }
        }

        return segments.Count == 0 ? "No preview fields available." : string.Join(" | ", segments);
    }
}
