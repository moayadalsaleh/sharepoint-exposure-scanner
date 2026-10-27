namespace SharePoint.Cleaner.Core.Models;

public sealed class ScanCheckResult
{
    public required string CheckId { get; init; }
    public required string Title { get; init; }
    public required RiskLevel RiskLevel { get; init; }
    public required CheckOutcome Outcome { get; init; }
    public required string Endpoint { get; init; }
    public required string Evidence { get; init; }
    public required string Recommendation { get; init; }
    public int? StatusCode { get; init; }
    public TimeSpan Duration { get; init; }
    public IReadOnlyList<ListItemFinding> ListItemFindings { get; init; } = [];

    /// <summary>For Exposed results: whether an anonymous re-run also succeeded. Null when not applicable or not verified.</summary>
    public bool? PubliclyAccessible { get; init; }

    public string Scope => Outcome != CheckOutcome.Exposed
        ? string.Empty
        : PubliclyAccessible switch
        {
            true => "Public (anonymous)",
            false => "Authenticated only",
            _ => "Unverified"
        };
}
