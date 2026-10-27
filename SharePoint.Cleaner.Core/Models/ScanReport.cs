namespace SharePoint.Cleaner.Core.Models;

public sealed class ScanReport
{
    public required Uri TargetUri { get; init; }
    public required DateTimeOffset StartedAt { get; init; }
    public required DateTimeOffset CompletedAt { get; init; }
    public required IReadOnlyList<ScanCheckResult> Results { get; init; }

    /// <summary>Product inferred from the MicrosoftSharePointTeamServices header, when present.</summary>
    public string? DetectedVersion { get; init; }

    public RiskLevel OverallRisk =>
        Results.Any(r => r.Outcome == CheckOutcome.Exposed && r.RiskLevel == RiskLevel.Critical) ? RiskLevel.Critical :
        Results.Any(r => r.Outcome == CheckOutcome.Exposed && r.RiskLevel == RiskLevel.High) ? RiskLevel.High :
        Results.Any(r => r.Outcome == CheckOutcome.Exposed && r.RiskLevel == RiskLevel.Medium) ? RiskLevel.Medium :
        Results.Any(r => r.Outcome == CheckOutcome.Exposed && r.RiskLevel == RiskLevel.Low) ? RiskLevel.Low :
        RiskLevel.None;

    public int ExposedCount => Results.Count(r => r.Outcome == CheckOutcome.Exposed);
    public int SecuredCount => Results.Count(r => r.Outcome == CheckOutcome.Secured);
    public int ErrorCount => Results.Count(r => r.Outcome == CheckOutcome.Error);
}
