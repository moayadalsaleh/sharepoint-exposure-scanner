namespace SharePoint.Cleaner.Core.Models;

public sealed class ScanRequest
{
    public required Uri TargetUri { get; init; }
    public TimeSpan CheckTimeout { get; init; } = TimeSpan.FromSeconds(12);
}
