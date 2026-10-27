namespace SharePoint.Cleaner.Core.Models;

public sealed class SharePointListItemRecord
{
    public required string ListTitle { get; init; }
    public required string Id { get; init; }
    public required string? Title { get; init; }
    public required string? FileRef { get; init; }
    public required string? Created { get; init; }
    public required string? Modified { get; init; }
    public required string Preview { get; init; }
    public IReadOnlyDictionary<string, string> AllFields { get; init; } = new Dictionary<string, string>();
}
