namespace SharePoint.Cleaner.Core.Models;

public sealed class SharePointListSummary
{
    public required string Title { get; init; }
    public int ItemCount { get; init; }
    public bool Hidden { get; init; }
}
