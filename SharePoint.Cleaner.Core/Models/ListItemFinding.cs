namespace SharePoint.Cleaner.Core.Models;

public sealed class ListItemFinding
{
    public required string ListTitle { get; init; }
    public required string ItemId { get; init; }
    public required string Endpoint { get; init; }
    public required string Preview { get; init; }
}
