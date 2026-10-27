using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Core.Contracts;

public interface ISharePointExplorer
{
    Task<IReadOnlyList<SharePointListSummary>> GetListsAsync(Uri siteUri, CancellationToken cancellationToken);
    Task<IReadOnlyList<SharePointListItemRecord>> GetListItemsAsync(
        Uri siteUri,
        string listTitle,
        int top,
        int minItemId,
        CancellationToken cancellationToken);
    Task<IReadOnlyList<SharePointSiteNode>> GetSiteTreeAsync(Uri siteUri, CancellationToken cancellationToken);
}
