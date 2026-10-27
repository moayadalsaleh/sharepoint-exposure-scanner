using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Core.Contracts;

public interface IScanCheck
{
    string CheckId { get; }
    string Title { get; }
    Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken);
}
