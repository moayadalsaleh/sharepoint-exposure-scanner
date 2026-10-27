using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Core.Contracts;

public interface IScanEngine
{
    Task<ScanReport> RunAsync(ScanRequest request, CancellationToken cancellationToken);
}
