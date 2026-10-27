using SharePoint.Cleaner.Core.Contracts;
using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Infrastructure.Scanning;

public sealed class DefaultScanEngine : IScanEngine
{
    private readonly IReadOnlyList<IScanCheck> _checks;

    public DefaultScanEngine(IEnumerable<IScanCheck> checks)
    {
        _checks = checks.ToList();
    }

    public async Task<ScanReport> RunAsync(ScanRequest request, CancellationToken cancellationToken)
    {
        var started = DateTimeOffset.UtcNow;

        using var handler = new HttpClientHandler
        {
            UseDefaultCredentials = true,
            PreAuthenticate = true,
            AutomaticDecompression = System.Net.DecompressionMethods.GZip |
                                     System.Net.DecompressionMethods.Deflate
        };
        using var httpClient = new HttpClient(handler);
        var context = new ScanContext(request.TargetUri, httpClient);

        using var anonymousHandler = new HttpClientHandler
        {
            UseDefaultCredentials = false,
            AutomaticDecompression = System.Net.DecompressionMethods.GZip |
                                     System.Net.DecompressionMethods.Deflate
        };
        using var anonymousClient = new HttpClient(anonymousHandler);
        var anonymousContext = new ScanContext(request.TargetUri, anonymousClient);

        string? detectedVersion = null;
        try
        {
            using var probeCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            probeCts.CancelAfter(request.CheckTimeout);
            using var probe = await httpClient.GetAsync(request.TargetUri, probeCts.Token);
            if (probe.Headers.TryGetValues("MicrosoftSharePointTeamServices", out var values))
            {
                detectedVersion = SharePointVersionInfo.Describe(values.FirstOrDefault());
            }
        }
        catch (Exception ex) when (ex is HttpRequestException or OperationCanceledException && !cancellationToken.IsCancellationRequested)
        {
        }

        var results = new List<ScanCheckResult>();
        foreach (var check in _checks)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var checkTimeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            checkTimeoutCts.CancelAfter(request.CheckTimeout);

            try
            {
                var result = await check.ExecuteAsync(context, checkTimeoutCts.Token);
                if (result.Outcome == CheckOutcome.Exposed)
                {
                    result = await VerifyAnonymouslyAsync(check, result, anonymousContext, request.CheckTimeout, cancellationToken);
                }

                results.Add(result);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                results.Add(new ScanCheckResult
                {
                    CheckId = check.CheckId,
                    Title = check.Title,
                    RiskLevel = RiskLevel.Medium,
                    Outcome = CheckOutcome.Error,
                    Endpoint = request.TargetUri.ToString(),
                    Evidence = "Check timed out.",
                    Recommendation = "Increase timeout or test connectivity before scanning.",
                    Duration = request.CheckTimeout
                });
            }
        }

        return new ScanReport
        {
            TargetUri = request.TargetUri,
            StartedAt = started,
            CompletedAt = DateTimeOffset.UtcNow,
            Results = results,
            DetectedVersion = detectedVersion
        };
    }

    // Re-runs an Exposed check without credentials to tell public exposure from authenticated-only access.
    private static async Task<ScanCheckResult> VerifyAnonymouslyAsync(
        IScanCheck check,
        ScanCheckResult authenticated,
        ScanContext anonymousContext,
        TimeSpan timeout,
        CancellationToken cancellationToken)
    {
        bool? isPublic = null;
        try
        {
            using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            cts.CancelAfter(timeout);
            var anonymous = await check.ExecuteAsync(anonymousContext, cts.Token);
            if (anonymous.Outcome == CheckOutcome.Exposed)
            {
                isPublic = true;
            }
            else if (anonymous.Outcome == CheckOutcome.Secured)
            {
                isPublic = false;
            }
        }
        catch (Exception ex) when (ex is not OperationCanceledException || !cancellationToken.IsCancellationRequested)
        {
        }

        return new ScanCheckResult
        {
            CheckId = authenticated.CheckId,
            Title = authenticated.Title,
            RiskLevel = authenticated.RiskLevel,
            Outcome = authenticated.Outcome,
            Endpoint = authenticated.Endpoint,
            Evidence = authenticated.Evidence,
            Recommendation = authenticated.Recommendation,
            StatusCode = authenticated.StatusCode,
            Duration = authenticated.Duration,
            ListItemFindings = authenticated.ListItemFindings,
            PubliclyAccessible = isPublic
        };
    }
}
