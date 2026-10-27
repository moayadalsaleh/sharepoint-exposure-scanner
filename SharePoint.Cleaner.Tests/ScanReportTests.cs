using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Tests;

public class ScanReportTests
{
    [Fact]
    public void OverallRisk_ShouldPreferHighestExposedRisk()
    {
        var report = BuildReport(
            new ScanCheckResult
            {
                CheckId = "A",
                Title = "Low",
                RiskLevel = RiskLevel.Low,
                Outcome = CheckOutcome.Exposed,
                Endpoint = "https://example.test/a",
                Evidence = "e",
                Recommendation = "r",
                Duration = TimeSpan.FromMilliseconds(1)
            },
            new ScanCheckResult
            {
                CheckId = "B",
                Title = "High",
                RiskLevel = RiskLevel.High,
                Outcome = CheckOutcome.Exposed,
                Endpoint = "https://example.test/b",
                Evidence = "e",
                Recommendation = "r",
                Duration = TimeSpan.FromMilliseconds(1)
            });

        Assert.Equal(RiskLevel.High, report.OverallRisk);
        Assert.Equal(2, report.ExposedCount);
    }

    [Fact]
    public void OverallRisk_ShouldBeNone_WhenNoExposedChecks()
    {
        var report = BuildReport(
            new ScanCheckResult
            {
                CheckId = "A",
                Title = "Secured",
                RiskLevel = RiskLevel.Medium,
                Outcome = CheckOutcome.Secured,
                Endpoint = "https://example.test/a",
                Evidence = "e",
                Recommendation = "r",
                Duration = TimeSpan.FromMilliseconds(1)
            },
            new ScanCheckResult
            {
                CheckId = "B",
                Title = "Error",
                RiskLevel = RiskLevel.High,
                Outcome = CheckOutcome.Error,
                Endpoint = "https://example.test/b",
                Evidence = "e",
                Recommendation = "r",
                Duration = TimeSpan.FromMilliseconds(1)
            });

        Assert.Equal(RiskLevel.None, report.OverallRisk);
        Assert.Equal(1, report.SecuredCount);
        Assert.Equal(1, report.ErrorCount);
    }

    private static ScanReport BuildReport(params ScanCheckResult[] results)
    {
        return new ScanReport
        {
            TargetUri = new Uri("https://example.test"),
            StartedAt = DateTimeOffset.UtcNow,
            CompletedAt = DateTimeOffset.UtcNow,
            Results = results
        };
    }
}
