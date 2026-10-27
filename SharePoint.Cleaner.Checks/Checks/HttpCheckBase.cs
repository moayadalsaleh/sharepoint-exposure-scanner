using System.Diagnostics;
using System.Text.Json;
using SharePoint.Cleaner.Core.Contracts;
using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Checks.Checks;

public abstract class HttpCheckBase : IScanCheck
{
    public abstract string CheckId { get; }
    public abstract string Title { get; }

    // Keeps the target's own path so scanning https://host/ar probes /ar/_api/... instead of the root web.
    protected static Uri Combine(Uri baseUri, string relativePath)
    {
        var normalized = relativePath.StartsWith('/') ? relativePath : "/" + relativePath;
        var basePath = baseUri.AbsolutePath.TrimEnd('/');
        return new Uri(baseUri.GetLeftPart(UriPartial.Authority) + basePath + normalized);
    }

    protected sealed record JsonProbe(Uri Endpoint, int? StatusCode, JsonElement? Root, TimeSpan Duration, string? Error)
    {
        public bool Succeeded => Root is not null;

        public bool Failed => Error is not null || StatusCode >= 500;

        public IReadOnlyList<JsonElement> Values
        {
            get
            {
                if (Root is { } root && SharePointJson.TryGetArray(root, out var array))
                {
                    return array.EnumerateArray().ToList();
                }

                return [];
            }
        }
    }

    protected static async Task<JsonProbe> ProbeJsonAsync(HttpClient httpClient, Uri endpoint, CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            using var response = await SharePointJson.GetAsync(httpClient, endpoint, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);

            JsonElement? root = null;
            if (response.IsSuccessStatusCode)
            {
                try
                {
                    using var doc = JsonDocument.Parse(body);
                    root = doc.RootElement.Clone();
                }
                catch (JsonException)
                {
                }
            }

            return new JsonProbe(endpoint, (int)response.StatusCode, root, sw.Elapsed, null);
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            return new JsonProbe(endpoint, null, null, sw.Elapsed, ex.Message);
        }
    }

    protected static string? GetString(JsonElement element, string name) =>
        element.ValueKind == JsonValueKind.Object && element.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.String
            ? value.GetString()
            : null;

    protected ScanCheckResult BuildResult(
        RiskLevel risk,
        CheckOutcome outcome,
        Uri endpoint,
        int? statusCode,
        TimeSpan duration,
        string evidence,
        string recommendation) => new()
    {
        CheckId = CheckId,
        Title = Title,
        RiskLevel = risk,
        Outcome = outcome,
        Endpoint = endpoint.ToString(),
        StatusCode = statusCode,
        Duration = duration,
        Evidence = evidence,
        Recommendation = recommendation
    };

    protected static async Task<(int StatusCode, string Body, TimeSpan Duration)> GetAsync(
        HttpClient httpClient,
        Uri endpoint,
        CancellationToken cancellationToken)
    {
        var sw = Stopwatch.StartNew();
        using var response = await httpClient.GetAsync(endpoint, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        sw.Stop();
        return ((int)response.StatusCode, body, sw.Elapsed);
    }

    public abstract Task<ScanCheckResult> ExecuteAsync(ScanContext context, CancellationToken cancellationToken);
}
