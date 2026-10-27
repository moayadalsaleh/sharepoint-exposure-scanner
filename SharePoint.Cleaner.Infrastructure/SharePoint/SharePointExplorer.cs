using System.Text.Json;
using SharePoint.Cleaner.Core.Contracts;
using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.Infrastructure.SharePoint;

public sealed class SharePointExplorer : ISharePointExplorer
{
    private readonly Func<HttpClient> _clientFactory;

    public SharePointExplorer()
        : this(CreateClient)
    {
    }

    public SharePointExplorer(Func<HttpClient> clientFactory)
    {
        _clientFactory = clientFactory ?? throw new ArgumentNullException(nameof(clientFactory));
    }

    public async Task<IReadOnlyList<SharePointListSummary>> GetListsAsync(Uri siteUri, CancellationToken cancellationToken)
    {
        using var client = _clientFactory();
        var endpoint = BuildApiUri(siteUri, "web/lists?$select=Title,ItemCount,Hidden&$filter=Hidden eq false");
        using var response = await SharePointJson.GetAsync(client, endpoint, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);

        response.EnsureSuccessStatusCode();

        using var json = JsonDocument.Parse(body);
        var results = new List<SharePointListSummary>();
        if (!SharePointJson.TryGetArray(json.RootElement, out var value))
        {
            return results;
        }

        foreach (var node in value.EnumerateArray())
        {
            results.Add(new SharePointListSummary
            {
                Title = node.TryGetProperty("Title", out var t) ? t.GetString() ?? string.Empty : string.Empty,
                ItemCount = node.TryGetProperty("ItemCount", out var c) ? c.GetInt32() : 0,
                Hidden = node.TryGetProperty("Hidden", out var h) && h.GetBoolean()
            });
        }

        return results
            .Where(x => !string.IsNullOrWhiteSpace(x.Title))
            .OrderBy(x => x.Title, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public async Task<IReadOnlyList<SharePointListItemRecord>> GetListItemsAsync(
        Uri siteUri,
        string listTitle,
        int top,
        int minItemId,
        CancellationToken cancellationToken)
    {
        using var client = _clientFactory();

        var normalizedTop = Math.Clamp(top, 1, 20000);
        var escapedTitle = Uri.EscapeDataString(listTitle.Replace("'", "''", StringComparison.Ordinal));
        var results = new List<SharePointListItemRecord>();
        var nextMinId = Math.Max(0, minItemId);

        // Page by Id (500 per request) so large lists load without hitting server paging limits.
        while (results.Count < normalizedTop)
        {
            var pageSize = Math.Min(500, normalizedTop - results.Count);
            var endpoint = BuildApiUri(siteUri,
                $"web/lists/getbytitle('{escapedTitle}')/items?$top={pageSize}&$filter=Id ge {nextMinId}&$orderby=Id asc");
            using var response = await SharePointJson.GetAsync(client, endpoint, cancellationToken);
            var body = await response.Content.ReadAsStringAsync(cancellationToken);

            response.EnsureSuccessStatusCode();

            using var json = JsonDocument.Parse(body);
            if (!SharePointJson.TryGetArray(json.RootElement, out var value))
            {
                break;
            }

            var pageCount = 0;
            long lastId = nextMinId;
            foreach (var item in value.EnumerateArray())
            {
                pageCount++;
                var allFields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var prop in item.EnumerateObject())
                {
                    if (prop.Value.ValueKind is not JsonValueKind.Object and not JsonValueKind.Array)
                        allFields[prop.Name] = prop.Value.ToString();
                }

                if (item.TryGetProperty("Id", out var idValue) && idValue.TryGetInt64(out var parsedId))
                {
                    lastId = parsedId;
                }

                results.Add(new SharePointListItemRecord
                {
                    ListTitle = listTitle,
                    Id = item.TryGetProperty("Id", out var id) ? id.ToString() : "N/A",
                    Title = item.TryGetProperty("Title", out var title) ? title.ToString() : null,
                    FileRef = item.TryGetProperty("FileRef", out var fileRef) ? fileRef.ToString() : null,
                    Created = item.TryGetProperty("Created", out var created) ? created.ToString() : null,
                    Modified = item.TryGetProperty("Modified", out var modified) ? modified.ToString() : null,
                    Preview = BuildPreview(item),
                    AllFields = allFields
                });
            }

            if (pageCount < pageSize || lastId < nextMinId)
            {
                break;
            }

            nextMinId = (int)Math.Min(int.MaxValue, lastId + 1);
        }

        return results;
    }

    public async Task<IReadOnlyList<SharePointSiteNode>> GetSiteTreeAsync(Uri siteUri, CancellationToken cancellationToken)
    {
        var state = new TreeState();
        return [await LoadSiteRecursiveAsync(_clientFactory, siteUri, siteUri, null, 0, 6, state, cancellationToken)];
    }

    private sealed class TreeState
    {
        public SemaphoreSlim Throttle { get; } = new(8);
        public volatile bool WebsUnavailable;
    }

    private sealed record ChildWeb(Uri Url, string? Title);

    private static async Task<SharePointSiteNode> LoadSiteRecursiveAsync(
        Func<HttpClient> clientFactory,
        Uri rootUri,
        Uri currentUri,
        string? knownTitle,
        int depth,
        int maxDepth,
        TreeState state,
        CancellationToken cancellationToken)
    {
        using var client = clientFactory();

        var title = knownTitle;
        if (string.IsNullOrWhiteSpace(title))
        {
            await state.Throttle.WaitAsync(cancellationToken);
            try
            {
                title = await GetWebTitleAsync(client, currentUri, cancellationToken);
            }
            finally
            {
                state.Throttle.Release();
            }
        }

        title ??= currentUri.AbsolutePath.Trim('/');
        SharePointSiteNode[] children = [];

        if (depth < maxDepth)
        {
            var childWebs = await GetChildWebsAsync(client, currentUri, state, cancellationToken);
            // Keep traversal inside same host and avoid loops.
            var valid = childWebs.Where(c =>
                string.Equals(c.Url.Host, rootUri.Host, StringComparison.OrdinalIgnoreCase)
                && !string.Equals(c.Url.AbsoluteUri.TrimEnd('/'), currentUri.AbsoluteUri.TrimEnd('/'), StringComparison.OrdinalIgnoreCase));

            children = await Task.WhenAll(valid.Select(c =>
                LoadSiteRecursiveAsync(clientFactory, rootUri, c.Url, c.Title, depth + 1, maxDepth, state, cancellationToken)));
        }

        return new SharePointSiteNode
        {
            Title = string.IsNullOrWhiteSpace(title) ? currentUri.AbsoluteUri : title,
            Url = currentUri.AbsoluteUri,
            Children = children
        };
    }

    private static async Task<IReadOnlyList<ChildWeb>> GetChildWebsAsync(HttpClient client, Uri currentUri, TreeState state, CancellationToken cancellationToken)
    {
        // "webs" needs permissions anonymous users often lack; "webinfos" is usually readable.
        foreach (var (resource, property) in new[] { ("web/webs?$select=Title,Url", "Url"), ("web/webinfos?$select=Title,ServerRelativeUrl", "ServerRelativeUrl") })
        {
            if (property == "Url" && state.WebsUnavailable)
            {
                continue;
            }

            try
            {
                HttpResponseMessage response;
                await state.Throttle.WaitAsync(cancellationToken);
                try
                {
                    response = await SharePointJson.GetAsync(client, BuildApiUri(currentUri, resource), cancellationToken);
                }
                finally
                {
                    state.Throttle.Release();
                }

                using var _ = response;
                if (!response.IsSuccessStatusCode)
                {
                    if (property == "Url" && response.StatusCode is System.Net.HttpStatusCode.Unauthorized or System.Net.HttpStatusCode.Forbidden)
                    {
                        state.WebsUnavailable = true;
                    }

                    continue;
                }

                var body = await response.Content.ReadAsStringAsync(cancellationToken);
                using var json = JsonDocument.Parse(body);
                if (!SharePointJson.TryGetArray(json.RootElement, out var value))
                {
                    continue;
                }

                var result = new List<ChildWeb>();
                foreach (var child in value.EnumerateArray())
                {
                    var url = child.TryGetProperty(property, out var u) ? u.GetString() : null;
                    if (string.IsNullOrWhiteSpace(url))
                    {
                        continue;
                    }

                    if (Uri.TryCreate(url, UriKind.Absolute, out var childUri)
                        || Uri.TryCreate(new Uri(currentUri.GetLeftPart(UriPartial.Authority)), url, out childUri))
                    {
                        result.Add(new ChildWeb(childUri, child.TryGetProperty("Title", out var t) ? t.GetString() : null));
                    }
                }

                return result;
            }
            catch (Exception ex) when (ex is HttpRequestException or JsonException or TaskCanceledException && !cancellationToken.IsCancellationRequested)
            {
            }
        }

        return [];
    }

    private static async Task<string?> GetWebTitleAsync(HttpClient client, Uri currentUri, CancellationToken cancellationToken)
    {
        var endpoint = BuildApiUri(currentUri, "web?$select=Title");
        using var response = await SharePointJson.GetAsync(client, endpoint, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        using var json = JsonDocument.Parse(body);
        var node = SharePointJson.Unwrap(json.RootElement);
        return node.ValueKind == JsonValueKind.Object && node.TryGetProperty("Title", out var title) ? title.GetString() : null;
    }

    private static Uri BuildApiUri(Uri siteUri, string resource)
    {
        return new Uri($"{siteUri.AbsoluteUri.TrimEnd('/')}/_api/{resource}");
    }

    private static string BuildPreview(JsonElement item)
    {
        var parts = new List<string>();
        foreach (var prop in item.EnumerateObject())
        {
            if (parts.Count >= 6)
            {
                break;
            }

            if (prop.Value.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
            {
                continue;
            }

            var value = prop.Value.ToString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                parts.Add($"{prop.Name}: {value}");
            }
        }

        return parts.Count == 0 ? "No preview fields." : string.Join(" | ", parts);
    }

    private static HttpClient CreateClient()
    {
        var handler = new HttpClientHandler
        {
            UseDefaultCredentials = true,
            PreAuthenticate = true,
            AutomaticDecompression = System.Net.DecompressionMethods.GZip |
                                     System.Net.DecompressionMethods.Deflate
        };
        return new HttpClient(handler);
    }
}
