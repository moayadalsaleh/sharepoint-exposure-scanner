using System.Text.Json;

namespace SharePoint.Cleaner.Core.Models;

/// <summary>
/// Helpers for talking to SharePoint 2013/2016/2019/SE/Online REST.
/// Newer farms honour odata=nometadata; SharePoint 2013 RTM and some proxies only accept odata=verbose,
/// which wraps results in { "d": { "results": [...] } }.
/// </summary>
public static class SharePointJson
{
    public static async Task<HttpResponseMessage> GetAsync(HttpClient client, Uri uri, CancellationToken cancellationToken)
    {
        var response = await SendAsync(client, uri, "application/json;odata=nometadata", cancellationToken);
        if (response.StatusCode is System.Net.HttpStatusCode.BadRequest
            or System.Net.HttpStatusCode.NotAcceptable
            or System.Net.HttpStatusCode.UnsupportedMediaType)
        {
            response.Dispose();
            response = await SendAsync(client, uri, "application/json;odata=verbose", cancellationToken);
        }

        return response;
    }

    private static Task<HttpResponseMessage> SendAsync(HttpClient client, Uri uri, string accept, CancellationToken cancellationToken)
    {
        var request = new HttpRequestMessage(HttpMethod.Get, uri);
        request.Headers.Accept.ParseAdd(accept);
        return client.SendAsync(request, cancellationToken);
    }

    /// <summary>Returns the object itself for the verbose "d" wrapper, otherwise the root.</summary>
    public static JsonElement Unwrap(JsonElement root) =>
        root.ValueKind == JsonValueKind.Object && root.TryGetProperty("d", out var d) ? d : root;

    /// <summary>Finds the record array in nometadata ("value"), verbose ("d.results") or plain array responses.</summary>
    public static bool TryGetArray(JsonElement root, out JsonElement array)
    {
        var node = Unwrap(root);
        if (node.ValueKind == JsonValueKind.Array)
        {
            array = node;
            return true;
        }

        foreach (var name in new[] { "value", "results" })
        {
            if (node.ValueKind == JsonValueKind.Object && node.TryGetProperty(name, out var found) && found.ValueKind == JsonValueKind.Array)
            {
                array = found;
                return true;
            }
        }

        array = default;
        return false;
    }
}
