namespace SharePoint.Cleaner.Core.Models;

public static class SharePointVersionInfo
{
    /// <summary>Maps the MicrosoftSharePointTeamServices header (e.g. "16.0.0.4351") to a product name.</summary>
    public static string? Describe(string? headerValue)
    {
        if (string.IsNullOrWhiteSpace(headerValue) || !Version.TryParse(headerValue.Trim(), out var v))
        {
            return null;
        }

        // Header format is major.minor.0.build, e.g. 16.0.0.10337.
        var build = v.Revision >= 0 ? v.Revision : Math.Max(v.Build, 0);
        var product = v.Major switch
        {
            12 => "SharePoint 2007",
            14 => "SharePoint 2010",
            15 => "SharePoint 2013",
            16 when build < 10000 => "SharePoint 2016",
            16 when build < 14326 => "SharePoint 2019",
            16 => "SharePoint Subscription Edition / Online",
            _ => "SharePoint"
        };

        return $"{product} (build {v}, inferred from header)";
    }
}
