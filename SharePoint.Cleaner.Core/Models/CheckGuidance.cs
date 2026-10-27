namespace SharePoint.Cleaner.Core.Models;

public sealed record GuidanceLink(string Title, string Url);

public sealed record CheckGuidance(
    string Summary,
    string WhyItMatters,
    IReadOnlyList<string> HowToFix,
    IReadOnlyList<GuidanceLink> References);
