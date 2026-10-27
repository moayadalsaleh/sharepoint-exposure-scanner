namespace SharePoint.Cleaner.Core.Models;

/// <summary>
/// Explanations, remediation steps and Microsoft Learn references for each check.
/// Every reference URL was verified to resolve; the explanations are this tool's own summary of the issue.
/// </summary>
public static class CheckGuidanceCatalog
{
    private static readonly GuidanceLink AnonymousAccess = new(
        "Manage anonymous access for a web application in SharePoint Server",
        "https://learn.microsoft.com/en-us/sharepoint/administration/manage-anonymous-access-for-a-web-application");

    private static readonly GuidanceLink Hardening = new(
        "Plan security hardening for SharePoint Server",
        "https://learn.microsoft.com/en-us/sharepoint/security-for-sharepoint-server/security-hardening");

    private static readonly GuidanceLink RestOverview = new(
        "Get to know the SharePoint REST service",
        "https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/get-to-know-the-sharepoint-rest-service");

    private static readonly GuidanceLink ListsRest = new(
        "Work with lists and list items with REST",
        "https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest");

    private static readonly GuidanceLink SitePermissions = new(
        "Customize permissions for a SharePoint site",
        "https://learn.microsoft.com/en-us/sharepoint/customize-sharepoint-site-permissions");

    private static readonly GuidanceLink SearchRest = new(
        "SharePoint Search REST API overview (see \"Enabling anonymous Search REST queries\")",
        "https://learn.microsoft.com/en-us/sharepoint/dev/general-development/sharepoint-search-rest-api-overview");

    private static readonly GuidanceLink Deprecated2019 = new(
        "What's deprecated or removed from SharePoint Server 2019",
        "https://learn.microsoft.com/en-us/sharepoint/what-s-new/what-s-deprecated-or-removed-from-sharepoint-server-2019");

    private static readonly GuidanceLink BuildNumbers = new(
        "SharePoint Server updates and build numbers",
        "https://learn.microsoft.com/en-us/officeupdates/sharepoint-updates");

    private static readonly GuidanceLink IisCustomHeaders = new(
        "IIS customHeaders (add or remove HTTP response headers)",
        "https://learn.microsoft.com/en-us/iis/configuration/system.webserver/httpprotocol/customheaders/");

    private static readonly GuidanceLink IisRequestFiltering = new(
        "IIS request filtering (block URLs and file extensions)",
        "https://learn.microsoft.com/en-us/iis/configuration/system.webserver/security/requestfiltering/");

    private static readonly GuidanceLink AspNetVersionHeader = new(
        "HttpRuntimeSection.EnableVersionHeader (X-AspNet-Version header)",
        "https://learn.microsoft.com/en-us/dotnet/api/system.web.configuration.httpruntimesection.enableversionheader");

    private static readonly CheckGuidance Fallback = new(
        "This check did not complete or has no detailed guidance yet.",
        "Review the evidence and endpoint shown for this result.",
        ["Confirm the site is reachable, then run the scan again."],
        [Hardening]);

    private static readonly Dictionary<string, CheckGuidance> Catalog = new(StringComparer.OrdinalIgnoreCase)
    {
        ["SP-H-001"] = new(
            "The user information page (/_layouts/userdisp.aspx) displays profile fields stored in the site's hidden User Information List. " +
            "The scanner requested it with id=1 and the page returned user-field markup.",
            "The id parameter is sequential, so anyone who can open this page can walk id=1,2,3... and harvest display names, e-mail addresses, " +
            "job titles and account names. That list feeds phishing, password spraying and social engineering.",
            [
                "If the site does not need public visitors, disable anonymous access for the web application or the site collection.",
                "If anonymous access is required, block /_layouts/*/userdisp.aspx for anonymous users at the reverse proxy / WAF (or IIS request filtering).",
                "Review who can read the User Information List and remove broad groups that do not need it.",
                "Re-run the scan to confirm the page no longer returns user fields."
            ],
            [AnonymousAccess, SitePermissions, Hardening]),

        ["SP-H-002"] = new(
            "owssvr.dll is a legacy Office/OWS Server handler. Its FileOpen dialog view returned browsable site content (list and library names) without a login.",
            "It reveals the structure of the site (lists, libraries, files) to anyone who can reach the URL, which speeds up targeted attacks " +
            "and can expose content that was never linked from the public pages.",
            [
                "Block /_vti_bin/owssvr.dll at the reverse proxy / WAF or with IIS request filtering unless a legacy Office client truly needs it.",
                "Remove anonymous access from lists and libraries that should not be public.",
                "Re-run the scan to confirm the endpoint no longer lists content."
            ],
            [IisRequestFiltering, AnonymousAccess, Hardening]),

        ["SP-H-003"] = new(
            "The SharePoint REST API (/_api/web/lists) answered without authentication, and list items could be read from the returned lists " +
            "(including document-library entries).",
            "REST gives structured, machine-readable access to list data. Anyone can script a full download of every list the scanner identity can read, " +
            "including fields that the public pages never display (internal notes, e-mail addresses, file paths).",
            [
                "Disable anonymous access to the web application, or limit it to the specific public lists/libraries that are meant to be public.",
                "Break permission inheritance on internal lists and remove \"Anonymous Users\" / \"Everyone\" from them.",
                "Move non-public configuration lists (for example Configuration, EmailData, Control Panel) to a restricted site.",
                "Consider blocking /_api/* for anonymous traffic at the reverse proxy if the public site does not use it.",
                "Re-run the scan; the lists endpoint should answer 401/403 for anonymous requests."
            ],
            [RestOverview, ListsRest, AnonymousAccess, SitePermissions]),

        ["SP-H-004"] = new(
            "The REST endpoints /_api/web/siteusers and /_api/web/sitegroups returned user and group records without authentication.",
            "A full list of accounts, display names and e-mail addresses lets attackers build target lists for phishing and password spraying, " +
            "and group names reveal how permissions are organised.",
            [
                "Disable anonymous access for the web application or site collection if it is not needed.",
                "If anonymous access is required, block /_api/web/siteusers, /_api/web/siteusergroups and /_api/web/sitegroups for anonymous users at the reverse proxy / WAF.",
                "Review who has Read access to the User Information List and remove broad groups.",
                "Re-run the scan to confirm the endpoints return 401/403."
            ],
            [RestOverview, AnonymousAccess, SitePermissions, Hardening]),

        ["SP-H-005"] = new(
            "The legacy WCF Data Services endpoint /_vti_bin/listdata.svc returned its service document, which lists every list as a queryable collection.",
            "ListData.svc is an older, second route to list data. It is often forgotten when /_api is locked down, so it can keep exposing data " +
            "through an unmonitored path.",
            [
                "Block /_vti_bin/listdata.svc at the reverse proxy / WAF or with IIS request filtering if no application depends on it.",
                "Apply the same permission restrictions as for the REST API: remove anonymous access from non-public lists.",
                "Re-run the scan to confirm the service document is no longer returned."
            ],
            [IisRequestFiltering, AnonymousAccess, Hardening]),

        ["SP-M-001"] = new(
            "One or more legacy SOAP/WCF service descriptions under /_vti_bin (lists.asmx, usergroup.asmx, webs.asmx, client.svc) are reachable.",
            "These services can read lists, users and site structure programmatically. Even when they enforce permissions, a reachable description " +
            "tells attackers exactly which operations to try and which SharePoint version is running.",
            [
                "Limit access to /_vti_bin/*.asmx and client.svc to the networks and users that need them.",
                "Block them for anonymous traffic at the reverse proxy / WAF if the public site does not use them.",
                "Follow Microsoft's hardening guidance for securing SharePoint web service files.",
                "Re-run the scan to confirm the service descriptions are no longer returned."
            ],
            [Hardening, Deprecated2019, IisRequestFiltering]),

        ["SP-M-002"] = new(
            "The REST endpoints /_api/web/webs or /_api/web/webinfos listed the subsites under this site without authentication.",
            "The subsite list shows the whole site structure, including internal or unlinked areas (for example /Admin), " +
            "so attackers know where to look next.",
            [
                "Disable anonymous access to the web application, or restrict anonymous access to the site(s) that are meant to be public.",
                "Block /_api/web/webs and /_api/web/webinfos for anonymous users at the reverse proxy / WAF.",
                "Make sure internal subsites do not inherit permissions that include anonymous users.",
                "Re-run the scan to confirm the endpoints return 401/403."
            ],
            [RestOverview, AnonymousAccess, SitePermissions]),

        ["SP-M-003"] = new(
            "The Search REST API (/_api/search/query) answered an anonymous wildcard query and returned results.",
            "Search can surface documents and pages that are not linked from any public page, and the results include metadata such as titles, paths and authors.",
            [
                "Decide whether anonymous Search REST queries are needed. If not, disable them or block /_api/search/* for anonymous users.",
                "If they are needed, restrict which query parameters are exposed to anonymous users (queryparametertemplate.xml).",
                "Exclude non-public lists and libraries from the search index (\"Allow this list/library to appear in search results\" = No).",
                "Re-run the scan to confirm the endpoint no longer returns results."
            ],
            [SearchRest, AnonymousAccess, Hardening]),

        ["SP-L-001"] = new(
            "FrontPage Server Extensions artifact (/_vti_inf.html) is present and returned FrontPage configuration information.",
            "These are remnants of a very old remote-authoring protocol. They disclose the extension version and server configuration " +
            "and widen the attack surface for no benefit.",
            [
                "Block /_vti_inf.html, /_vti_pvt/* and /_vti_bin/_vti_aut/* at the reverse proxy / WAF or with IIS request filtering.",
                "Disable FrontPage remote authoring if it is still enabled for the web application.",
                "Re-run the scan to confirm the files are no longer returned."
            ],
            [Hardening, IisRequestFiltering]),

        ["SP-L-002"] = new(
            "The HTTP response headers disclose the SharePoint build number and/or ASP.NET / framework versions " +
            "(MicrosoftSharePointTeamServices, X-AspNet-Version, X-Powered-By).",
            "The exact build number lets an attacker look up which security updates are missing and pick a matching exploit.",
            [
                "Remove the X-Powered-By header in IIS (customHeaders) and the X-AspNet-Version header (enableVersionHeader=\"false\").",
                "Strip or rewrite the MicrosoftSharePointTeamServices header at the reverse proxy / WAF.",
                "Keep the farm patched to the latest cumulative update regardless of what is disclosed.",
                "Re-run the scan to confirm the headers are gone."
            ],
            [BuildNumbers, IisCustomHeaders, AspNetVersionHeader])
    };

    public static CheckGuidance For(string checkId) =>
        Catalog.TryGetValue(checkId, out var guidance) ? guidance : Fallback;
}
