# SharePoint Cleaner

A free Windows (WPF, .NET 8) tool that scans a SharePoint site for exposed REST APIs, legacy services and information leaks, and explains each finding in plain English with links to Microsoft documentation.

> **Authorized use only.** Scan only sites you own or have explicit written permission to test. The authors accept no liability for misuse. Never publish findings about a site you do not own; report them privately to its owner.

## Features

- **Security Scan**: 10 read-only checks. Select a result to see the evidence, the endpoint tested, what the issue is, why it matters, how to fix it, and related Microsoft Learn links.
- **Windows credentials + anonymous verification**: checks run with your Windows credentials; every check that comes back Exposed is then re-run anonymously. The result shows an access scope: **Public (anonymous)** means anyone can read it, **Authenticated only** means only the scanning account could, and **Unverified** means the anonymous re-run was inconclusive.
- **Explorer**: site and subsite tree (REST `webs`, with a `webinfos` fallback), sortable grid with list and item counts, lists and list items per site.

## Checks

| ID | Risk | What it looks for |
|----|------|-------------------|
| SP-H-001 | High | User information page (`/_layouts/userdisp.aspx`) |
| SP-H-002 | High | `owssvr.dll` list content |
| SP-H-003 | Critical | Lists and items readable via `/_api/web/lists` |
| SP-H-004 | High | User and group enumeration (`siteusers`, `sitegroups`) |
| SP-H-005 | High | Legacy `/_vti_bin/listdata.svc` |
| SP-M-001 | Medium | `/_vti_bin/*.asmx` and `client.svc` service descriptions |
| SP-M-002 | Medium | Subsite enumeration (`webs`, `webinfos`) |
| SP-M-003 | Medium | Anonymous Search REST (`/_api/search/query`) |
| SP-L-001 | Low | FrontPage Server Extensions (`/_vti_inf.html`) |
| SP-L-002 | Low | Version disclosure headers |

All checks only send GET requests and read a few records. User emails are never displayed.

## Run

```powershell
dotnet build SharePoint.Cleaner.sln
dotnet run --project .\SharePoint.Cleaner.App\SharePoint.Cleaner.App.csproj
dotnet test
```

Requires Windows and the .NET 8 SDK.

## Supported SharePoint versions

Built for on-premises **SharePoint 2013, 2016, 2019** (and Subscription Edition), and also works on SharePoint Online where endpoints exist.

- REST responses are read in both the lightweight (`odata=nometadata`) and the verbose (`odata=verbose`, `d.results`) formats. If a farm rejects the lightweight format, the request is retried as verbose (needed for 2013 RTM).
- The version is inferred from the `MicrosoftSharePointTeamServices` header (15.x = 2013, 16.0.0.<10000 = 2016, <14326 = 2019, newer = Subscription Edition) and shown in the status bar. Farms that hide the header show "not disclosed".
- SharePoint 2010 has no `/_api`, so the REST-based checks do not apply there.
- Only tested against live sites whose versions were not confirmed; please report version-specific problems.

## Limitations

- Not a full vulnerability scanner: no CVE checks, and no SharePoint Online (Microsoft 365) specific checks yet.
- Heuristic detection can produce false positives or negatives. Verify findings manually.
- The Explorer tab uses your Windows credentials, so it shows what your account can see.
- Guidance text is this project's own summary. Microsoft links are related documentation, not statements that a feature is deprecated.

## Built with AI

This entire project was built through conversation with an AI coding assistant: the WPF app, the exposure checks, the SharePoint 2013/2016/2019 compatibility, the tests and this README. The human role was directing, testing against real sites and making product decisions. It is a small example of what AI-assisted development can do. Review the code before relying on it.

## Contributing

Issues and pull requests are welcome, especially new checks. Each check lives in `SharePoint.Cleaner.Checks` and its guidance in `CheckGuidanceCatalog.cs`.

## License

MIT, see [LICENSE](LICENSE).
