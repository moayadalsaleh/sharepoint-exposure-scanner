using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Input;
using Microsoft.Win32;
using SharePoint.Cleaner.App.Commands;
using SharePoint.Cleaner.App.Models;
using SharePoint.Cleaner.Core.Contracts;
using SharePoint.Cleaner.Core.Models;

namespace SharePoint.Cleaner.App.ViewModels;

public sealed class MainViewModel : INotifyPropertyChanged
{
    private readonly IScanEngine _scanEngine;
    private readonly ISharePointExplorer _sharePointExplorer;
    private string _targetUrl = "https://contoso.sharepoint.com";
    private string _statusText = "Ready";
    private bool _isScanning;
    private string _selectedListTitle = string.Empty;
    private int _itemsTop = 500;
    private int _itemsMinId;

    public MainViewModel(IScanEngine scanEngine, ISharePointExplorer sharePointExplorer)
    {
        _scanEngine = scanEngine;
        _sharePointExplorer = sharePointExplorer;
        ScanCommand = new AsyncRelayCommand(ScanAsync, () => !IsScanning);
        LoadListsCommand = new AsyncRelayCommand(LoadListsAsync, () => !IsScanning);
        LoadListItemsCommand = new AsyncRelayCommand(LoadListItemsAsync, () => !IsScanning);
        LoadSiteTreeCommand = new AsyncRelayCommand(LoadSiteTreeAsync, () => !IsScanning);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public string TargetUrl
    {
        get => _targetUrl;
        set
        {
            if (_targetUrl == value)
            {
                return;
            }

            _targetUrl = value;
            OnPropertyChanged();
            if (SelectedSiteNode is null)
            {
                OnPropertyChanged(nameof(ExplorerSiteUrl));
            }
        }
    }

    public string StatusText
    {
        get => _statusText;
        private set
        {
            if (_statusText == value)
            {
                return;
            }

            _statusText = value;
            OnPropertyChanged();
        }
    }

    public bool IsScanning
    {
        get => _isScanning;
        private set
        {
            if (_isScanning == value)
            {
                return;
            }

            _isScanning = value;
            OnPropertyChanged();
            if (ScanCommand is AsyncRelayCommand cmd)
            {
                cmd.RaiseCanExecuteChanged();
            }
            if (LoadListsCommand is AsyncRelayCommand loadLists)
            {
                loadLists.RaiseCanExecuteChanged();
            }
            if (LoadListItemsCommand is AsyncRelayCommand loadItems)
            {
                loadItems.RaiseCanExecuteChanged();
            }
            if (LoadSiteTreeCommand is AsyncRelayCommand loadTree)
            {
                loadTree.RaiseCanExecuteChanged();
            }
        }
    }

    public ICommand ScanCommand { get; }
    public ICommand LoadListsCommand { get; }
    public ICommand LoadListItemsCommand { get; }
    public ICommand LoadSiteTreeCommand { get; }

    public ObservableCollection<ScanCheckResult> Results { get; } = [];

    private ScanCheckResult? _selectedResult;
    public ScanCheckResult? SelectedResult
    {
        get => _selectedResult;
        set
        {
            if (ReferenceEquals(_selectedResult, value))
            {
                return;
            }

            _selectedResult = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(SelectedGuidance));
            OnPropertyChanged(nameof(HasSelectedResult));
        }
    }

    public CheckGuidance? SelectedGuidance => _selectedResult is null ? null : CheckGuidanceCatalog.For(_selectedResult.CheckId);
    public bool HasSelectedResult => _selectedResult is not null;
    public ObservableCollection<SharePointListSummary> Lists { get; } = [];
    public ObservableCollection<SharePointListItemRecord> ListItems { get; } = [];
    public ObservableCollection<SharePointSiteNode> SiteNodes { get; } = [];
    public ObservableCollection<SharePointSiteNode> SiteRows { get; } = [];
    public ObservableCollection<ItemFieldRow> ItemFields { get; } = [];

    private SharePointSiteNode? _selectedSiteNode;
    public SharePointSiteNode? SelectedSiteNode
    {
        get => _selectedSiteNode;
        set
        {
            if (_selectedSiteNode == value)
            {
                return;
            }

            _selectedSiteNode = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(ExplorerSiteUrl));

            if (value is not null)
            {
                _ = LoadListsAsync();
            }
        }
    }

    private SharePointListSummary? _selectedList;
    public SharePointListSummary? SelectedList
    {
        get => _selectedList;
        set
        {
            if (_selectedList == value)
            {
                return;
            }

            _selectedList = value;
            OnPropertyChanged();
            SelectedListTitle = value?.Title ?? string.Empty;

            if (value is not null)
            {
                _ = LoadListItemsAsync();
            }
        }
    }

    private bool _expandAllSites = true;
    public bool ExpandAllSites
    {
        get => _expandAllSites;
        set
        {
            if (_expandAllSites == value)
            {
                return;
            }

            _expandAllSites = value;
            OnPropertyChanged();
        }
    }

    private CancellationTokenSource? _listsCts;
    private CancellationTokenSource? _itemsCts;

    public string ExplorerSiteUrl => SelectedSiteNode?.Url ?? TargetUrl;

    private SharePointListItemRecord? _selectedListItem;
    public SharePointListItemRecord? SelectedListItem
    {
        get => _selectedListItem;
        set
        {
            if (_selectedListItem == value) return;
            _selectedListItem = value;
            OnPropertyChanged();
            ItemFields.Clear();
            if (value is not null)
            {
                foreach (var kv in value.AllFields)
                    ItemFields.Add(new ItemFieldRow { Field = kv.Key, Value = kv.Value });
            }
        }
    }

    public string SelectedListTitle
    {
        get => _selectedListTitle;
        set
        {
            if (_selectedListTitle == value)
            {
                return;
            }

            _selectedListTitle = value;
            OnPropertyChanged();
        }
    }

    public int ItemsTop
    {
        get => _itemsTop;
        set
        {
            if (_itemsTop == value)
            {
                return;
            }

            _itemsTop = value;
            OnPropertyChanged();
        }
    }

    public int ItemsMinId
    {
        get => _itemsMinId;
        set
        {
            if (_itemsMinId == value)
            {
                return;
            }

            _itemsMinId = value;
            OnPropertyChanged();
        }
    }

    private async Task ScanAsync()
    {
        if (!Uri.TryCreate(TargetUrl.Trim(), UriKind.Absolute, out var targetUri))
        {
            MessageBox.Show("Please enter a valid SharePoint URL.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            IsScanning = true;
            StatusText = "Scanning...";
            Results.Clear();
            SelectedResult = null;

            var report = await _scanEngine.RunAsync(new ScanRequest
            {
                TargetUri = targetUri,
                CheckTimeout = TimeSpan.FromSeconds(10)
            }, CancellationToken.None);

            foreach (var result in report.Results)
            {
                Results.Add(result);
            }

            SelectedResult = Results.FirstOrDefault(r => r.Outcome == CheckOutcome.Exposed) ?? Results.FirstOrDefault();

            StatusText = $"Completed | Version: {report.DetectedVersion ?? "not disclosed"} | Risk: {report.OverallRisk} | Exposed: {report.ExposedCount} | Secured: {report.SecuredCount} | Errors: {report.ErrorCount}";
        }
        catch (Exception ex)
        {
            StatusText = "Scan failed.";
            MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsScanning = false;
        }
    }

    private async Task LoadListsAsync()
    {
        if (!TryGetTargetUri(out var targetUri))
        {
            return;
        }

        _listsCts?.Cancel();
        _itemsCts?.Cancel();
        var cts = _listsCts = new CancellationTokenSource();
        var siteUri = GetExplorerUri(targetUri);

        Lists.Clear();
        ListItems.Clear();
        ItemFields.Clear();
        StatusText = $"Loading lists for {siteUri}...";

        try
        {
            var lists = await _sharePointExplorer.GetListsAsync(siteUri, cts.Token);
            if (cts.IsCancellationRequested)
            {
                return;
            }

            foreach (var list in lists)
            {
                Lists.Add(list);
            }

            StatusText = $"{siteUri} - {Lists.Count} lists. Select a list to load its items.";
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            if (!cts.IsCancellationRequested)
            {
                StatusText = $"Failed to load lists for {siteUri}: {ex.Message}";
            }
        }
    }

    private async Task LoadListItemsAsync()
    {
        if (!TryGetTargetUri(out var targetUri))
        {
            return;
        }

        var listTitle = SelectedListTitle;
        if (string.IsNullOrWhiteSpace(listTitle))
        {
            StatusText = "Select a list first.";
            return;
        }

        _itemsCts?.Cancel();
        var cts = _itemsCts = new CancellationTokenSource();

        ListItems.Clear();
        ItemFields.Clear();
        StatusText = $"Loading items from {listTitle}...";

        try
        {
            var items = await _sharePointExplorer.GetListItemsAsync(
                GetExplorerUri(targetUri),
                listTitle,
                ItemsTop,
                ItemsMinId,
                cts.Token);
            if (cts.IsCancellationRequested)
            {
                return;
            }

            foreach (var item in items)
            {
                ListItems.Add(item);
            }

            StatusText = $"Loaded {ListItems.Count} items from {listTitle}.";
        }
        catch (OperationCanceledException)
        {
        }
        catch (Exception ex)
        {
            if (!cts.IsCancellationRequested)
            {
                StatusText = $"Failed to load items from {listTitle}: {ex.Message}";
            }
        }
    }

    private async Task LoadSiteTreeAsync()
    {
        if (!TryGetTargetUri(out var targetUri))
        {
            return;
        }

        try
        {
            IsScanning = true;
            StatusText = "Loading site tree...";
            _countCts?.Cancel();
            SelectedSiteNode = null;
            SelectedList = null;
            Lists.Clear();
            ListItems.Clear();
            ItemFields.Clear();
            SiteNodes.Clear();
            SiteRows.Clear();

            var nodes = await _sharePointExplorer.GetSiteTreeAsync(targetUri, CancellationToken.None);
            foreach (var node in nodes)
            {
                SiteNodes.Add(node);
                AddRows(node, 0);
            }

            StatusText = $"Site tree loaded ({nodes.Sum(CountSites)} sites). Counting list items...";
            IsScanning = false;
            await CountItemsAsync(nodes);
        }
        catch (Exception ex)
        {
            StatusText = "Failed to load site tree.";
            MessageBox.Show(ex.Message, "Explorer", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsScanning = false;
        }
    }

    private CancellationTokenSource? _countCts;

    private async Task CountItemsAsync(IReadOnlyList<SharePointSiteNode> roots)
    {
        _countCts?.Cancel();
        var cts = _countCts = new CancellationTokenSource();
        var all = new List<SharePointSiteNode>();
        void Flatten(SharePointSiteNode n) { all.Add(n); foreach (var c in n.Children) Flatten(c); }
        foreach (var r in roots) Flatten(r);

        using var throttle = new SemaphoreSlim(6);
        var done = 0;

        async Task CountOne(SharePointSiteNode node)
        {
            try
            {
                await throttle.WaitAsync(cts.Token);
                try
                {
                    var lists = await _sharePointExplorer.GetListsAsync(new Uri(node.Url), cts.Token);
                    node.OwnListCount = lists.Count;
                    node.OwnItemCount = lists.Sum(l => l.ItemCount);
                }
                finally
                {
                    throttle.Release();
                }
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch
            {
                node.CountFailed = true;
            }

            foreach (var r in roots) RecomputeTotals(r);
            if (!cts.IsCancellationRequested)
            {
                StatusText = $"Counting list items... {++done}/{all.Count} sites";
            }
        }

        await Task.WhenAll(all.Select(CountOne));
        if (cts.IsCancellationRequested)
        {
            return;
        }

        var grand = roots.Sum(r => r.TotalItemCount ?? 0);
        StatusText = $"Done: {all.Count} sites, {grand:N0} list items in total. Select a site to see its lists.";
    }

    private static int RecomputeTotals(SharePointSiteNode node)
    {
        var total = node.OwnItemCount ?? 0;
        foreach (var child in node.Children)
        {
            total += RecomputeTotals(child);
        }

        node.TotalItemCount = total;
        return total;
    }

    private int AddRows(SharePointSiteNode node, int depth)
    {
        node.Depth = depth;
        SiteRows.Add(node);
        var descendants = 0;
        foreach (var child in node.Children)
        {
            descendants += 1 + AddRows(child, depth + 1);
        }

        node.SubsiteCount = descendants;
        return descendants;
    }

    private static int CountSites(SharePointSiteNode node) => 1 + node.Children.Sum(CountSites);

    private bool TryGetTargetUri(out Uri targetUri)
    {
        if (!Uri.TryCreate(TargetUrl.Trim(), UriKind.Absolute, out targetUri!))
        {
            MessageBox.Show("Please enter a valid SharePoint URL.", "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        return true;
    }

    private Uri GetExplorerUri(Uri fallbackUri)
    {
        return SelectedSiteNode is not null && Uri.TryCreate(SelectedSiteNode.Url, UriKind.Absolute, out var selectedUri)
            ? selectedUri
            : fallbackUri;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
