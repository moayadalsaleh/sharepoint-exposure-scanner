using System.ComponentModel;

namespace SharePoint.Cleaner.Core.Models;

public sealed class SharePointSiteNode : INotifyPropertyChanged
{
    private int? _ownListCount;
    private int? _ownItemCount;
    private int? _totalItemCount;
    private bool _countFailed;

    public required string Title { get; init; }
    public required string Url { get; init; }
    public required IReadOnlyList<SharePointSiteNode> Children { get; init; }

    // Nesting level (0 = root) and number of subsites below this site; set when the flat grid is built.
    public int Depth { get; set; }
    public int SubsiteCount { get; set; }
    public string Path => new Uri(Url).AbsolutePath;

    public event PropertyChangedEventHandler? PropertyChanged;

    public int? OwnListCount
    {
        get => _ownListCount;
        set { _ownListCount = value; Raise(nameof(OwnListCount)); Raise(nameof(CountText)); }
    }

    // Items in this site's own (non-hidden) lists.
    public int? OwnItemCount
    {
        get => _ownItemCount;
        set { _ownItemCount = value; Raise(nameof(OwnItemCount)); Raise(nameof(CountText)); }
    }

    // Items in this site plus all of its subsites.
    public int? TotalItemCount
    {
        get => _totalItemCount;
        set { _totalItemCount = value; Raise(nameof(TotalItemCount)); Raise(nameof(CountText)); }
    }

    public bool CountFailed
    {
        get => _countFailed;
        set { _countFailed = value; Raise(nameof(CountFailed)); Raise(nameof(CountText)); }
    }

    public string CountText
    {
        get
        {
            if (CountFailed)
            {
                return "n/a";
            }

            if (OwnItemCount is null)
            {
                return "counting...";
            }

            var own = $"{OwnItemCount:N0} items in {OwnListCount} lists";
            return Children.Count > 0 && TotalItemCount is not null
                ? $"{own} | total with subsites: {TotalItemCount:N0}"
                : own;
        }
    }

    private void Raise(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
}
