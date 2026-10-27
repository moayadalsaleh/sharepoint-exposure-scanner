using System.Windows;
using SharePoint.Cleaner.Core.Models;
using SharePoint.Cleaner.App.ViewModels;

namespace SharePoint.Cleaner.App;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }

    private void SiteTree_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (DataContext is MainViewModel viewModel)
        {
            viewModel.SelectedSiteNode = e.NewValue as SharePointSiteNode;
        }
    }

    private void Hyperlink_RequestNavigate(object sender, System.Windows.Navigation.RequestNavigateEventArgs e)
    {
        if (Uri.TryCreate(e.Uri?.ToString(), UriKind.Absolute, out var uri) && uri.Scheme == Uri.UriSchemeHttps)
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
        }

        e.Handled = true;
    }
}