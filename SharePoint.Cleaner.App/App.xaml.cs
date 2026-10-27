using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using SharePoint.Cleaner.App.ViewModels;
using SharePoint.Cleaner.Checks.Checks;
using SharePoint.Cleaner.Core.Contracts;
using SharePoint.Cleaner.Infrastructure.Scanning;
using SharePoint.Cleaner.Infrastructure.SharePoint;

namespace SharePoint.Cleaner.App;

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application
{
	private ServiceProvider? _serviceProvider;

	protected override void OnStartup(StartupEventArgs e)
	{
		base.OnStartup(e);

		var services = new ServiceCollection();
		ConfigureServices(services);
		_serviceProvider = services.BuildServiceProvider();

		var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
		mainWindow.Show();
	}

	protected override void OnExit(ExitEventArgs e)
	{
		_serviceProvider?.Dispose();
		base.OnExit(e);
	}

	private static void ConfigureServices(IServiceCollection services)
	{
		services.AddSingleton<IScanCheck, UserInformationExposureCheck>();
		services.AddSingleton<IScanCheck, OwssvrExposureCheck>();
		services.AddSingleton<IScanCheck, ListItemsExposureCheck>();
		services.AddSingleton<IScanCheck, WebServicesExposureCheck>();
		services.AddSingleton<IScanCheck, FrontPageExtensionsCheck>();
		services.AddSingleton<IScanCheck, UserEnumerationExposureCheck>();
		services.AddSingleton<IScanCheck, ListDataServiceExposureCheck>();
		services.AddSingleton<IScanCheck, SiteStructureExposureCheck>();
		services.AddSingleton<IScanCheck, SearchRestExposureCheck>();
		services.AddSingleton<IScanCheck, VersionDisclosureCheck>();

		services.AddSingleton<IScanEngine, DefaultScanEngine>();
		services.AddSingleton<ISharePointExplorer, SharePointExplorer>();

		services.AddTransient<MainViewModel>();
		services.AddTransient<MainWindow>();
	}
}

