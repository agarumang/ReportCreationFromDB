using Microsoft.Extensions.DependencyInjection;
using ReportGenerator.Services;
using ReportGenerator.ViewModels;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;

namespace ReportGenerator
{
    public partial class App : Application
    {
        private ServiceProvider _serviceProvider;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var services = new ServiceCollection();
            ConfigureServices(services);
            _serviceProvider = services.BuildServiceProvider();

            var main = _serviceProvider.GetRequiredService<MainWindow>();

            // Ensure the window opens maximized by default (also respected if set in XAML).
            main.WindowState = WindowState.Maximized;

            main.Show();
        }

        private void ConfigureServices(IServiceCollection services)
        {
            // Services
            services.AddSingleton<IDatabaseService, DatabaseService>();

            // ViewModels
            services.AddTransient<MainViewModel>();

            // Views
            services.AddTransient<MainWindow>();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            _serviceProvider?.Dispose();
        }
    }
}