using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ReportGenerator.ViewModels;
using System.Diagnostics;
using System.Windows.Controls.Primitives;

namespace ReportGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow(MainViewModel vm)
        {
            try
            {
                InitializeComponent();
            }
            catch (System.Windows.Markup.XamlParseException xpe)
            {
                // Log detailed info for diagnostics
                Trace.TraceError("XamlParseException during InitializeComponent: {0}", xpe.ToString());

                // Provide concise actionable message to the user/developer
                var inner = xpe.InnerException;
                string details = inner != null
                    ? $"{inner.GetType().Name}: {inner.Message}"
                    : xpe.Message;

                MessageBox.Show(
                    "The application failed to initialize its UI resources.\n\n" +
                    $"{details}\n\n" +
                    "Common causes: missing image resource, incorrect pack URI, or resource Build Action not set to 'Resource'.\n\n" +
                    "Check Resources/Styles.xaml and the Resources folder for the referenced image and its Build Action.",
                    "Startup Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                // Shutdown gracefully to avoid further exceptions from partially initialized UI
                Application.Current?.Shutdown();
                return;
            }

            DataContext = vm;
            Loaded += async (s, e) => await vm.InitializeAsync();

            // Borderless WPF windows can maximize beyond the visible work area (taskbar),
            // pushing the bottom scrollbar off-screen. Constrain to the work area.
            SourceInitialized += (_, __) =>
            {
                var wa = SystemParameters.WorkArea;
                MaxWidth = wa.Width;
                MaxHeight = wa.Height;
            };
        }

        private void ResultsGrid_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
        {
            // Force a deterministic width so total columns can exceed viewport width,
            // ensuring horizontal scrolling is possible.
            e.Column.MinWidth = 120;
            e.Column.Width = new DataGridLength(140);
        }
    }
}
