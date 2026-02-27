using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ReportGenerator.Views
{
    public partial class HeaderControl : UserControl
    {
        public HeaderControl()
        {
            InitializeComponent();

            // Allow dragging the window by holding the header background or title
            this.MouseLeftButtonDown += HeaderControl_MouseLeftButtonDown;
        }

        private void HeaderControl_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is DependencyObject d && FindAncestor<Button>(d) != null)
            {
                return;
            }

            var wnd = Window.GetWindow(this);
            if (wnd != null && e.ButtonState == MouseButtonState.Pressed)
            {
                if (e.ClickCount == 2)
                {
                    try
                    {
                        wnd.WindowState = wnd.WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
                    }
                    catch { /* ignore */ }
                    return;
                }

                try { wnd.DragMove(); } catch { /* ignore if not draggable */ }
            }
        }

        private static T FindAncestor<T>(DependencyObject start) where T : DependencyObject
        {
            var cur = start;
            while (cur != null)
            {
                if (cur is T t) return t;
                cur = System.Windows.Media.VisualTreeHelper.GetParent(cur);
            }
            return null;
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            var wnd = Window.GetWindow(this);
            wnd?.Close();
        }
    }
}