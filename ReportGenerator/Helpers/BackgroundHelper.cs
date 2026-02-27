using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;

namespace ReportGenerator.Helpers
{
    // Attached property that accepts string/Color/Brush and applies a Brush to the element's Background.
    public static class BackgroundHelper
    {
        public static readonly DependencyProperty BackgroundExProperty =
            DependencyProperty.RegisterAttached(
                "BackgroundEx",
                typeof(object),
                typeof(BackgroundHelper),
                new PropertyMetadata(null, OnBackgroundExChanged));

        public static object GetBackgroundEx(DependencyObject obj) => obj.GetValue(BackgroundExProperty);
        public static void SetBackgroundEx(DependencyObject obj, object value) => obj.SetValue(BackgroundExProperty, value);

        private static void OnBackgroundExChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var brush = BrushHelper.ConvertToBrush(e.NewValue);
            if (brush == null)
            {
                // If conversion fails, do not set an invalid value. Optionally clear.
                return;
            }

            // Prefer direct strongly-typed APIs:
            if (d is Control control)
            {
                control.Background = brush;
                return;
            }

            if (d is Panel panel)
            {
                panel.Background = brush;
                return;
            }

            if (d is Border border)
            {
                border.Background = brush;
                return;
            }

            // Fall back to reflection: try to find a static BackgroundProperty on the type or base types.
            var dp = FindBackgroundDependencyProperty(d.GetType());
            if (dp != null)
            {
                d.SetValue(dp, brush);
            }
        }

        private static DependencyProperty FindBackgroundDependencyProperty(Type type)
        {
            while (type != null)
            {
                var fi = type.GetField("BackgroundProperty", BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
                if (fi != null && typeof(DependencyProperty).IsAssignableFrom(fi.FieldType))
                {
                    return (DependencyProperty)fi.GetValue(null);
                }
                type = type.BaseType;
            }
            return null;
        }
    }
}