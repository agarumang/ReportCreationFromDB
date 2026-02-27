using System;
using System.Windows;
using System.Windows.Media;

namespace ReportGenerator.Helpers
{
    public static class BrushHelper
    {
        // Convert an input (Brush, Color, or string) into a Brush instance, or null if conversion fails.
        public static Brush ConvertToBrush(object value)
        {
            if (value == null) return null;

            if (value is Brush brush) return brush;

            if (value is Color color) return new SolidColorBrush(color);

            if (value is string s)
            {
                s = s.Trim();
                if (string.IsNullOrEmpty(s)) return null;

                try
                {
                    // BrushConverter handles many string forms (e.g. "#FF9D2626", "Red")
                    var converted = (Brush)new BrushConverter().ConvertFromString(s);
                    return converted;
                }
                catch
                {
                    // Fallback: try ColorConverter -> SolidColorBrush
                    try
                    {
                        var col = (Color)ColorConverter.ConvertFromString(s);
                        return new SolidColorBrush(col);
                    }
                    catch
                    {
                        return null;
                    }
                }
            }

            return null;
        }
    }
}