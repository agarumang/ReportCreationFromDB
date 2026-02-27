using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace ReportGenerator.Converters
{
    public class BoolToVisibilityConverter : IValueConverter
    {
        // If parameter == "Invert" the boolean is inverted before conversion.
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool flag = false;
            if (value is bool b) flag = b;
            else if (value is bool?) flag = ((bool?)value) ?? false;

            if (parameter is string p && string.Equals(p, "Invert", StringComparison.OrdinalIgnoreCase))
            {
                flag = !flag;
            }

            return flag ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility v)
            {
                bool result = v == Visibility.Visible;
                if (parameter is string p && string.Equals(p, "Invert", StringComparison.OrdinalIgnoreCase))
                {
                    result = !result;
                }
                return result;
            }
            return false;
        }
    }
}