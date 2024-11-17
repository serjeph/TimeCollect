using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace TimeCollect.Converters
{
    /// <summary>
    /// Converts a boolean value to a Visibility value. This is often used
    /// to control the visibility of UI element based on a boolean property.
    /// </summary>
    public class BoolToVisibilityConverter : IValueConverter
    {
        /// <summary>
        /// Converts a boolean value to a Visibility value.
        /// </summary>
        /// <param name="value">The value produced by the binding source.</param>
        /// <param name="targetType">The type of the binding target property.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>Visibility.Visible if the input is true; otherwise, Visibility.Collapse</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed; // Default to collapse if not a boolean
        }

        /// <summary>
        /// Converts a Visibility value back to a boolean value.
        /// </summary>
        /// <param name="value">The value produced by the binding source.</param>
        /// <param name="targetType">The type of the binding target property.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>True if the Visibility is Visible; otherwise, false.</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility visibility)
            {
                return visibility == Visibility.Visible;
            }
            return false; // Default to false if not Visibility
        }
    }
}
