using System;
using System.Globalization;
using System.Windows.Data;

namespace TimeCollect.Converters
{

    /// <summary>
    /// Converts a boolean value to another boolean value. This converter
    /// is typically used to control the enabled/disabled state of a UI
    /// element based on a boolean property in the date source.
    /// </summary>
    public class BoolToIsRunEnabledConverter : IValueConverter
    {
        /// <summary>
        /// Converts a boolean value to a boolean value.
        /// </summary>
        /// <param name="value"> The value produced by the binding source.</param>
        /// <param name="targetType">The type of the binding target property.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>THe boolean value if the input is a boolean; otherwise, false.</returns>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue;
            }
            return false;
        }

        /// <summary>
        /// Converts a boolean value back to a boolean value.
        /// </summary>
        /// <param name="value"> The value produced by the binding source.</param>
        /// <param name="targetType">The type of the binding target property.</param>
        /// <param name="parameter">The converter parameter to use.</param>
        /// <param name="culture">The culture to use in the converter.</param>
        /// <returns>THe boolean value if the input is a boolean; otherwise, false.</returns>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue;
            }
            return false;
        }
    }
}
