using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Data;

namespace TimeCollect.Converters
{
    public class SheetNamesConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            try
            {
                if (value is ObservableCollection<string> sheetNames)
                {
                    return string.Join(",", sheetNames);
                }

            }
            catch (Exception ex)
            {

                Console.WriteLine($"Error in SheetNamesConverter.Convert: {ex.Message}");
            }

            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {

            try
            {
                if (value is string strValue)
                {
                    return new ObservableCollection<string>(strValue.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries));
                }

            }
            catch (Exception ex)
            {

                Console.WriteLine($"Error in SheetNamesConverter.ConvertBack: {ex.Message}");
            }
            return new ObservableCollection<string>();
        }

    }
}
