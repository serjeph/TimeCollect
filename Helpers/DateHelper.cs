using System;
using System.Globalization;

namespace TimeCollect.Helpers
{
    public static class DateHelper
    {
        public static string IsActual(string dateString)
        {
            // Use the correct format string for yyyy-M-d
            if (DateTime.TryParseExact(
                dateString,
                "yyyy-M-d",
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime date))
            {
                return date < DateTime.Today ? "Y" : "N";
            }
            else
            {
                return "Invalid Date";
            }
        }
    }
}
