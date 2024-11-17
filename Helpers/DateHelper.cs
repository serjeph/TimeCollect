using System;

namespace TimeCollect.Helpers
{
    public static class DateHelper
    {
        public static string IsActual(int year, int month, int day)
        {
            try
            {
                DateTime date = new DateTime(year, month, day);
                return date < DateTime.Today ? "Y" : "N";
            }
            catch (ArgumentOutOfRangeException)
            {
                return "Invalid Date";
            }
        }
    }
}
