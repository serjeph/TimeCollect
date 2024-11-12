using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace TimeCollect.Helpers
{
    public static class DateTimeHelper
    {
        public static string GetWeekType(int year, int month, int day, string filename = "weekType.json")
        {
            try
            {
                string json = File.ReadAllText(filename);
                Dictionary<string, string> weekTypes = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

                DateTime dateObject = new DateTime(year, month, day); //To create DateTime object directly

                CultureInfo culture = CultureInfo.CurrentCulture;
                Calendar calendar = culture.Calendar;
                CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
                DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;
                int weekNumber = calendar.GetWeekOfYear(dateObject, weekRule, firstDayOfWeek);

                return weekTypes.TryGetValue(weekNumber.ToString(), out string weekType) ? weekType : "not found";


            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("JSON file not found/.");
                return null;
            }
            catch (JsonReaderException)
            {
                Console.WriteLine("Invalid JSON");
                return null;
            }
        }
    }
}
