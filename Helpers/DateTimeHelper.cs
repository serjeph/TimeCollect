using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace TimeCollect.Helpers
{
    public static class DateTimeHelper
    {
        public static string GetWeekType(string dateString, string filename = "weekType.json")
        {
            string weekTypeFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", filename);
            try
            {
                string json = File.ReadAllText(weekTypeFilePath);
                Dictionary<string, string> weekTypes = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);

                DateTime dateObject = DateTime.ParseExact(dateString, "yyyy-M-d", null);

                CultureInfo culture = CultureInfo.CurrentCulture;
                Calendar calendar = culture.Calendar;
                CalendarWeekRule weekRule = culture.DateTimeFormat.CalendarWeekRule;
                DayOfWeek firstDayOfWeek = culture.DateTimeFormat.FirstDayOfWeek;
                int weekNumber = calendar.GetWeekOfYear(dateObject, weekRule, firstDayOfWeek);

                if (weekTypes.ContainsKey(weekNumber.ToString()))
                {
                    return weekTypes[weekNumber.ToString()];
                }
                else
                {
                    return null;
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("JSON file missing.");
                return null;
            }
            catch (JsonReaderException)
            {
                Console.WriteLine($"Invalid JSON");
                return null;
            }
        }
    }
}
