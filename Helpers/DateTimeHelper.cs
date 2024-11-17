using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using TimeCollect.Models;

namespace TimeCollect.Helpers
{
    public static class DateTimeHelper
    {
        public static string GetWeekType(int year, int month, int day, string filename = "weekNames.json")
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string weekNamesFilePath = Path.Combine(appDataPath, "TimeCollect", filename);
            try
            {
                CultureInfo culture = CultureInfo.CurrentCulture;
                Calendar calendar = culture.Calendar;
                DateTime date = new DateTime(year, month, day);

                var json = File.ReadAllText(weekNamesFilePath);
                var weekTypes = JsonConvert.DeserializeObject<ObservableCollection<WeekType>>(json);

                var weekType = weekTypes.FirstOrDefault(w => date >= w.DateStart && date <= w.DateEnd);

                return weekType?.WeekTypeName ?? "-";

            }
            catch (FileNotFoundException)
            {
                Console.WriteLine("JSON file not found.");
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
