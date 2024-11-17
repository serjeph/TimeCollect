using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows;
using TimeCollect.Models;

namespace TimeCollect.Helpers
{
    public class WeekTypeHelper
    {
        public static ObservableCollection<WeekType> SetWeekTypes(DateTime yearStartDate)
        {
            CultureInfo culture = CultureInfo.CurrentCulture;
            Calendar calendar = culture.Calendar;
            DayOfWeek dayOfWeek = calendar.GetDayOfWeek(yearStartDate);

            // Ensure yearStartDate is a Sunday
            if (dayOfWeek != DayOfWeek.Sunday)
            {
                MessageBox.Show($"Date must be a Sunday. Date set: {dayOfWeek}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

            // Calculate the number of weeks in the year
            int numWeeks = (yearStartDate.AddDays(366).Year == yearStartDate.AddYears(1).Year) ? 52 : 51;

            var data = new ObservableCollection<WeekType>();

            for (int i = 0; i < numWeeks; i++)
            {
                DateTime startDate = yearStartDate.AddDays(i * 7);
                DateTime endDate = startDate.AddDays(6);

                data.Add(new WeekType()
                {
                    DateStart = startDate,
                    DateEnd = endDate,
                    WeekName = SetWeekName(startDate),
                    WeekTypeName = SetWeekTypeName(startDate)
                });
            }

            return data;
        }

        public static ObservableCollection<WeekType> GetWeekTypes(DateTime yearStartDate, string filePath = "weekNames.json")
        {

            // Ensure yearStartDate is Sunday
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string jsonFilePath = Path.Combine(appDataPath, "TimeCollect", filePath);
            var weekData = new ObservableCollection<WeekType>();

            if (File.Exists(jsonFilePath))
            {

                var jsonData = File.ReadAllText(jsonFilePath);
                var data = JsonConvert.DeserializeObject<ObservableCollection<WeekType>>(jsonData);

                foreach (var item in data)
                {
                    weekData.Add(new WeekType()
                    {
                        DateStart = item.DateStart,
                        DateEnd = item.DateEnd,
                        WeekName = item.WeekName,
                        WeekTypeName = item.WeekTypeName
                    });
                }

                return weekData;
            }
            else
            {
                return SetWeekTypes(yearStartDate);

            }
        }


        private static string SetWeekTypeName(DateTime dateStart)
        {
            string weekTypeName = "-";

            DateTime weekStart = dateStart;
            DateTime weekEnd = dateStart.AddDays(6);
            DateTime lastWeekStart = dateStart.AddDays(-7);
            DateTime lastWeekEnd = dateStart.AddDays(-1);
            DateTime lastLastWeekEnd = dateStart.AddDays(-13);
            DateTime lastLastWeekStart = dateStart.AddDays(-14);
            DateTime lastTwoWeekEnd = dateStart.AddDays(-20);
            DateTime lastTwoWeekStart = dateStart.AddDays(-21);
            DateTime lastThreeWeekEnd = dateStart.AddDays(-27);

            // first week of the year
            if (weekStart.Year != weekEnd.Year && weekEnd.Month == 1)
            {
                weekTypeName = "1A";
            }


            if (weekStart.Year == weekEnd.Year)
            {
                // if in January
                if (weekStart.Month == 1)
                {
                    if (weekStart.Month != lastWeekStart.Month)
                    {
                        weekTypeName = $"{weekStart.Month}B";
                    }
                    else if (weekStart.Month != lastLastWeekStart.Month && weekStart.Month == lastWeekEnd.Month)
                    {
                        weekTypeName = $"{weekStart.Month}C";
                    }
                    else if (weekStart.Month != lastTwoWeekStart.Month && weekStart.Month == lastLastWeekStart.Month)
                    {
                        weekTypeName = $"{weekStart.Month}D";
                    }
                    else
                    {
                        weekTypeName = $"{weekStart.Month}E";
                    }
                }


                if (weekStart.Month != weekEnd.Month)
                {
                    weekTypeName = $"{weekStart.Month}to{weekEnd.Month}";
                }

                if (weekStart.Month == weekEnd.Month && weekStart.Month > 1)
                {
                    if (weekStart.Month != lastWeekStart.Month)
                    {
                        weekTypeName = $"{weekStart.Month}A";
                    }
                    else if (weekStart.Month != lastLastWeekEnd.Month && weekStart.Month == lastWeekEnd.Month)
                    {
                        weekTypeName = $"{weekStart.Month}B";
                    }
                    else if (weekStart.Month != lastTwoWeekEnd.Month && weekStart.Month == lastLastWeekEnd.Month)
                    {
                        weekTypeName = $"{weekStart.Month}C";
                    }
                    else if (weekStart.Month != lastThreeWeekEnd.Month && weekStart.Month == lastTwoWeekEnd.Month)
                    {
                        weekTypeName = $"{weekStart.Month}D";
                    }
                    else
                    {
                        weekTypeName = $"{weekStart.Month}E";
                    }
                }

            }

            return weekTypeName;

        }

        private static string SetWeekName(DateTime dateStart)
        {
            DateTime weekStart = dateStart;
            DateTime weekEnd = dateStart.AddDays(6);

            if (weekStart.Month != weekEnd.Month)
            {
                return $"{weekStart:MMM}{weekStart.Day:d}-{weekEnd:MMM}{weekEnd.Day:d}";
            }
            else
            {
                return $"{weekStart:MMM}{weekStart.Day:d}-{weekEnd.Day:d}";
            }

        }




    }
}
