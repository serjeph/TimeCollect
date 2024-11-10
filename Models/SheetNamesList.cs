using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;

namespace TimeCollect.Models
{
    public class SheetNamesList
    {
        public ObservableCollection<string> SheetNames { get; set; }

        public SheetNamesList()
        {
            SheetNames = new ObservableCollection<string>();
        }

        public void AddSheetName(string sheetName)
        {
            SheetNames.Add(sheetName);
        }

        public void RemoveSheetName(string sheetName)
        {
            SheetNames.Remove(sheetName);
        }

        public bool ContainsSheetName(string sheetName)
        {
            return SheetNames.Contains(sheetName);
        }

        public void SaveToJson(string fileName = "sheetNames.json")
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string filePath = Path.Combine(appDataPath, "TimeCollect", fileName);

            try
            {
                string jsonString = JsonSerializer.Serialize(this);
                File.WriteAllText(filePath, jsonString);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving sheet names to JSON: {ex.Message}");
            }
        }

        public static SheetNamesList LoadFromJson(string fileName = "sheetNames.json")
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string filePath = Path.Combine(appDataPath, "TimeCollect", fileName);
            try
            {

                if (!File.Exists(filePath))
                {
                    File.WriteAllText(filePath, "{}");
                }

                string jsonString = File.ReadAllText(filePath);
                return JsonSerializer.Deserialize<SheetNamesList>(jsonString);

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading sheet names from JSON: {ex.Message}");
                return new SheetNamesList();
            }


        }

        public string SheetNamesAsString
        {
            get => string.Join(",", SheetNames);
        }

    }
}
