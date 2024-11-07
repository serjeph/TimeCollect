using ClosedXML.Excel;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using Npgsql;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace TimeCollect
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<Employee> Employees { get; set; }

        private readonly string _credentialsPath = "credentials.json";


        // add other properties and commands here later...
        public ICommand RunDataCommand { get; set; }
        public ICommand SaveCredentialsCommand { get; set; }
        public string LogMessages { get; set; }

        public MainViewModel()
        {
            Employees = new ObservableCollection<Employee>();
            RunDataCommand = new RelayCommand(RunData);
            SaveCredentialsCommand = new RelayCommand(SaveCredentials);
        }

        private async void RunData()
        {
            try
            {
                //1. Fetch data from Google Sheets for each employee
                var credential = await GetUserCredential();
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "TimeCollect"
                });

                // 2. Process data for each sheet
                var allCleanedData = new List<IList<object>>(); //List to store all cleaned data
                var sheetNames = SheetNames.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                string outputDirectory = @"D:\Documents\TimeCollect";
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                foreach (var sheetName in sheetNames)
                {
                    var sheetData = new List<IList<object>>();

                    foreach (var employee in Employees)
                    {

                        var range = $"{sheetName}!A7:AR39";
                        var request = service.Spreadsheets.Values.Get(employee.SpreadsheetId, range);
                        var response = await request.ExecuteAsync();
                        var values = response.Values;

                        if (values != null && values.Count > 0)
                        {
                            LogMessages += $"Fetched {values.Count} rows for {employee.Nickname} from sheet {sheetName}\n";

                            // Fill blank cells with 0.00
                            for (int i = 0; i < values.Count; i++)
                            {
                                for (int j = 0; j < values[i].Count; j++)
                                {
                                    if (string.IsNullOrEmpty(values[i][j].ToString()))
                                    {
                                        values[i][j] = "0.00";
                                    }
                                }
                            }

                            sheetData.AddRange(CleanData(values));
                        }
                        else
                        {
                            LogMessages += $"No data found for {employee.Nickname} from sheet {sheetName}\n";
                        }

                    }

                    // 3. Process the combined data for the current sheet
                    if (sheetData.Count > 0)
                    {
                        allCleanedData.AddRange(sheetData);
                        ExportToExcel(sheetData, sheetName, Path.Combine(outputDirectory, $"output_{sheetName}.xlsx"));
                    }

                }

                // 4. Process all the cleand data from all sheets combined
                if (allCleanedData.Count > 0)
                {
                    ExportToExcel(allCleanedData, "AllSheets", Path.Combine(outputDirectory, "output_all_sheets.xlsx"));
                    InsertIntoDatabase(allCleanedData);
                }

            }
            catch (Exception ex)
            {
                // 5. Handle exceptions
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task<UserCredential> GetUserCredential()
        {
            //Load credentials from the saved credentials.json file
            using (var stream = new FileStream(_credentialsPath, FileMode.Open, FileAccess.Read))
            {
                return await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.FromStream(stream).Secrets,
                    new[] { SheetsService.Scope.SpreadsheetsReadonly },
                    "user", CancellationToken.None);
            }
        }

        private void SaveCredentials()
        {


            //2. Create a ClientSecret object
            var secrets = new ClientSecrets
            {
                ClientId = ClientId, // from the ViewModel
                ClientSecret = ClientSecret, // from the ViewModel
            };

            //3. Serialize the ClientSecrets object to JSON
            var json = JsonConvert.SerializeObject(secrets, Formatting.Indented);

            //4. Save the JSON to the credentials file
            File.WriteAllText(_credentialsPath, json);

            //5. Provide feedback to the user
            LogMessages += $"Credentials saved to {_credentialsPath}";
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private string _sheetNames;
        public string SheetNames
        {
            get => _sheetNames;
            set
            {
                _sheetNames = value;
                OnPropertyChanged(nameof(SheetNames));
            }
        }

        private string _clientId;
        public string ClientId
        {
            get => _clientId;
            set
            {
                _clientId = value;
                OnPropertyChanged($"{nameof(ClientId)}");
            }
        }

        private string _clientSecret;
        public string ClientSecret
        {
            get => _clientSecret;
            set
            {
                _clientSecret = value;
                OnPropertyChanged(nameof(ClientSecret));
            }
        }

        private IList<IList<object>> CleanData(IList<IList<object>> data)
        {
            // Get the indices of columns to be deleted
            // Example: Delete columns 2, 4 ,6
            // to be edit later
            var columnsToDelete = new List<int>() { 1, 3, 5 };

            // Iterate through each row in the data
            for (int i = 0; i < data.Count; i++)
            {
                //Delete columns in reverse order to avoid index issues
                for (int j = 0; j < columnsToDelete.Count; j++)
                {
                    int columnIndex = columnsToDelete[i];
                    if (columnIndex >= 0 && columnIndex < data[i].Count)
                    {
                        data[i].RemoveAt(columnIndex);
                    }
                }
            }

            return data;

        }

        private void ExportToExcel(IList<IList<object>> data, string sheetName, string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(sheetName);
                int row = 1;
                foreach (var rowData in data)
                {
                    int col = 1;
                    foreach (var cellData in rowData)
                    {
                        // Use pattern matching for type checking
                        switch (cellData)
                        {
                            case string str:
                                worksheet.Cell(row, col).Value = str;
                                break;
                            case int num:
                                worksheet.Cell(row, col).Value = num;
                                break;
                            case float f:
                                worksheet.Cell(row, col).Value = Math.Round(f, 2);
                                break;
                            // ... add more cases for other types as needed
                            default:
                                worksheet.Cell(row, col).Value = cellData.ToString(); // Use SetValue for general cases
                                break;
                        }
                        col++;
                    }
                    row++;
                }
                workbook.SaveAs(filePath);
            }
        }

        private void InsertIntoDatabase(IList<IList<object>> data)
        {
            try
            {
                // 1. Define your connection string
                string connString = "Host=your_host;Database=your_database;Username=your_user;Password=your_password";

                // 2. Establish a connection to the database
                using (var conn = new NpgsqlConnection(connString))
                {
                    conn.Open();

                    // 3. Define the SQL command
                    using (var cmd = new NpgsqlCommand())
                    {
                        cmd.Connection = conn;

                        //TODO: edit this later...
                        cmd.CommandText = "INSERT INTO your_table (column1, column2, ...) VALUES (@value1, @value2, ...)";

                        // 4. Add parameters and values based on table structure
                        // Example:
                        //TODO: edit this later...
                        cmd.Parameters.AddWithValue("column1", NpgsqlTypes.NpgsqlDbType.Varchar);
                        cmd.Parameters.AddWithValue("column2", NpgsqlTypes.NpgsqlDbType.Varchar);

                        // 5. Iterate through the data and insert each row
                        foreach (var rowData in data)
                        {
                            // Assuming rowData[0] corresponds to column 1, rowData[1] to column2, etc...
                            cmd.Parameters["column1"].Value = rowData[0];
                            cmd.Parameters["column2"].Value = rowData[1];
                            // ... set values for other parameters


                            cmd.ExecuteNonQuery();
                        }

                    }
                }
                // 6. Provide feedback
                LogMessages += $"Inserted {data.Count} rows into the database.\n";

            }
            catch (Exception ex)
            {
                // 7. Handle exceptions
                MessageBox.Show($"An error occurred while inserting into the database: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);

            }
        }

    }

}

