using Google.Apis.Auth.OAuth2;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using TimeCollect.Helpers;
using TimeCollect.Services;

namespace TimeCollect
{
    public class MainViewModel : INotifyPropertyChanged
    {

        private readonly string _credentialsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", "credentials.json");
        private readonly GoogleSheetsService _googleSheetsService;
        private readonly DatabaseService _databaseService;

        public ObservableCollection<Employee> Employees { get; set; }
        public ICommand RunDataCommand { get; set; }
        public ICommand SaveCredentialsCommand { get; set; }

        private string _clientId;
        public string ClientId
        {
            get => _clientId;
            set
            {
                _clientId = value;
                OnPropertyChanged(nameof(ClientId));
            }
        }

        private string _clientSecret;
        public string ClientSecret
        {
            get => _clientSecret;
            set
            {
                _clientSecret = value;
                OnPropertyChanged(nameof(_clientSecret));
            }
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

        private string _outputDirectory = @"D:\Documents\TimeCollect";
        public string OutputDirectory
        {
            get => _outputDirectory;
            set
            {
                _outputDirectory = value;
                OnPropertyChanged(nameof(OutputDirectory));
            }
        }

        private bool _isRunEnabled = true;
        public bool IsRunEnabled
        {
            get => _isRunEnabled;
            set
            {
                _isRunEnabled = value;
                OnPropertyChanged(nameof(IsRunEnabled));
            }
        }

        private Visibility _isLoading = Visibility.Collapsed;
        public Visibility IsLoading
        {
            get => _isLoading;
            set
            {
                _isLoading = value;
                OnPropertyChanged(nameof(IsLoading));
            }
        }

        public string LogMessages { get; set; }

        public MainViewModel()
        {

            Employees = new ObservableCollection<Employee>();
            RunDataCommand = new RelayCommand(async () => await RunData());
            SaveCredentialsCommand = new RelayCommand(SaveCredentials);

            _googleSheetsService = new GoogleSheetsService(_credentialsPath);

            string connString = "Host=localhost;Database=timecollect;Username=postgres;Password=admin;Port=5433";
            _databaseService = new DatabaseService(connString);
        }

        private async Task RunData()
        {
            try
            {
                IsRunEnabled = false;
                IsLoading = Visibility.Visible;

                var service = await _googleSheetsService.CreateSheetsService();

                var allCleanedData = new List<IList<object>>();
                var sheetNames = SheetNames.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                string outputDirectory = OutputDirectory;
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

                            for (int i = 0; i < values.Count; i++)
                            {
                                for (int j = 0; j < values[i].Count; j++)
                                {
                                    if (string.IsNullOrEmpty(values[i][j]?.ToString()))
                                    {
                                        // to convert blank cells to 0.00
                                        values[i][j] = "0.00";
                                    }
                                }
                            }

                            sheetData.AddRange(CleanData(values));

                        }
                        else
                        {
                            LogMessages += $"No data found for {employee.Nickname} from sheet {sheetName}\n"; ;
                        }
                    }

                    if (sheetData.Count > 0)
                    {
                        allCleanedData.AddRange(sheetData);
                        ExcelHelper.ExportToExcel(sheetData, sheetName, Path.Combine(outputDirectory, $"output_{sheetName}.xlsx"));
                    }
                }

                if (allCleanedData.Count > 0)
                {
                    ExcelHelper.ExportToExcel(allCleanedData, "AllSheets", Path.Combine(outputDirectory, "output_all_sheets.xlsx"));
                    _databaseService.InsertData(allCleanedData);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsRunEnabled = true;
                IsLoading = Visibility.Collapsed;
            }
        }

        private IList<IList<object>> CleanData(IList<IList<object>> data)
        {
            var columnsToDelete = new List<int>() { 3, 4, 5, 6, 7, 8, 9, 10, 12, 18, 23, 28, 33, 38, 43 };

            for (int i = 0; i < data.Count; i++)
            {
                for (int j = columnsToDelete.Count - 1; j >= 0; j--)
                {
                    int columnIndex = columnsToDelete[j];
                    if (columnIndex >= 0 && columnIndex < data[i].Count)
                    {
                        data[i].RemoveAt(columnIndex);
                    }
                }
            }
            return data;
        }

        private void SaveCredentials()
        {
            var secrets = new ClientSecrets
            {
                ClientId = ClientId,
                ClientSecret = ClientSecret
            };

            var json = JsonConvert.SerializeObject(secrets, Formatting.Indented);

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string filepath = Path.Combine(appDataPath, "TimeCollect", "credentials.json");
            Directory.CreateDirectory(Path.GetDirectoryName(filepath));

            File.WriteAllText(filepath, json);

            MessageBox.Show("Credentials saved successfully", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}

