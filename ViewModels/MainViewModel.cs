using Google.Apis.Auth.OAuth2;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using TimeCollect.Commands;
using TimeCollect.Helpers;
using TimeCollect.Models;
using TimeCollect.Services;

namespace TimeCollect
{
    public class MainViewModel : INotifyPropertyChanged
    {

        private readonly string _credentialsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", "credentials.json");
        private readonly GoogleSheetsService _googleSheetsService;
        private readonly DatabaseService _databaseService;

        public ObservableCollection<Employee> Employees { get; set; }

        public string DatabaseHost { get; set; } = "localhost"; // Default host
        public string DatabaseName { get; set; }
        public string DatabaseUsername { get; set; } = "postgres"; // Default username
        public string DatabasePassword { get; set; }
        public int DatabasePort { get; set; } = 5433; // Default port
        private readonly string _databaseSettingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", "database_settings.json");


        public ICommand RunDataCommand { get; set; }
        public ICommand SaveCredentialsCommand { get; set; }
        public ICommand SaveDatabaseSettingsCommand { get; set; }



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

        private string _projectId;
        public string ProjectId
        {
            get => _projectId;
            set
            {
                _projectId = value;
                OnPropertyChanged(nameof(ProjectId));
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

        private string _outputDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "TimeCollect");
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

        private string _logMessages;
        public string LogMessages
        {
            get => _logMessages;
            set
            {
                _logMessages = value;
                OnPropertyChanged(nameof(LogMessages));
            }
        }

        // Constructor
        public MainViewModel()
        {
            Employees = new ObservableCollection<Employee>();
            DotNetEnv.Env.Load();
#if DEBUG
            SheetNames = "202409,202410";
#endif

            // Load employee data from file
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string employeeFilePath = Path.Combine(appDataPath, "TimeCollect", "employees.json");


            RunDataCommand = new RelayCommand(async () => await RunData());
            SaveCredentialsCommand = new RelayCommand(SaveCredentials);
            SaveDatabaseSettingsCommand = new RelayCommand(SaveDatabaseSettings);

            _googleSheetsService = new GoogleSheetsService(_credentialsPath);

            if (File.Exists(_databaseSettingsPath))
            {
                try
                {
                    string json = File.ReadAllText(_databaseSettingsPath);
                    var settings = JsonConvert.DeserializeObject<DatabaseSettings>(json);

                    DatabaseHost = settings.Host;
                    DatabaseName = settings.Database;
                    DatabaseUsername = settings.Username;
                    DatabasePassword = settings.Password;
                    DatabasePort = settings.Port;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading database settings: {ex.Message}");
                    throw;
                }
            }

            _databaseService = new DatabaseService(new DatabaseSettings
            {
                Host = DatabaseHost,
                Database = DatabaseName,
                Username = DatabaseUsername,
                Password = DatabasePassword,
                Port = DatabasePort
            });
        }


        private async Task RunData()
        {
            LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Data Processing Started...\n";
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
                    int uuid = 0;
                    var sheetData = new List<IList<object>>();
                    var transformedValues = new List<IList<object>>();

                    List<string> workType = new List<string>();
                    workType.AddRange(Enumerable.Repeat("0.00", 4));
                    workType.AddRange(Enumerable.Repeat("間接", 5));
                    workType.AddRange(Enumerable.Repeat("直接", 20));

                    foreach (var employee in Employees)
                    {
                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{sheetName}]-[{employee.Nickname}] Data fetching...\n";
                        var range = $"{sheetName}!A7:AR39";
                        var request = service.Spreadsheets.Values.Get(employee.SpreadsheetId, range);
                        var response = await request.ExecuteAsync();
                        var values = response.Values;

                        values[0].Add("0.00");
                        values[0].Add("0.00");
                        values[0].Add("0.00");
                        values[0].Add("0.00");

                        var employeeData = new List<IList<object>>();

                        if (values != null && values.Count > 0)
                        {

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
                            employeeData.AddRange(CleanData(values));

                            Console.WriteLine(employeeData);

                            for (int i = 0; i < 9; i++)
                            {
                                employeeData[1][i] = employeeData[0][i];
                            }

                            for (int i = 0; i < 3; i++)
                            {
                                foreach (int j in new List<int> { 10, 14, 18, 22, 26 })
                                {
                                    employeeData[0][i + j] = employeeData[0][j - 1];
                                }
                            }
                            Console.WriteLine(employeeData);

                            for (int col = 0; col < 25; col++)
                            {
                                for (int row = 0; row < employeeData.Count - 2; row++)
                                {
                                    int year = int.Parse(employeeData[row + 2][0].ToString());
                                    int month = int.Parse(employeeData[row + 2][1].ToString());
                                    int day = int.Parse(employeeData[row + 2][2].ToString());

                                    string weekType = DateTimeHelper.GetWeekType($"{year}-{month}-{day}");
                                    string isActual = DateHelper.IsActual($"{year}-{month}-{day}");

                                    transformedValues.Add(new List<object>()
                                    {
                                        uuid + 1,
                                        employee.EmployeeId,
                                        row + 1,
                                        year,
                                        month,
                                        day,
                                        weekType,
                                        employee.Nickname,
                                        employeeData[0][col + 4],
                                        employeeData[1][col + 4],
                                        workType[col + 4],
                                        isActual,
                                        float.Parse(employeeData[row + 2][col + 4].ToString())
                                    });


                                    uuid++;

                                }
                            }
                        }
                        else
                        {
                            LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] No data found for {employee.Nickname} from sheet {sheetName}\n"; ;
                        }

                        Console.WriteLine(transformedValues);

                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{sheetName}]-[{employee.Nickname}] Data fetched...\n";
                    }

                    if (transformedValues.Count > 0)
                    {
                        allCleanedData.AddRange(transformedValues);

                        // Get column headers from the database
                        var columnHeaders = _databaseService.GetColumnHeaders($"timesheet_{sheetName}");

                        ExcelHelper.ExportToExcel(transformedValues, sheetName, Path.Combine(outputDirectory, $"output_{sheetName}.xlsx"), columnHeaders);
                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Monthly data was exported to {outputDirectory}\\output_{sheetName}.xlsx\n";
                        _databaseService.InsertData(transformedValues, sheetName);
                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Data Uploaded to Database with {transformedValues.Count} rows.\n";
                    }
                }

                if (allCleanedData.Count > 0)
                {
                    var columnHeaders = new List<string>
                    {
                        "uuid",
                        "employee_id",
                        "row_id",
                        "年",
                        "月",
                        "日",
                        "week_type",
                        "名前",
                        "工号",
                        "task_type",
                        "work_type",
                        "is_actual",
                        "時間"
                    };

                    ExcelHelper.ExportToExcel(allCleanedData, "AllSheets", Path.Combine(outputDirectory, "output_all_sheets.xlsx"), columnHeaders);
                    LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Data exported to excel file: {outputDirectory} \n";

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
                LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Data Processing Completed!\n";
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
            var secrets = new
            {
                installed = new
                {
                    client_id = ClientId,
                    project_id = ProjectId,
                    auth_uri = "https://accounts.google.com/o/oauth2/auth",
                    token_uri = "https://oauth2.googleapis.com/token",
                    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs",
                    client_secret = ClientSecret,
                    redirect_uris = new[] { "http://localhost" }
                }
            };

            string json = JsonConvert.SerializeObject(secrets, Formatting.Indented);

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string filepath = Path.Combine(appDataPath, "TimeCollect", "credentials.json");
            string tokenFilePath = Path.Combine(appDataPath, "TimeCollect", "token.json");
            Directory.CreateDirectory(Path.GetDirectoryName(filepath));

            File.WriteAllText(filepath, json);
            if (!File.Exists(tokenFilePath))
            {
                File.WriteAllText(tokenFilePath, "{}");
            }

            MessageBox.Show("Credentials saved successfully", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public void LoadCredentials()
        {
            if (File.Exists(_credentialsPath))
            {
                try
                {
                    string json = File.ReadAllText(_credentialsPath);
                    var secrets = JsonConvert.DeserializeObject<ClientSecrets>(json);

                    ClientId = secrets.ClientId;
                    ClientSecret = secrets.ClientSecret;
                }
                catch (Exception ex)
                {
                    //Handle exceptions
                    Console.WriteLine($"Error loading credentials: {ex.Message}");
                }
            }
        }


        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void SaveDatabaseSettings()
        {
            var settings = new DatabaseSettings
            {
                Host = DatabaseHost,
                Database = DatabaseName,
                Username = DatabaseUsername,
                Password = DatabasePassword,
                Port = DatabasePort
            };

            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            Directory.CreateDirectory(Path.GetDirectoryName(_databaseSettingsPath));
            File.WriteAllText(_databaseSettingsPath, json);

            MessageBox.Show("Database settings saved succesfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

}

