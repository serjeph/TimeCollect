using ClosedXML.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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

namespace TimeCollect.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {


        private readonly string _credentialsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", "credentials.json");
        private readonly string _googleCredentialsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", "googlecredentials.json");
        private readonly string _sheetNamesFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", "sheetNames.json");
        private readonly string _weekNamesFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TimeCollect", "weekNames.json");

        private readonly GoogleSheetsService _googleSheetsService;
        //private readonly DatabaseService _databaseService;

        // to be bound with xaml
        public ObservableCollection<WeekType> WeekTypes { get; set; }
        public ObservableCollection<SheetNamesList> SheetNames { get; set; }
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
        public ICommand SaveSheetNamesCommand { get; set; }
        public ICommand ExportLogCommand { get; set; }
        public ICommand SaveWeekTypeCommand { get; set; }
        public ICommand ReloadWeekTypeCommand { get; set; }




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
                OnPropertyChanged(nameof(ClientSecret));
            }
        }



        public DateTime YearStartDate { get; set; } = new DateTime(2023, 12, 31); // Default date



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

        private bool _isLoading;
        public bool IsLoading
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
            // Load employee data from file
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string employeeFilePath = Path.Combine(appDataPath, "TimeCollect", "employees.json");
            string sheetNamesFilePath = Path.Combine(appDataPath, "TimeCollect", "sheetNames.json");


            // Initialize collections
            WeekTypes = new ObservableCollection<WeekType>();
            SheetNames = new ObservableCollection<SheetNamesList>();
            Employees = new ObservableCollection<Employee>();

            // Initialize UI commands
            RunDataCommand = new RelayCommand(async () => await RunData());
            SaveCredentialsCommand = new RelayCommand(SaveCredentials);
            SaveDatabaseSettingsCommand = new RelayCommand(SaveDatabaseSettings);
            ExportLogCommand = new RelayCommand(ExportLog);
            SaveWeekTypeCommand = new RelayCommand(SaveWeekTypeData);
            SaveSheetNamesCommand = new RelayCommand(SaveSheetNames);
            ReloadWeekTypeCommand = new RelayCommand(ReloadWeekType);

            // Initialize services
            _googleSheetsService = new GoogleSheetsService(_credentialsPath);

            //_databaseService = new DatabaseService(new DatabaseSettings
            //{
            //    Host = DatabaseHost,
            //    Database = DatabaseName,
            //    Username = DatabaseUsername,
            //    Password = DatabasePassword,
            //    Port = DatabasePort
            //});

            // Load data
            LoadYearStartDate();
            LoadWeekTypeData();
            LoadDatabaseSettings();
            LoadCredentials();
            LoadSheetNamesData();
        }


        private async Task RunData()
        {

            var columnHeaders = new List<string>
                    {
                        "uuid",
                        "社員番号",
                        "年",
                        "月",
                        "日",
                        "WeekType",
                        "名前",
                        "工号",
                        "種別",
                        "間接/直接",
                        "原寸/3D/管理",
                        "時間"
                    };

            LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Timesheet Collection Started...\n";
            try
            {
                IsRunEnabled = false;
                IsLoading = true;

                var service = await _googleSheetsService.CreateSheetsService();

                var allCleanedData = new List<IList<object>>();

                string outputDirectory = OutputDirectory;
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                var sheetNames = SheetNames.ToList();
                foreach (var sheetName in sheetNames)
                {
                    int uuid = 0;
                    var sheetData = new List<IList<object>>();
                    var transformedValues = new List<IList<object>>();

                    List<string> workType = new List<string>();
                    workType.AddRange(Enumerable.Repeat("0.00", 4));
                    workType.AddRange(Enumerable.Repeat("間接", 5));
                    workType.AddRange(Enumerable.Repeat("直接", 20));

                    List<string> team = new List<string>();
                    team.AddRange(Enumerable.Repeat("0.00", 4));
                    team.AddRange(Enumerable.Repeat("-", 2));
                    team.AddRange(Enumerable.Repeat("管理", 1));
                    team.AddRange(Enumerable.Repeat("-", 2));
                    team.AddRange(Enumerable.Repeat("原寸", 20));

                    List<string> team3D = new List<string>();
                    team3D.AddRange(Enumerable.Repeat("0.00", 4));
                    team3D.AddRange(Enumerable.Repeat("-", 2));
                    team3D.AddRange(Enumerable.Repeat("管理", 1));
                    team3D.AddRange(Enumerable.Repeat("-", 2));
                    team3D.AddRange(Enumerable.Repeat("3D", 20));


                    foreach (var employee in Employees)
                    {
                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{sheetName.SheetName}]-[{employee.Nickname}] Data fetching...\n";
                        var range = $"{sheetName.SheetName}!A7:AR39";
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

                                    string weekType = DateTimeHelper.GetWeekType(year, month, day);
                                    string isActual = DateHelper.IsActual(year, month, day);

                                    transformedValues.Add(new List<object>()
                                    {
                                        uuid + 1,
                                        employee.EmployeeId,
                                        year,
                                        month,
                                        day,
                                        weekType,
                                        employee.Nickname,
                                        employeeData[0][col + 4],
                                        employeeData[1][col + 4],
                                        workType[col + 4],
                                        employee.Team,
                                        float.Parse(employeeData[row + 2][col + 4].ToString())
                                    });


                                    uuid++;

                                }
                            }
                        }
                        else
                        {
                            LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] No data found for {employee.Nickname} from sheet {sheetName.SheetName}\n"; ;
                        }

                        Console.WriteLine(transformedValues);

                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{sheetName.SheetName}]-[{employee.Nickname}] Data fetched...\n";
                    }

                    if (transformedValues.Count > 0)
                    {


                        ExcelHelper.ExportToExcel(transformedValues, sheetName.SheetName, Path.Combine(outputDirectory, $"output_{sheetName.SheetName}.xlsx"), columnHeaders);
                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{sheetName.SheetName}] Monthly timesheet data was collected to {outputDirectory}\\output_{sheetName.SheetName}.xlsx\n";

                        allCleanedData.AddRange(transformedValues);
                        //string logMessages = LogMessages;
                        //_databaseService.InsertData(transformedValues, sheetName.SheetName, ref logMessages);
                        //LogMessages = logMessages;
                        LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Timesheets collected to Excel with {transformedValues.Count} rows.\n";
                    }
                }

                string combinedFilePath = Path.Combine(outputDirectory, "combined_output.xlsx");
                CombineExcelFiles(SheetNames, outputDirectory, combinedFilePath);
                LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Combined output location: {outputDirectory} \n";

                if (allCleanedData.Count > 0)
                {

                    ExcelHelper.ExportToExcel(allCleanedData, "AllSheets", Path.Combine(outputDirectory, "output_all_sheets.xlsx"), columnHeaders);
                    LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Data exported to excel file: {outputDirectory} \n";

                }

            }
            catch (Exception ex)
            {
                LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] An error occurred: {ex.Message} \n";
            }
            finally
            {
                IsRunEnabled = true;
                IsLoading = false;
                LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Data Processing Completed!\n";
                ExportLog();
            }
        }

        private void LoadDatabaseSettings()
        {
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
        }

        private void ExportLog()
        {
            try
            {
                string logFilePath = Path.Combine(OutputDirectory, "log.txt");
                File.WriteAllText(logFilePath, LogMessages);
                MessageBox.Show($"Log exported to {logFilePath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting log: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadWeekTypeData()
        {
            try
            {
                var data = WeekTypeHelper.GetWeekTypes(YearStartDate);

                foreach (var weekType in data)
                {
                    WeekTypes.Add(weekType);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading week type data: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void LoadYearStartDate()
        {
            try
            {
                if (File.Exists(_weekNamesFilePath))
                {
                    var jsonData = File.ReadAllText(_weekNamesFilePath);
                    var data = JsonConvert.DeserializeObject<ObservableCollection<WeekType>>(jsonData);
                    YearStartDate = data[0].DateStart;
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Error loading year start date: {ex.Message}");
            }
        }

        private void SaveWeekTypeData()
        {
            try
            {
                string jsonData = JsonConvert.SerializeObject(WeekTypes, Formatting.Indented);
                Directory.CreateDirectory(Path.GetDirectoryName(_weekNamesFilePath));
                File.WriteAllText(_weekNamesFilePath, jsonData);

                MessageBox.Show("Week type data saved sucessfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving week type data: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }



        private void ReloadWeekType()
        {
            try
            {
                var data = WeekTypeHelper.SetWeekTypes(YearStartDate);
                WeekTypes.Clear();
                foreach (var weekType in data)
                {
                    WeekTypes.Add(weekType);
                }

                MessageBox.Show("Week types reloaded sucessfully!", "Sucess", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {

                Console.WriteLine($"Error reloading week types: {ex.Message}");
            }

        }

        private void CombineExcelFiles(ObservableCollection<SheetNamesList> SheetNames, string outputDirectory, string combinedFilePath)
        {
            try
            {
                if (File.Exists(combinedFilePath))
                {
                    File.Delete(combinedFilePath);
                }

                using (var combinedWorkbook = new XLWorkbook())
                {
                    var sheetNames = SheetNames.ToList();
                    foreach (var sheetName in sheetNames)
                    {
                        string sheetFilePath = Path.Combine(outputDirectory, $"output_{sheetName.SheetName}.xlsx");
                        using (var sheetWorkbook = new XLWorkbook(sheetFilePath))
                        {
                            var worksheet = sheetWorkbook.Worksheet(1); // Get the first worksheet
                            worksheet.CopyTo(combinedWorkbook, sheetName.SheetName); //Copy to the combined workbook
                        }
                    }

                    combinedWorkbook.SaveAs(combinedFilePath);
                }
                LogMessages += $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] All monthly data was exported to {combinedFilePath}\n";

            }
            catch (Exception ex)
            {

                MessageBox.Show($"Error combining Excel files: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
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
            if (File.Exists(_credentialsPath))
            {
                var googleCredentials = new GoogleCredentials
                {
                    GoogleClientId = ClientId,
                    GoogleProjectId = ProjectId,
                    GoogleClientSecret = ClientSecret
                };


                string googleJson = JsonConvert.SerializeObject(googleCredentials, Formatting.Indented);

                Directory.CreateDirectory(Path.GetDirectoryName(_googleCredentialsPath));

                File.WriteAllText(_googleCredentialsPath, googleJson);

                MessageBox.Show("Credentials saved successfully", "Success", MessageBoxButton.OK, MessageBoxImage.Information);

            }


        }

        public void LoadCredentials()
        {
            if (!File.Exists(_credentialsPath))
            {
                MessageBox.Show("Please provide GCP credentials json", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                string jsonData = File.ReadAllText(_credentialsPath);
                JObject json = JObject.Parse(jsonData);
                JObject installedObject = (JObject)json["installed"];

                ClientId = installedObject["client_id"].ToString();
                ProjectId = installedObject["project_id"].ToString();
                ClientSecret = installedObject["client_secret"].ToString();
            }
            catch (Exception ex)
            {
                // Handle exceptions (e.g., log, rethrow, or display a user-friendly message)
                Console.WriteLine($"Error loading credentials: {ex.Message}");
            }
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

        private void SaveSheetNames()
        {
            try
            {
                string jsonData = JsonConvert.SerializeObject(SheetNames);
                Directory.CreateDirectory(Path.GetDirectoryName(_sheetNamesFilePath));
                File.WriteAllText(_sheetNamesFilePath, jsonData);

                MessageBox.Show("Sheet name settings successfully!", "Sucess", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving sheet name data: {ex.Message}");
            }

        }

        private void LoadSheetNamesData()
        {
            {
                try
                {
                    string json = File.ReadAllText(_sheetNamesFilePath);
                    var sheetNames = JsonConvert.DeserializeObject<ObservableCollection<SheetNamesList>>(json);
                    foreach (var sheetName in sheetNames)
                    {
                        SheetNames.Add(sheetName);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading Sheet Names: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// PropertyChanged handler
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

