using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
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
            //1. Fetch data from Google Sheets for each employee
            var credential = await GetUserCredential();
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "TimeCollect"
            });

            foreach (var employee in Employees)
            {
                var sheetNames = SheetNames.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (var sheetName in sheetNames)
                {
                    var range = $"{sheetName}!A:Z"; //i'll edit this range later...
                    var request = service.Spreadsheets.Values.Get(employee.SpreadsheetId, range);
                    var response = await request.ExecuteAsync();
                    var values = response.Values;

                    //2. Process the fetched data (clean, export, insert) - for now, just log
                    if (values != null && values.Count > 0)
                    {
                        LogMessages += $"Fetched {values.Count} rows for {employee.Nickname}\n";
                        //TODO: Implement CleanData, ExportToExcel, InsertIntoDatabase
                    }
                    else
                    {
                        LogMessages += $"No data found for {employee.Nickname}\n";
                    }
                }
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

    }
}
