using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace TimeCollect
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            DataContext = new MainViewModel();
            ((MainViewModel)DataContext).LoadCredentials();
            InitializeComponent();

            // Attach event handlers to PasswordBoxes
            databasePasswordBox.PasswordChanged += DatabasePasswordBox_PasswordChanged;
            clientSecretPasswordBox.PasswordChanged += ClientSecretPasswordBox_PasswordChanged;

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);


            // Load employee data from fole (if it exists)
            string employeeFilePath = Path.Combine(appDataPath, "TimeCollect", "employees.json");

            if (File.Exists(employeeFilePath))
            {
                Console.WriteLine(employeeFilePath);
                try
                {
                    string jsonData = File.ReadAllText(employeeFilePath);
                    var employees = JsonConvert.DeserializeObject<ObservableCollection<Employee>>(jsonData);
                    foreach (var employee in employees)
                    {
                        ((MainViewModel)DataContext).Employees.Add(employee);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading employee data: {ex.Message}");
                    //throw;
                }
            }
            else
            {
                // if the file doesn't exist, initialize on empty ObservationCollection
                ((MainViewModel)DataContext).Employees = new ObservableCollection<Employee>();
            }
            //Bind UI elements to ViewModel properties
            sheetNamesTextBox.SetBinding(TextBox.TextProperty, new Binding("SheetNames"));
            //outputDirectoryTextBox.SetBinding(TextBox.TextProperty, new Binding("OutputDirectory"));
        }

        private void ClearLogButton_Click(object sender, RoutedEventArgs e)
        {
            logTextBox.Text = ""; //Clear the logTextBoxs
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // 1. get the employees data from the viewmodel
            var employees = ((MainViewModel)DataContext).Employees;

            // 2. serialize the data to JSON
            string jsonData = JsonConvert.SerializeObject(employees);

            // 3. save the json to a file
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string filePath = Path.Combine(appDataPath, "TimeCollect", "employees.json");
            Console.WriteLine(filePath);
            // Ensuring the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            File.WriteAllText(filePath, jsonData);
        }

        private void LogTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            logTextBox.ScrollToEnd();
        }

        private void ClientSecretPasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (sender is PasswordBox passwordBox && DataContext is MainViewModel viewModel)
            {
                viewModel.ClientSecret = passwordBox.SecurePassword;
            }
        }

        private void DatabasePasswordBox_PasswordChanged(object sender, RoutedEventArgs e)
        {
            if (sender is PasswordBox passwordBox && DataContext is MainViewModel viewModel)
            {
                viewModel.DatabasePassword = passwordBox.SecurePassword;
            }
        }
    }
}
