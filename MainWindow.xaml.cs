using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using TimeCollect.Models;
using TimeCollect.ViewModels;

namespace TimeCollect
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly MainViewModel viewModel;
        private readonly string employeeFilePath;



        public MainWindow()
        {

            viewModel = new MainViewModel();
            DataContext = viewModel;

            viewModel.LoadCredentials();

            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            employeeFilePath = Path.Combine(appDataPath, "TimeCollect", "employees.json");



            LoadEmployeeData();
            InitializeComponent();

            Closing += Window_Closing;
        }


        private void LoadEmployeeData()
        {
            if (File.Exists(employeeFilePath))
            {
                try
                {
                    string jsonData = File.ReadAllText(employeeFilePath);
                    var employees = JsonConvert.DeserializeObject<ObservableCollection<Employee>>(jsonData);
                    viewModel.Employees.Clear();
                    foreach (var employee in employees)
                    {
                        viewModel.Employees.Add(employee);
                    }
                }
                catch (Exception ex)
                {

                    Console.WriteLine($"Error loading employee data: {ex.Message}");
                }
            }
        }

        private void SaveEmployeeData()
        {
            try
            {
                string jsonData = JsonConvert.SerializeObject(viewModel.Employees, Formatting.Indented);
                Directory.CreateDirectory(Path.GetDirectoryName(employeeFilePath));
                File.WriteAllText(employeeFilePath, jsonData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving employee data: {ex.Message}");
            }
        }


        private void ClearLogButton_Click(object sender, RoutedEventArgs e)
        {
            viewModel.LogMessages = "";
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveEmployeeData();
        }

        private void LogTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            logTextBox.ScrollToEnd();
        }

    }
}
