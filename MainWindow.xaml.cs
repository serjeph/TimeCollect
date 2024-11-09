﻿using Newtonsoft.Json;
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

            // Grant permissions to AppData\Roaming\TimeCollect directory
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            //string timeCollectDirectory = Path.Combine(appDataPath, "TimeCollect");
            //
            //if (!Directory.Exists(timeCollectDirectory))
            //{
            //    Directory.CreateDirectory(timeCollectDirectory);
            //}
            //
            //try
            //{
            //    SecurityIdentifier sid = new SecurityIdentifier(WellKnownSidType.BuiltinUsersSid, null);
            //    //string userSid = System.Security.Principal.WindowsIdentity.GetCurrent().User.Value;
            //    //Console.WriteLine(userSid);
            //    DirectorySecurity directorySecurity = Directory.GetAccessControl(timeCollectDirectory);
            //    FileSystemAccessRule accessRule = new FileSystemAccessRule(
            //        sid,
            //        FileSystemRights.FullControl,
            //        InheritanceFlags.ContainerInherit | InheritanceFlags.ObjectInherit,
            //        PropagationFlags.None,
            //        AccessControlType.Allow);
            //    Console.WriteLine(accessRule);
            //    directorySecurity.AddAccessRule(accessRule);
            //    Directory.SetAccessControl(timeCollectDirectory, directorySecurity);
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"Error setting directory permissions: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            //}

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
            clientIdTextBox.SetBinding(TextBox.TextProperty, new Binding("ClientId"));
            clientSecretTextBox.SetBinding(TextBox.TextProperty, new Binding("ClientSecret"));
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
    }
}
