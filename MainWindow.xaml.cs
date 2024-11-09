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
            InitializeComponent();

            // Load employee data from file
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string filePath = Path.Combine(appDataPath, "TimeCollect", "employees.json");

            if (File.Exists(filePath))
            {
                try
                {
                    string jsonData = File.ReadAllText(filePath);
                    var employees = JsonConvert.DeserializeObject<ObservableCollection<Employee>>(jsonData);
                    ((MainViewModel)DataContext).Employees = employees;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message.ToString());
                }
            }
            else
            {
                // if the file doesn't exist, initialize on empty ObservationCollection
                ((MainViewModel)DataContext).Employees = new ObservableCollection<Employee>();
            }

            DataContext = new MainViewModel();

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
