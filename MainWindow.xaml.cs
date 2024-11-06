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
            DataContext = new MainViewModel();

            //Bind the sheetNameTextBox to the SheetNames property in the view model
            sheetNamesTextBox.SetBinding(TextBox.TextProperty, new Binding("SheetNames"));
        }
    }
}
