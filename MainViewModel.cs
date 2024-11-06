using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;

namespace TimeCollect
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<Employee> Employees { get; set; }


        // add other properties and commands here later...
        public ICommand RunDataCommand { get; set; }
        public ICommand SaveCredentialsCommand { get; set; }

        public MainViewModel()
        {
            Employees = new ObservableCollection<Employee>();
            RunDataCommand = new RelayCommand(RunData);
            SaveCredentialsCommand = new RelayCommand(SaveCredentials);
        }

        private void RunData()
        {
            //TODO: Implement data processing logic here
        }

        private void SaveCredentials()
        {
            //TODO: Implement credential saving logic here
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
