using System.Collections.ObjectModel;
using System.ComponentModel;

namespace TimeCollect
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public ObservableCollection<Employee> Employees { get; set; }


        // add other properties and commands here later...


        public MainViewModel()
        {
            Employees = new ObservableCollection<Employee>();
        }

        // Implement INotifyPropertyChange interface later...

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
