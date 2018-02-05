using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Input;
using AgileAutomations.Model;

namespace AgileAutomations.Interface
{
    public interface IMainWindowViewModel
    {
        string Feedback { get; set; }
        bool FileSelected { get; set; }
        ObservableCollection<ContactFormData> ContactFormDataCollection { get; set; }
        ICommand SubmitFormCommand { get; }
        ICommand ExtractRowsCommand { get; }

        event PropertyChangedEventHandler PropertyChanged;
    }
}