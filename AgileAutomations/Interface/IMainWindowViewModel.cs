using System.ComponentModel;
using System.Windows.Input;

namespace AgileAutomations.Interface
{
    public interface IMainWindowViewModel
    {
        string Feedback { get; set; }
        ICommand ImportCommand { get; }
        ICommand SelectFileCommand { get; }

        event PropertyChangedEventHandler PropertyChanged;
    }
}