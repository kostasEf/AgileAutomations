using System.Collections.ObjectModel;
using AgileAutomations.Model;

namespace AgileAutomations.Interface
{
    public interface IExcelHelper
    {
        void AddReferenceNumbers(string path, ObservableCollection<ContactFormData> contactFormDataList);
        ObservableCollection<ContactFormData> ExtractRows(string path);
        void OpenExcelFile(string filePath);
    }
}