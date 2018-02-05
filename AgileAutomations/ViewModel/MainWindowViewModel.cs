using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using AgileAutomations.Helper;
using AgileAutomations.Interface;
using AgileAutomations.Model;
using AgileAutomations.Properties;

namespace AgileAutomations.ViewModel
{
    public class MainWindowViewModel : INotifyPropertyChanged, IMainWindowViewModel
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public ObservableCollection<ContactFormData> ContactFormDataCollection { get; set; }
        private readonly IExcelHelper excelHelper;
        private readonly IFileBrowser fileBrowser;
        private readonly IHtmlHelper htmlHelper;
        private string filePath;

        public ICommand ExtractRowsCommand { get; private set; }
        public ICommand SubmitFormCommand { get; private set; }

        public MainWindowViewModel(IExcelHelper excelHelper, IFileBrowser fileBrowser, IHtmlHelper htmlHelper)
        {
            this.fileBrowser = fileBrowser;
            this.excelHelper = excelHelper;
            this.htmlHelper = htmlHelper;

            ExtractRowsCommand = new DelegateCommand(ExtractRows);
            SubmitFormCommand = new DelegateCommand(SubmitForm);
            ContactFormDataCollection = new ObservableCollection<ContactFormData>();
        }

        private string feedback;
        public string Feedback
        {
            get => feedback;
            set
            {
                if(feedback == value)
                {
                    return;
                }

                feedback = value;

                OnPropertyChanged();
            }
        }

        private bool fileSelected;
        public bool FileSelected
        {
            get => fileSelected;
            set
            {
                if (fileSelected == value)
                {
                    return;
                }

                fileSelected = value;

                OnPropertyChanged();
            }
        }

        private void ExtractRows()
        {
            filePath = fileBrowser.GetFullName();

            if (!string.IsNullOrWhiteSpace(filePath))
            {
                ContactFormDataCollection = excelHelper.ExtractRows(filePath);

                if (ContactFormDataCollection.Any())
                {
                    Feedback = $"Rows found: {ContactFormDataCollection.Count}";
                    FileSelected = true;
                }
            }
        }

        private void SubmitForm()
        {
            foreach (ContactFormData contactFormData in ContactFormDataCollection)
            {
                contactFormData.Reference = htmlHelper.SubmitContactForm(contactFormData);
            }

            excelHelper.AddReferenceNumbers(filePath, ContactFormDataCollection);

            Feedback = "Success";
            FileSelected = false;

            excelHelper.OpenExcelFile(filePath);
        }
        
        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}