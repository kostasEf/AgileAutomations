using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
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
        private IList<ContactFormData> contactFormDataList;
        private IExcelParser excelParser;
        private IFileBrowser fileBrowser;
        private IHtmlHelper htmlHelper;
        private string filePath;

        public ICommand SelectFileCommand { get; private set; }
        public ICommand ImportCommand { get; private set; }

        public MainWindowViewModel(IExcelParser excelParser, IFileBrowser fileBrowser, IHtmlHelper htmlHelper)
        {
            this.fileBrowser = fileBrowser;
            this.excelParser = excelParser;
            this.htmlHelper = htmlHelper;

            SelectFileCommand = new DelegateCommand(SelectFile);
            ImportCommand = new DelegateCommand(SubmitForm);
        }

        private string feedback;
        public string Feedback {
            get => feedback;
            set
            {
                if(feedback == value)
                    return;

                feedback = value;

                OnPropertyChanged();
            }
        }

        private void SelectFile()
        {
            filePath = fileBrowser.GetFullName();

            if (!string.IsNullOrWhiteSpace(filePath))
            {
                contactFormDataList = excelParser.ExtractRows(filePath);
                Feedback = $"Rows found: {contactFormDataList.Count}";
            }
        }

        private void SubmitForm()
        {
            foreach (var contactFormData in contactFormDataList)
            {
                contactFormData.Reference = htmlHelper.SubmitContactForm(contactFormData);
            }

            excelParser.AddReferenceNumbers(filePath, contactFormDataList.ToList());

            Feedback = "Success";

            OpenExcelFile();
        }

        private void OpenExcelFile()
        {
            var start = new ProcessStartInfo
            {
                FileName = filePath
            };

            using (var process = new Process { StartInfo = start })
            {
                process.Start();
                process.WaitForExit();
            }
        }

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}