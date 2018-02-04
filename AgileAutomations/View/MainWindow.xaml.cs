using System.Collections.Generic;
using System.Linq;
using AgileAutomations.Helper;
using AgileAutomations.Interface;
using System.Windows;
using AgileAutomations.Model;

namespace AgileAutomations
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private IFileBrowser fileBrowser;
        private ExcelParser excelParser;
        private IEnumerable<ContactFormData> contactFormDataList;

        public MainWindow()
        {
            InitializeComponent();

            fileBrowser = new FileBrowser();
            excelParser = new ExcelParser();
        }

        private void SelectFile(object sender, RoutedEventArgs e)
        {
            contactFormDataList = excelParser.ExtractData(fileBrowser.GetFullName());
            Rows.Content = $"Rows found: {contactFormDataList.Count()}";
        }

        private void Import(object sender, RoutedEventArgs e)
        {
            throw new System.NotImplementedException();
        }
    }
}
    