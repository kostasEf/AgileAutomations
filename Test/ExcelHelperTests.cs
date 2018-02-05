using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using AgileAutomations.Helper;
using AgileAutomations.Interface;
using AgileAutomations.Model;
using NUnit.Framework;
using OfficeOpenXml;

namespace Test
{
    [TestFixture]
    public class ExcelHelperTests
    {
        private IExcelHelper excelHelper;
        //I am aware that the best way would be to create an excel file every time the test runs
        private const string FILE_PATH = @"C:\TestData\AgileAutomations.xlsx";
        private ObservableCollection<ContactFormData> contactFormDataCollection;

        [OneTimeSetUp]
        public void Initialize()
        {
            excelHelper = new ExcelHelper();

            contactFormDataCollection = new ObservableCollection<ContactFormData>
            {
                new ContactFormData {Reference = "123"},
                new ContactFormData {Reference = "321"}
            };
        }

        [Test]
        public void ExtractRows_WhenCalled_ReturnsPopulatedCollection()
        {
            Assert.IsTrue(excelHelper.ExtractRows(FILE_PATH).Any());
        }

        [Test]
        public void AddReferenceNumbers_WhenCalled_ReturnsPopulatedCollection()
        {
            //Setup
            bool result = true;

            //Act
            excelHelper.AddReferenceNumbers(FILE_PATH, contactFormDataCollection);

            using (var excelPackage = new ExcelPackage(new FileInfo(FILE_PATH)))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                var row = 2;
                int column = typeof(ContactFormData).GetProperties().Count();
                foreach (ContactFormData contactFormData in contactFormDataCollection)
                {
                    if (workSheet != null)
                    {
                        if (workSheet.Cells[row, column].Value.ToString() != contactFormData.Reference)
                        {
                            result = false;
                        }
                    }
                    row++;
                }
            }

            //Assert
            Assert.IsTrue(result);
        }
    }
}