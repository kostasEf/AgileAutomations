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
        private const string FILE_PATH = @"C:\TestData\AgileAutomationsTestData.xlsx";
        private ObservableCollection<ContactFormData> contactFormDataCollection;

        [OneTimeSetUp]
        public void Initialize()
        {
            excelHelper = new ExcelHelper();

            contactFormDataCollection = new ObservableCollection<ContactFormData>
            {
                new ContactFormData {Name = "Name1", Email = "Email1", Subject = "Subject1", Message = "Message1", Reference = "123"},
                new ContactFormData {Name = "Name2", Email = "Email2", Subject = "Subject2", Message = "Message2", Reference = "321"},
                new ContactFormData {Name = "Name3", Email = "Email3", Subject = "Subject3", Message = "Message3", Reference = "678"}
            };

            File.Delete(FILE_PATH);
            CreateAndPopulateExcel();
        }

        private void CreateAndPopulateExcel()
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(FILE_PATH)))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("Data");
                var columns = typeof(ContactFormData).GetProperties();
                var rows = contactFormDataCollection.Count;

                for (int i = 1; i <= columns.Length; i++)
                {
                    workSheet.Cells[1, i].Value = columns[i - 1].Name;
                }

                for (int i = 0; i < rows; i++)
                {
                    workSheet.Cells[i + 2, 1].Value = contactFormDataCollection[i].Name;
                    workSheet.Cells[i + 2, 2].Value = contactFormDataCollection[i].Email;
                    workSheet.Cells[i + 2, 3].Value = contactFormDataCollection[i].Subject;
                    workSheet.Cells[i + 2, 4].Value = contactFormDataCollection[i].Message;
                    workSheet.Cells[i + 2, 5].Value = contactFormDataCollection[i].Reference;
                }
                
                excelPackage.Save();
            }
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