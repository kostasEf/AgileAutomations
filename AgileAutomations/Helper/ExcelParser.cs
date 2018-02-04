using System.Collections.Generic;
using System.IO;
using System.Linq;
using AgileAutomations.Interface;
using AgileAutomations.Model;
using OfficeOpenXml;

namespace AgileAutomations.Helper
{
    public class ExcelParser : IExcelParser
    {
        public IList<ContactFormData> ExtractRows(string path)
        {  
            using (var excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                var workSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                return PopulateContactFormDataList(workSheet, true).ToList();
            }
        }

        public void AddReferenceNumbers(string path, IList<ContactFormData> contactFormDataList)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                var workSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                var row = 2;
                foreach (var contactFormData in contactFormDataList)
                {
                    workSheet.Cells[row, 5].Value = contactFormData.Reference;
                    row++;
                }
                excelPackage.Save();
            }          
        }

        private IList<ContactFormData> PopulateContactFormDataList(ExcelWorksheet workSheet, bool firstRowHeader)
        {
            IList<ContactFormData> contactFormDataList = new List<ContactFormData>();

            if (workSheet != null)
            {
                var headers = new Dictionary<string, int>();

                for (var rowIndex = workSheet.Dimension.Start.Row; rowIndex <= workSheet.Dimension.End.Row; rowIndex++)
                {
                    if (rowIndex == 1 && firstRowHeader)
                    {
                        headers = GetExcelHeader(workSheet, rowIndex);
                    }
                    else
                    {
                        contactFormDataList.Add(new ContactFormData
                        {
                            Name = ParseWorksheetValue(workSheet, headers, rowIndex, "Name"),
                            Email = ParseWorksheetValue(workSheet, headers, rowIndex, "Email"),
                            Subject = ParseWorksheetValue(workSheet, headers, rowIndex, "Subject"),
                            Message = ParseWorksheetValue(workSheet, headers, rowIndex, "Message"),
                        });

                    }
                }
            }

            return contactFormDataList;
        }

        private Dictionary<string, int> GetExcelHeader(ExcelWorksheet workSheet, int rowIndex)
        {
            var headers = new Dictionary<string, int>();

            if (workSheet != null)
            {
                for (var columnIndex = workSheet.Dimension.Start.Column; columnIndex <= workSheet.Dimension.End.Column; columnIndex++)
                {
                    if (workSheet.Cells[rowIndex, columnIndex].Value != null)
                    {
                        var columnName = workSheet.Cells[rowIndex, columnIndex].Value.ToString();

                        if (!headers.ContainsKey(columnName) && !string.IsNullOrEmpty(columnName))
                        {
                            headers.Add(columnName, columnIndex);
                        }
                    }
                }
            }

            return headers;
        }

        private string ParseWorksheetValue(ExcelWorksheet workSheet, Dictionary<string, int> headers, int rowIndex, string columnName)
        {
            var value = string.Empty;
            var columnIndex = headers.ContainsKey(columnName) ? headers[columnName] : (int?)null;

            if (workSheet != null && columnIndex != null && workSheet.Cells[rowIndex, columnIndex.Value].Value != null)
            {
                value = workSheet.Cells[rowIndex, columnIndex.Value].Value.ToString();
            }

            return value;
        }
    }
}