using System.Collections.Generic;
using System.IO;
using System.Linq;
using AgileAutomations.Model;
using OfficeOpenXml;

namespace AgileAutomations.Helper
{
    public class ExcelParser
    {
        public IEnumerable<ContactFormData> ExtractData(string path)
        {  
            using (var excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                var workSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                return PopulateContactFormDataList(workSheet, true);
            }
        }

 

        private IEnumerable<ContactFormData> PopulateContactFormDataList(ExcelWorksheet workSheet, bool firstRowHeader)
        {
            IList<ContactFormData> contactFormDataList = new List<ContactFormData>();

            if (workSheet != null)
            {
                var header = new Dictionary<string, int>();

                for (var rowIndex = workSheet.Dimension.Start.Row; rowIndex <= workSheet.Dimension.End.Row; rowIndex++)
                {
                    if (rowIndex == 1 && firstRowHeader)
                    {
                        header = ExcelHelper.GetExcelHeader(workSheet, rowIndex);
                    }
                    else
                    {
                        contactFormDataList.Add(new ContactFormData
                        {
                            Name = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Name"),
                            Email = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Email"),
                            Subject = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Subject"),
                            Message = ExcelHelper.ParseWorksheetValue(workSheet, header, rowIndex, "Message"),
                        });

                    }
                }
            }

            return contactFormDataList;
        }
    }

    public static class ExcelHelper
    {

        public static Dictionary<string, int> GetExcelHeader(ExcelWorksheet workSheet, int rowIndex)
        {
            var header = new Dictionary<string, int>();

            if (workSheet != null)
            {
                for (var columnIndex = workSheet.Dimension.Start.Column; columnIndex <= workSheet.Dimension.End.Column; columnIndex++)
                {
                    if (workSheet.Cells[rowIndex, columnIndex].Value != null)
                    {
                        var columnName = workSheet.Cells[rowIndex, columnIndex].Value.ToString();

                        if (!header.ContainsKey(columnName) && !string.IsNullOrEmpty(columnName))
                        {
                            header.Add(columnName, columnIndex);
                        }
                    }
                }
            }

            return header;
        }


        public static string ParseWorksheetValue(ExcelWorksheet workSheet, Dictionary<string, int> header, int rowIndex, string columnName)
        {
            var value = string.Empty;
            var columnIndex = header.ContainsKey(columnName) ? header[columnName] : (int?)null;

            if (workSheet != null && columnIndex != null && workSheet.Cells[rowIndex, columnIndex.Value].Value != null)
            {
                value = workSheet.Cells[rowIndex, columnIndex.Value].Value.ToString();
            }

            return value;
        }
    }
}