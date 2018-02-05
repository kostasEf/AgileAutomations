using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using AgileAutomations.Interface;
using AgileAutomations.Model;
using OfficeOpenXml;

namespace AgileAutomations.Helper
{
    // ReSharper disable once ClassNeverInstantiated.Global
    public class ExcelHelper : IExcelHelper
    {
        public ObservableCollection<ContactFormData> ExtractRows(string path)
        {  
            using (var excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                return PopulateContactFormDataList(workSheet, true);
            }
        }

        public void AddReferenceNumbers(string path, ObservableCollection<ContactFormData> contactFormDataList)
        {
            using (var excelPackage = new ExcelPackage(new FileInfo(path)))
            {
                ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                var row = 2;
                int column = typeof(ContactFormData).GetProperties().Length;
                foreach (ContactFormData contactFormData in contactFormDataList)
                {
                    if (workSheet != null)
                    {
                        workSheet.Cells[row, column].Value = Convert.ToInt32(contactFormData.Reference);
                    }

                    row++;
                }
                excelPackage.Save();
            }          
        }

        public void OpenExcelFile(string filePath)
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

        private ObservableCollection<ContactFormData> PopulateContactFormDataList(ExcelWorksheet workSheet, bool firstRowHeader)
        {
            var contactFormDataCollection = new ObservableCollection<ContactFormData>();

            if (workSheet != null)
            {
                var headers = new Dictionary<string, int>();

                for (int rowIndex = workSheet.Dimension.Start.Row; rowIndex <= workSheet.Dimension.End.Row; rowIndex++)
                {
                    if (rowIndex == 1 && firstRowHeader)
                    {
                        headers = GetExcelHeader(workSheet, rowIndex);
                    }
                    else
                    {
                        contactFormDataCollection.Add(new ContactFormData
                        {
                            Name = ParseWorksheetValue(workSheet, headers, rowIndex, "Name"),
                            Email = ParseWorksheetValue(workSheet, headers, rowIndex, "Email"),
                            Subject = ParseWorksheetValue(workSheet, headers, rowIndex, "Subject"),
                            Message = ParseWorksheetValue(workSheet, headers, rowIndex, "Message")
                        });

                    }
                }
            }

            return contactFormDataCollection;
        }

        private Dictionary<string, int> GetExcelHeader(ExcelWorksheet workSheet, int rowIndex)
        {
            var headers = new Dictionary<string, int>();

            if (workSheet != null)
            {
                for (int columnIndex = workSheet.Dimension.Start.Column; columnIndex <= workSheet.Dimension.End.Column; columnIndex++)
                {
                    if (workSheet.Cells[rowIndex, columnIndex].Value != null)
                    {
                        string columnName = workSheet.Cells[rowIndex, columnIndex].Value.ToString();

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
            string value = string.Empty;
            var columnIndex = headers.ContainsKey(columnName) ? headers[columnName] : (int?)null;

            if (workSheet != null && columnIndex != null && workSheet.Cells[rowIndex, columnIndex.Value].Value != null)
            {
                value = workSheet.Cells[rowIndex, columnIndex.Value].Value.ToString();
            }

            return value;
        }
    }
}