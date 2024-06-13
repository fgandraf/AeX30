using AeX30.Domain.Entities;
using AeX30.Infra.Repository;
using OfficeOpenXml;
using System.IO;

namespace AeX30.App.Services
{
    public class ReportService
    {

        public bool SetReport(string templatePath, string saveAsPath, Report report)
        {
            if (IsValid(templatePath))
            {
                ReportRepository.SetReport(templatePath, saveAsPath, report);
                return true;
            }
            else
                return false;
        }

        private bool IsValid(string filePath)
        {

            string footer;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                footer = worksheet.HeaderFooter.OddFooter.LeftAlignedText;
            }

            string sheetName;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                sheetName = worksheet.Name;
            }


            bool fileExists = File.Exists(filePath);
            bool sheetNameIsValid = sheetName == "RAE";
            bool footerIsValid = !string.IsNullOrEmpty(footer);

            return fileExists && sheetNameIsValid && footerIsValid;
        }

    }
}