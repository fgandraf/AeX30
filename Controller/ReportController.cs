using AeX30.Entities;
using AeX30.Repository;
using System.IO;

namespace AeX30.Controller
{
    public class ReportController
    {
        public static void SetReport(string templatePath, string saveAsPath, Report report)
        {
            if (IsValid(templatePath))
                ReportRepository.SetReport(templatePath, saveAsPath, report);
        }

        private static bool IsValid(string filePath)
        {
            string footer = ReportRepository.GetLeftFooter(filePath);
            string sheetName = ReportRepository.GetSheetName(filePath);

            bool fileExists = File.Exists(filePath);
            bool sheetNameIsValid = sheetName == "RAE";
            bool footerIsValid = footer != "" || footer != null;

            return fileExists && sheetNameIsValid && footerIsValid;
        }

    }
}







