using AeX30.Domain.Entities;
using AeX30.Infra.Repository;
using System.IO;

namespace AeX30.App.Services
{
    public class ReportService
    {
        private ProposalRepository _proposalRepository;

        public ReportService()
            => _proposalRepository = new ProposalRepository();


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
            string footer = _proposalRepository.GetLeftFooter(filePath);
            string sheetName = _proposalRepository.GetSheetName(filePath);

            bool fileExists = File.Exists(filePath);
            bool sheetNameIsValid = sheetName == "RAE";
            bool footerIsValid = !string.IsNullOrEmpty(footer);

            return fileExists && sheetNameIsValid && footerIsValid;
        }

    }
}