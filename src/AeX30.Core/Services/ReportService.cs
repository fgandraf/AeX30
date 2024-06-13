using AeX30.Core.Entities;
using AeX30.Core.Interfaces;
using OfficeOpenXml;
using System;
using System.IO;

namespace AeX30.Core.Services
{
    public class ReportService : IReportService
    {
        public bool SaveReport(string templatePath, string saveAsPath, Report report)
        {
            var footer = LoadLeftFooter(templatePath);
            var sheetName = LoadSheetName(templatePath);

            var isValid = ValidateTemplate(templatePath,footer, sheetName);

            if (!isValid)
                return false;

            var cellReference = ReportCellMap.Get();
            var values = ConvertReportToDynamic(report);

            using var package = new ExcelPackage(new FileInfo(templatePath));
            var worksheet = package.Workbook.Worksheets["RAE"];

            PopulateWorksheet(worksheet, cellReference, values);

            worksheet.Calculate();

            package.SaveAs(new FileInfo(saveAsPath));
            return true;
        }

        private string LoadLeftFooter(string filePath)
        {
            if (!File.Exists(filePath))
                return string.Empty;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            return worksheet.HeaderFooter.OddFooter.LeftAlignedText;
        }

        private string LoadSheetName(string filePath)
        {
            if (!File.Exists(filePath))
                return string.Empty;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            return worksheet.Name;
        }

        private bool ValidateTemplate(string filePath, string footer, string sheetName)
        {
            bool fileExists = File.Exists(filePath);
            bool sheetNameIsValid = sheetName == "RAE";
            bool footerIsValid = !string.IsNullOrEmpty(footer);

            return fileExists && sheetNameIsValid && footerIsValid;
        }

        private dynamic[] ConvertReportToDynamic(Report report)
            {
                dynamic[] values = {
                    report.Request.Reference[1],
                    report.Request.Reference[2],
                    Convert.ToInt32(report.Request.Reference[3]),
                    report.Request.Reference[4],
                    report.Request.Reference[5],
                    report.Request.Reference[6],
                    report.Proposal.ProponenteNome,
                    report.Proposal.ProponenteCPF.Number,
                    report.Proposal.ProponenteDDD,
                    report.Proposal.ProponenteFone.Number,
                    report.Proposal.ResponsavelNome,
                    report.Proposal.ResponsavelCauCrea,
                    report.Proposal.ResponsavelUF,
                    report.Proposal.ResponsavelCPF.Number,
                    report.Proposal.ResponsavelDDD,
                    report.Proposal.ResponsavelFone.Number,
                    report.Proposal.ImovelEndereco,
                    report.Proposal.ImovelComplemento,
                    report.Proposal.ImovelBairro,
                    report.Proposal.ImovelCep.Number,
                    report.Proposal.ImovelMunicipio,
                    report.Proposal.ImovelUF,
                    report.Proposal.ImovelValorTerreno.Number,
                    report.Proposal.ImovelMatricula,
                    report.Proposal.ImovelOficio,
                    report.Proposal.ImovelComarca,
                    report.Proposal.ImovelComarcaUF,
                    Convert.ToDouble(report.Proposal.ServicoItem01),
                    Convert.ToDouble(report.Proposal.ServicoItem02),
                    Convert.ToDouble(report.Proposal.ServicoItem03),
                    Convert.ToDouble(report.Proposal.ServicoItem04),
                    Convert.ToDouble(report.Proposal.ServicoItem05),
                    Convert.ToDouble(report.Proposal.ServicoItem06),
                    Convert.ToDouble(report.Proposal.ServicoItem07),
                    Convert.ToDouble(report.Proposal.ServicoItem08),
                    Convert.ToDouble(report.Proposal.ServicoItem09),
                    Convert.ToDouble(report.Proposal.ServicoItem10),
                    Convert.ToDouble(report.Proposal.ServicoItem11),
                    Convert.ToDouble(report.Proposal.ServicoItem12),
                    Convert.ToDouble(report.Proposal.ServicoItem13),
                    Convert.ToDouble(report.Proposal.ServicoItem14),
                    Convert.ToDouble(report.Proposal.ServicoItem15),
                    Convert.ToDouble(report.Proposal.ServicoItem16),
                    Convert.ToDouble(report.Proposal.ServicoItem17),
                    Convert.ToDouble(report.Proposal.ServicoItem18),
                    Convert.ToDouble(report.Proposal.ServicoItem19),
                    Convert.ToDouble(report.Proposal.ServicoItem20),
                    report.MensuradoAcumulado == "" ? 0.00 : report.MensuradoAcumulado,
                    report.ContratoInicio == "  /  /" ? "" : Convert.ToDateTime(report.ContratoInicio),
                    report.ContratoTermino == "  /  /" ? "" : Convert.ToDateTime(report.ContratoTermino),
                    Convert.ToDouble(report.Proposal.CronogramaExecutado),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa1),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa2),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa3),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa4),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa5),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa6),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa7),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa8),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa9),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa10),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa11),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa12),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa13),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa14),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa15),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa16),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa17),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa18),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa19),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa20),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa21),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa22),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa23),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa24),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa25),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa26),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa27),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa28),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa29),
                    Convert.ToDouble(report.Proposal.CronogramaEtapa30)
                };

                return values;
            }

        private void PopulateWorksheet(ExcelWorksheet worksheet, string[] cellReference, dynamic[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (cellReference[i].StartsWith("AG") && Convert.ToDouble(values[i]) == 0.00)
                    continue;

                var cell = worksheet.Cells[cellReference[i]];

                if (values[i] is double || values[i] is int)
                {
                    cell.Value = Convert.ToDouble(values[i]);
                    cell.Style.Numberformat.Format = "0.00";
                }
                else
                {
                    cell.Value = values[i];
                }
            }
        }

    }
}