using AeX30.Core.Entities;
using System.IO;
using AeX30.Core.ValueObject;
using System;
using OfficeOpenXml;
using AeX30.Core.Interfaces;

namespace AeX30.Core.Services
{
    public class ProposalService : IProposalService
    {
        public Proposal LoadFromFile(string filePath)
        {
            var footer = LoadLeftFooter(filePath);
            var sheetName = LoadSheetName(filePath);
            var cellReference = ProposalCellMap.Get(footer);

            var isValid = ValidateLoadingFile(filePath, sheetName, footer, cellReference);

            if (!isValid)
                return null;

            try
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                var sheet = package.Workbook.Worksheets[0];

                var proposal = new Proposal(
                    tipo: sheetName,
                    vigencia: footer,
                    proponenteNome: sheet.Cells[cellReference[0]].Text.Trim(),
                    proponenteCPF: new Cpf(sheet.Cells[cellReference[1]].Text),
                    proponenteDDD: sheet.Cells[cellReference[2]].Text,
                    proponenteFone: new PhoneNumber(sheet.Cells[cellReference[3]].Text),
                    responsavelNome: sheet.Cells[cellReference[4]].Text,
                    responsavelCauCrea: sheet.Cells[cellReference[5]].Text,
                    responsavelUF: sheet.Cells[cellReference[6]].Text,
                    responsavelCPF: new Cpf(sheet.Cells[cellReference[7]].Text),
                    responsavelDDD: sheet.Cells[cellReference[8]].Text,
                    responsavelFone: new PhoneNumber(sheet.Cells[cellReference[9]].Text),
                    imovelEndereco: sheet.Cells[cellReference[10]].Text,
                    imovelComplemento: sheet.Cells[cellReference[11]].Text,
                    imovelCep: new ZipCode(sheet.Cells[cellReference[12]].Text),
                    imovelBairro: sheet.Cells[cellReference[13]].Text,
                    imovelMunicipio: sheet.Cells[cellReference[14]].Text,
                    imovelUF: sheet.Cells[cellReference[15]].Text,
                    imovelValorTerreno: new Money(sheet.Cells[cellReference[16]].GetValue<double>().ToString()),
                    imovelMatricula: sheet.Cells[cellReference[17]].Text,
                    imovelOficio: sheet.Cells[cellReference[18]].Text,
                    imovelComarca: string.Empty,
                    imovelComarcaUF: string.Empty,
                    servicoItem01: sheet.Cells[cellReference[21]].GetValue<double>().ToString(),
                    servicoItem02: sheet.Cells[cellReference[22]].GetValue<double>().ToString(),
                    servicoItem03: sheet.Cells[cellReference[23]].GetValue<double>().ToString(),
                    servicoItem04: sheet.Cells[cellReference[24]].GetValue<double>().ToString(),
                    servicoItem05: sheet.Cells[cellReference[25]].GetValue<double>().ToString(),
                    servicoItem06: sheet.Cells[cellReference[26]].GetValue<double>().ToString(),
                    servicoItem07: sheet.Cells[cellReference[27]].GetValue<double>().ToString(),
                    servicoItem08: sheet.Cells[cellReference[28]].GetValue<double>().ToString(),
                    servicoItem09: sheet.Cells[cellReference[29]].GetValue<double>().ToString(),
                    servicoItem10: sheet.Cells[cellReference[30]].GetValue<double>().ToString(),
                    servicoItem11: sheet.Cells[cellReference[31]].GetValue<double>().ToString(),
                    servicoItem12: sheet.Cells[cellReference[32]].GetValue<double>().ToString(),
                    servicoItem13: sheet.Cells[cellReference[33]].GetValue<double>().ToString(),
                    servicoItem14: sheet.Cells[cellReference[34]].GetValue<double>().ToString(),
                    servicoItem15: sheet.Cells[cellReference[35]].GetValue<double>().ToString(),
                    servicoItem16: sheet.Cells[cellReference[36]].GetValue<double>().ToString(),
                    servicoItem17: sheet.Cells[cellReference[37]].GetValue<double>().ToString(),
                    servicoItem18: sheet.Cells[cellReference[38]].GetValue<double>().ToString(),
                    servicoItem19: sheet.Cells[cellReference[39]].GetValue<double>().ToString(),
                    servicoItem20: sheet.Cells[cellReference[40]].GetValue<double>().ToString(),
                    cronogramaExecutado: sheet.Cells[cellReference[41]].GetValue<double>().ToString(),
                    cronogramaEtapa1: sheet.Cells[cellReference[42]].GetValue<double>().ToString(),
                    cronogramaEtapa2: sheet.Cells[cellReference[43]].GetValue<double>().ToString(),
                    cronogramaEtapa3: sheet.Cells[cellReference[44]].GetValue<double>().ToString(),
                    cronogramaEtapa4: sheet.Cells[cellReference[45]].GetValue<double>().ToString(),
                    cronogramaEtapa5: sheet.Cells[cellReference[46]].GetValue<double>().ToString(),
                    cronogramaEtapa6: sheet.Cells[cellReference[47]].GetValue<double>().ToString(),
                    cronogramaEtapa7: sheet.Cells[cellReference[48]].GetValue<double>().ToString(),
                    cronogramaEtapa8: sheet.Cells[cellReference[49]].GetValue<double>().ToString(),
                    cronogramaEtapa9: sheet.Cells[cellReference[50]].GetValue<double>().ToString(),
                    cronogramaEtapa10: sheet.Cells[cellReference[51]].GetValue<double>().ToString(),
                    cronogramaEtapa11: sheet.Cells[cellReference[52]].GetValue<double>().ToString(),
                    cronogramaEtapa12: sheet.Cells[cellReference[53]].GetValue<double>().ToString(),
                    cronogramaEtapa13: sheet.Cells[cellReference[54]].GetValue<double>().ToString(),
                    cronogramaEtapa14: sheet.Cells[cellReference[55]].GetValue<double>().ToString(),
                    cronogramaEtapa15: sheet.Cells[cellReference[56]].GetValue<double>().ToString(),
                    cronogramaEtapa16: sheet.Cells[cellReference[57]].GetValue<double>().ToString(),
                    cronogramaEtapa17: sheet.Cells[cellReference[58]].GetValue<double>().ToString(),
                    cronogramaEtapa18: sheet.Cells[cellReference[59]].GetValue<double>().ToString(),
                    cronogramaEtapa19: sheet.Cells[cellReference[60]].GetValue<double>().ToString(),
                    cronogramaEtapa20: sheet.Cells[cellReference[61]].GetValue<double>().ToString(),
                    cronogramaEtapa21: sheet.Cells[cellReference[62]].GetValue<double>().ToString(),
                    cronogramaEtapa22: sheet.Cells[cellReference[63]].GetValue<double>().ToString(),
                    cronogramaEtapa23: sheet.Cells[cellReference[64]].GetValue<double>().ToString(),
                    cronogramaEtapa24: sheet.Cells[cellReference[65]].GetValue<double>().ToString(),
                    cronogramaEtapa25: sheet.Cells[cellReference[66]].GetValue<double>().ToString(),
                    cronogramaEtapa26: sheet.Cells[cellReference[67]].GetValue<double>().ToString(),
                    cronogramaEtapa27: sheet.Cells[cellReference[68]].GetValue<double>().ToString(),
                    cronogramaEtapa28: sheet.Cells[cellReference[69]].GetValue<double>().ToString(),
                    cronogramaEtapa29: sheet.Cells[cellReference[70]].GetValue<double>().ToString(),
                    cronogramaEtapa30: sheet.Cells[cellReference[71]].GetValue<double>().ToString()
                );

                return proposal;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
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

        private bool ValidateLoadingFile(string filePath, string sheetName, string footer, string[] cellReference)
        {
            if (!File.Exists(filePath))
                return false;

            bool sheetNameIsValid = sheetName == "Proposta_Constr_Individual";
            bool footerIsValid = !string.IsNullOrEmpty(footer);
            bool cellReferenceIsValid = cellReference != null;

            return sheetNameIsValid && footerIsValid && cellReferenceIsValid;
        }
    }
}