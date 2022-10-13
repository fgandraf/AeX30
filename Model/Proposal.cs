﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO;
using System;

namespace AeX30.Model
{

    public class Proposal
    {
        public string Tipo { get; set; }
        public string Vigencia { get; set; }
        public string ProponenteNome { get; set; }
        public string ProponenteCPF { get; set; }
        public string ProponenteDDD { get; set; }
        public string ProponenteFone { get; set; }
        public string ResponsavelNome { get; set; }
        public string ReponsavelCauCrea { get; set; }
        public string ResponsavelUF { get; set; }
        public string ResponsavelCPF { get; set; }
        public string ResponsavelDDD { get; set; }
        public string ResponsavelFone { get; set; }
        public string ImovelEndereco { get; set; }
        public string ImovelComplemento { get; set; }
        public string ImovelCep { get; set; }
        public string ImovelBairro { get; set; }
        public string ImovelMunicipio { get; set; }
        public string ImovelUF { get; set; }
        public string ImovelValorTerreno { get; set; }
        public string ImovelMatricula { get; set; }
        public string ImovelOficio { get; set; }
        public string ImovelComarca { get; set; }
        public string ImovelComarcaUF { get; set; }
        public string ServicoItem01 { get; set; }
        public string ServicoItem02 { get; set; }
        public string ServicoItem03 { get; set; }
        public string ServicoItem04 { get; set; }
        public string ServicoItem05 { get; set; }
        public string ServicoItem06 { get; set; }
        public string ServicoItem07 { get; set; }
        public string ServicoItem08 { get; set; }
        public string ServicoItem09 { get; set; }
        public string ServicoItem10 { get; set; }
        public string ServicoItem11 { get; set; }
        public string ServicoItem12 { get; set; }
        public string ServicoItem13 { get; set; }
        public string ServicoItem14 { get; set; }
        public string ServicoItem15 { get; set; }
        public string ServicoItem16 { get; set; }
        public string ServicoItem17 { get; set; }
        public string ServicoItem18 { get; set; }
        public string ServicoItem19 { get; set; }
        public string ServicoItem20 { get; set; }
        public string CronogramaExecutado { get; set; }
        public string CronogramaEtapa1 { get; set; } = "0";
        public string CronogramaEtapa2 { get; set; } = "0";
        public string CronogramaEtapa3 { get; set; } = "0";
        public string CronogramaEtapa4 { get; set; } = "0";
        public string CronogramaEtapa5 { get; set; } = "0";
        public string CronogramaEtapa6 { get; set; } = "0";
        public string CronogramaEtapa7 { get; set; } = "0";
        public string CronogramaEtapa8 { get; set; } = "0";
        public string CronogramaEtapa9 { get; set; } = "0";
        public string CronogramaEtapa10 { get; set; } = "0";
        public string CronogramaEtapa11 { get; set; } = "0";
        public string CronogramaEtapa12 { get; set; } = "0";
        public string CronogramaEtapa13 { get; set; } = "0";
        public string CronogramaEtapa14 { get; set; } = "0";
        public string CronogramaEtapa15 { get; set; } = "0";
        public string CronogramaEtapa16 { get; set; } = "0";
        public string CronogramaEtapa17 { get; set; } = "0";
        public string CronogramaEtapa18 { get; set; } = "0";
        public string CronogramaEtapa19 { get; set; } = "0";
        public string CronogramaEtapa20 { get; set; } = "0";
        public string CronogramaEtapa21 { get; set; } = "0";
        public string CronogramaEtapa22 { get; set; } = "0";
        public string CronogramaEtapa23 { get; set; } = "0";
        public string CronogramaEtapa24 { get; set; } = "0";
        public string CronogramaEtapa25 { get; set; } = "0";
        public string CronogramaEtapa26 { get; set; } = "0";
        public string CronogramaEtapa27 { get; set; } = "0";
        public string CronogramaEtapa28 { get; set; } = "0";
        public string CronogramaEtapa29 { get; set; } = "0";
        public string CronogramaEtapa30 { get; set; } = "0";



        public static string GetSheetName(string filePath)
        {
            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);

            return wbook.GetSheetName(0);
        }

        public static string GetLeftFooter(string filePath)
        {
            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);
            ISheet sheet = wbook.GetSheet(wbook.GetSheetName(0));

            return sheet.Footer.Left;
        }

        public Proposal GetProposal(string filePath, string[] cellReference)
        {
            try
            {
                FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                HSSFWorkbook wbook = new HSSFWorkbook(file);
                ISheet sheet = wbook.GetSheetAt(0);

                Proposal proposal = new Proposal
                {
                    Vigencia = sheet.Footer.Left,
                    Tipo = wbook.GetSheetName(0) == "Proposta" ? "PFUI" : "PCI",
                    ProponenteNome = sheet.GetRow(new CellReference(cellReference[00]).Row).GetCell(new CellReference(cellReference[00]).Col).ToString().TrimEnd(),
                    ProponenteCPF = FormatString.CPF(sheet.GetRow(new CellReference(cellReference[01]).Row).GetCell(new CellReference(cellReference[01]).Col).ToString()),
                    ProponenteDDD = sheet.GetRow(new CellReference(cellReference[02]).Row).GetCell(new CellReference(cellReference[02]).Col).ToString(),
                    ProponenteFone = FormatString.Fone(sheet.GetRow(new CellReference(cellReference[03]).Row).GetCell(new CellReference(cellReference[03]).Col).ToString()),
                    ResponsavelNome = sheet.GetRow(new CellReference(cellReference[04]).Row).GetCell(new CellReference(cellReference[04]).Col).ToString(),
                    ReponsavelCauCrea = sheet.GetRow(new CellReference(cellReference[05]).Row).GetCell(new CellReference(cellReference[05]).Col).ToString(),
                    ResponsavelUF = sheet.GetRow(new CellReference(cellReference[06]).Row).GetCell(new CellReference(cellReference[06]).Col).ToString(),
                    ResponsavelCPF = FormatString.CPF(sheet.GetRow(new CellReference(cellReference[07]).Row).GetCell(new CellReference(cellReference[07]).Col).ToString()),
                    ResponsavelDDD = sheet.GetRow(new CellReference(cellReference[08]).Row).GetCell(new CellReference(cellReference[08]).Col).ToString(),
                    ResponsavelFone = FormatString.Fone(sheet.GetRow(new CellReference(cellReference[09]).Row).GetCell(new CellReference(cellReference[09]).Col).ToString()),
                    ImovelEndereco = sheet.GetRow(new CellReference(cellReference[10]).Row).GetCell(new CellReference(cellReference[10]).Col).ToString(),
                    ImovelComplemento = sheet.GetRow(new CellReference(cellReference[11]).Row).GetCell(new CellReference(cellReference[11]).Col).ToString(),
                    ImovelCep = FormatString.CEP(sheet.GetRow(new CellReference(cellReference[12]).Row).GetCell(new CellReference(cellReference[12]).Col).ToString()),
                    ImovelBairro = sheet.GetRow(new CellReference(cellReference[13]).Row).GetCell(new CellReference(cellReference[13]).Col).ToString(),
                    ImovelMunicipio = sheet.GetRow(new CellReference(cellReference[14]).Row).GetCell(new CellReference(cellReference[14]).Col).ToString(),
                    ImovelUF = sheet.GetRow(new CellReference(cellReference[15]).Row).GetCell(new CellReference(cellReference[15]).Col).ToString(),
                    ImovelValorTerreno = sheet.GetRow(new CellReference(cellReference[16]).Row).GetCell(new CellReference(cellReference[16]).Col).ToString(),
                    ImovelMatricula = sheet.GetRow(new CellReference(cellReference[17]).Row).GetCell(new CellReference(cellReference[17]).Col).ToString(),
                    ImovelOficio = sheet.GetRow(new CellReference(cellReference[18]).Row).GetCell(new CellReference(cellReference[18]).Col).ToString(),
                    ImovelComarca = wbook.GetSheetName(0) == "Proposta" ? sheet.GetRow(new CellReference(cellReference[19]).Row).GetCell(new CellReference(cellReference[19]).Col).ToString() : "",
                    ImovelComarcaUF = wbook.GetSheetName(0) == "Proposta" ? sheet.GetRow(new CellReference(cellReference[20]).Row).GetCell(new CellReference(cellReference[20]).Col).ToString() : "",
                    ServicoItem01 = sheet.GetRow(new CellReference(cellReference[21]).Row).GetCell(new CellReference(cellReference[21]).Col).NumericCellValue.ToString(),
                    ServicoItem02 = sheet.GetRow(new CellReference(cellReference[22]).Row).GetCell(new CellReference(cellReference[22]).Col).NumericCellValue.ToString(),
                    ServicoItem03 = sheet.GetRow(new CellReference(cellReference[23]).Row).GetCell(new CellReference(cellReference[23]).Col).NumericCellValue.ToString(),
                    ServicoItem04 = sheet.GetRow(new CellReference(cellReference[24]).Row).GetCell(new CellReference(cellReference[24]).Col).NumericCellValue.ToString(),
                    ServicoItem05 = sheet.GetRow(new CellReference(cellReference[25]).Row).GetCell(new CellReference(cellReference[25]).Col).NumericCellValue.ToString(),
                    ServicoItem06 = sheet.GetRow(new CellReference(cellReference[26]).Row).GetCell(new CellReference(cellReference[26]).Col).NumericCellValue.ToString(),
                    ServicoItem07 = sheet.GetRow(new CellReference(cellReference[27]).Row).GetCell(new CellReference(cellReference[27]).Col).NumericCellValue.ToString(),
                    ServicoItem08 = sheet.GetRow(new CellReference(cellReference[28]).Row).GetCell(new CellReference(cellReference[28]).Col).NumericCellValue.ToString(),
                    ServicoItem09 = sheet.GetRow(new CellReference(cellReference[29]).Row).GetCell(new CellReference(cellReference[29]).Col).NumericCellValue.ToString(),
                    ServicoItem10 = sheet.GetRow(new CellReference(cellReference[30]).Row).GetCell(new CellReference(cellReference[30]).Col).NumericCellValue.ToString(),
                    ServicoItem11 = sheet.GetRow(new CellReference(cellReference[31]).Row).GetCell(new CellReference(cellReference[31]).Col).NumericCellValue.ToString(),
                    ServicoItem12 = sheet.GetRow(new CellReference(cellReference[32]).Row).GetCell(new CellReference(cellReference[32]).Col).NumericCellValue.ToString(),
                    ServicoItem13 = sheet.GetRow(new CellReference(cellReference[33]).Row).GetCell(new CellReference(cellReference[33]).Col).NumericCellValue.ToString(),
                    ServicoItem14 = sheet.GetRow(new CellReference(cellReference[34]).Row).GetCell(new CellReference(cellReference[34]).Col).NumericCellValue.ToString(),
                    ServicoItem15 = sheet.GetRow(new CellReference(cellReference[35]).Row).GetCell(new CellReference(cellReference[35]).Col).NumericCellValue.ToString(),
                    ServicoItem16 = sheet.GetRow(new CellReference(cellReference[36]).Row).GetCell(new CellReference(cellReference[36]).Col).NumericCellValue.ToString(),
                    ServicoItem17 = sheet.GetRow(new CellReference(cellReference[37]).Row).GetCell(new CellReference(cellReference[37]).Col).NumericCellValue.ToString(),
                    ServicoItem18 = sheet.GetRow(new CellReference(cellReference[38]).Row).GetCell(new CellReference(cellReference[38]).Col).NumericCellValue.ToString(),
                    ServicoItem19 = sheet.GetRow(new CellReference(cellReference[39]).Row).GetCell(new CellReference(cellReference[39]).Col).NumericCellValue.ToString(),
                    ServicoItem20 = sheet.GetRow(new CellReference(cellReference[40]).Row).GetCell(new CellReference(cellReference[40]).Col).NumericCellValue.ToString(),
                    CronogramaExecutado = sheet.GetRow(new CellReference(cellReference[41]).Row).GetCell(new CellReference(cellReference[41]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa1 = sheet.GetRow(new CellReference(cellReference[42]).Row).GetCell(new CellReference(cellReference[42]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa2 = sheet.GetRow(new CellReference(cellReference[43]).Row).GetCell(new CellReference(cellReference[43]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa3 = sheet.GetRow(new CellReference(cellReference[44]).Row).GetCell(new CellReference(cellReference[44]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa4 = sheet.GetRow(new CellReference(cellReference[45]).Row).GetCell(new CellReference(cellReference[45]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa5 = sheet.GetRow(new CellReference(cellReference[46]).Row).GetCell(new CellReference(cellReference[46]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa6 = sheet.GetRow(new CellReference(cellReference[47]).Row).GetCell(new CellReference(cellReference[47]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa7 = sheet.GetRow(new CellReference(cellReference[48]).Row).GetCell(new CellReference(cellReference[48]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa8 = sheet.GetRow(new CellReference(cellReference[49]).Row).GetCell(new CellReference(cellReference[49]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa9 = sheet.GetRow(new CellReference(cellReference[50]).Row).GetCell(new CellReference(cellReference[50]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa10 = sheet.GetRow(new CellReference(cellReference[51]).Row).GetCell(new CellReference(cellReference[52]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa11 = sheet.GetRow(new CellReference(cellReference[52]).Row).GetCell(new CellReference(cellReference[52]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa12 = sheet.GetRow(new CellReference(cellReference[53]).Row).GetCell(new CellReference(cellReference[53]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa13 = sheet.GetRow(new CellReference(cellReference[54]).Row).GetCell(new CellReference(cellReference[54]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa14 = sheet.GetRow(new CellReference(cellReference[55]).Row).GetCell(new CellReference(cellReference[55]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa15 = sheet.GetRow(new CellReference(cellReference[56]).Row).GetCell(new CellReference(cellReference[56]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa16 = sheet.GetRow(new CellReference(cellReference[57]).Row).GetCell(new CellReference(cellReference[57]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa17 = sheet.GetRow(new CellReference(cellReference[58]).Row).GetCell(new CellReference(cellReference[58]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa18 = sheet.GetRow(new CellReference(cellReference[59]).Row).GetCell(new CellReference(cellReference[59]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa19 = sheet.GetRow(new CellReference(cellReference[60]).Row).GetCell(new CellReference(cellReference[60]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa20 = sheet.GetRow(new CellReference(cellReference[61]).Row).GetCell(new CellReference(cellReference[61]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa21 = sheet.GetRow(new CellReference(cellReference[62]).Row).GetCell(new CellReference(cellReference[62]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa22 = sheet.GetRow(new CellReference(cellReference[63]).Row).GetCell(new CellReference(cellReference[63]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa23 = sheet.GetRow(new CellReference(cellReference[64]).Row).GetCell(new CellReference(cellReference[64]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa24 = sheet.GetRow(new CellReference(cellReference[65]).Row).GetCell(new CellReference(cellReference[65]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa25 = sheet.GetRow(new CellReference(cellReference[66]).Row).GetCell(new CellReference(cellReference[66]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa26 = sheet.GetRow(new CellReference(cellReference[67]).Row).GetCell(new CellReference(cellReference[67]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa27 = sheet.GetRow(new CellReference(cellReference[68]).Row).GetCell(new CellReference(cellReference[68]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa28 = sheet.GetRow(new CellReference(cellReference[69]).Row).GetCell(new CellReference(cellReference[69]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa29 = sheet.GetRow(new CellReference(cellReference[70]).Row).GetCell(new CellReference(cellReference[70]).Col).NumericCellValue.ToString(),
                    CronogramaEtapa30 = sheet.GetRow(new CellReference(cellReference[71]).Row).GetCell(new CellReference(cellReference[71]).Col).NumericCellValue.ToString()
                };
                return proposal;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

    }
}