using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System.IO;
using System;
using AeX30.Domain.Entities;

namespace AeX30.Infra.Repository
{

    public class ProposalRepository
    {
        public Proposal GetProposal(string filePath, string[] cellReference)
        {
            try
            {
                FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                HSSFWorkbook wbook = new HSSFWorkbook(file);
                ISheet sheet = wbook.GetSheetAt(0);

                Proposal proposal = new Proposal(
                    tipo: wbook.GetSheetName(0),
                    vigencia: sheet.Footer.Left,
                    proponenteNome: sheet.GetRow(new CellReference(cellReference[00]).Row).GetCell(new CellReference(cellReference[00]).Col).ToString().TrimEnd(),
                    proponenteCPF: sheet.GetRow(new CellReference(cellReference[01]).Row).GetCell(new CellReference(cellReference[01]).Col).ToString(),
                    proponenteDDD: sheet.GetRow(new CellReference(cellReference[02]).Row).GetCell(new CellReference(cellReference[02]).Col).ToString(),
                    proponenteFone: sheet.GetRow(new CellReference(cellReference[03]).Row).GetCell(new CellReference(cellReference[03]).Col).ToString(),
                    responsavelNome: sheet.GetRow(new CellReference(cellReference[04]).Row).GetCell(new CellReference(cellReference[04]).Col).ToString(),
                    responsavelCauCrea: sheet.GetRow(new CellReference(cellReference[05]).Row).GetCell(new CellReference(cellReference[05]).Col).ToString(),
                    responsavelUF: sheet.GetRow(new CellReference(cellReference[06]).Row).GetCell(new CellReference(cellReference[06]).Col).ToString(),
                    responsavelCPF: sheet.GetRow(new CellReference(cellReference[07]).Row).GetCell(new CellReference(cellReference[07]).Col).ToString(),
                    responsavelDDD: sheet.GetRow(new CellReference(cellReference[08]).Row).GetCell(new CellReference(cellReference[08]).Col).ToString(),
                    responsavelFone: sheet.GetRow(new CellReference(cellReference[09]).Row).GetCell(new CellReference(cellReference[09]).Col).ToString(),
                    imovelEndereco: sheet.GetRow(new CellReference(cellReference[10]).Row).GetCell(new CellReference(cellReference[10]).Col).ToString(),
                    imovelComplemento: sheet.GetRow(new CellReference(cellReference[11]).Row).GetCell(new CellReference(cellReference[11]).Col).ToString(),
                    imovelCep: sheet.GetRow(new CellReference(cellReference[12]).Row).GetCell(new CellReference(cellReference[12]).Col).ToString(),
                    imovelBairro: sheet.GetRow(new CellReference(cellReference[13]).Row).GetCell(new CellReference(cellReference[13]).Col).ToString(),
                    imovelMunicipio: sheet.GetRow(new CellReference(cellReference[14]).Row).GetCell(new CellReference(cellReference[14]).Col).ToString(),
                    imovelUF: sheet.GetRow(new CellReference(cellReference[15]).Row).GetCell(new CellReference(cellReference[15]).Col).ToString(),
                    imovelValorTerreno: sheet.GetRow(new CellReference(cellReference[16]).Row).GetCell(new CellReference(cellReference[16]).Col).NumericCellValue.ToString(),
                    imovelMatricula: sheet.GetRow(new CellReference(cellReference[17]).Row).GetCell(new CellReference(cellReference[17]).Col).ToString(),
                    imovelOficio: sheet.GetRow(new CellReference(cellReference[18]).Row).GetCell(new CellReference(cellReference[18]).Col).ToString(),
                    imovelComarca: wbook.GetSheetName(0) == "Proposta" ? sheet.GetRow(new CellReference(cellReference[19]).Row).GetCell(new CellReference(cellReference[19]).Col).ToString() : "",
                    imovelComarcaUF: wbook.GetSheetName(0) == "Proposta" ? sheet.GetRow(new CellReference(cellReference[20]).Row).GetCell(new CellReference(cellReference[20]).Col).ToString() : "",
                    servicoItem01: sheet.GetRow(new CellReference(cellReference[21]).Row).GetCell(new CellReference(cellReference[21]).Col).NumericCellValue.ToString(),
                    servicoItem02: sheet.GetRow(new CellReference(cellReference[22]).Row).GetCell(new CellReference(cellReference[22]).Col).NumericCellValue.ToString(),
                    servicoItem03: sheet.GetRow(new CellReference(cellReference[23]).Row).GetCell(new CellReference(cellReference[23]).Col).NumericCellValue.ToString(),
                    servicoItem04: sheet.GetRow(new CellReference(cellReference[24]).Row).GetCell(new CellReference(cellReference[24]).Col).NumericCellValue.ToString(),
                    servicoItem05: sheet.GetRow(new CellReference(cellReference[25]).Row).GetCell(new CellReference(cellReference[25]).Col).NumericCellValue.ToString(),
                    servicoItem06: sheet.GetRow(new CellReference(cellReference[26]).Row).GetCell(new CellReference(cellReference[26]).Col).NumericCellValue.ToString(),
                    servicoItem07: sheet.GetRow(new CellReference(cellReference[27]).Row).GetCell(new CellReference(cellReference[27]).Col).NumericCellValue.ToString(),
                    servicoItem08: sheet.GetRow(new CellReference(cellReference[28]).Row).GetCell(new CellReference(cellReference[28]).Col).NumericCellValue.ToString(),
                    servicoItem09: sheet.GetRow(new CellReference(cellReference[29]).Row).GetCell(new CellReference(cellReference[29]).Col).NumericCellValue.ToString(),
                    servicoItem10: sheet.GetRow(new CellReference(cellReference[30]).Row).GetCell(new CellReference(cellReference[30]).Col).NumericCellValue.ToString(),
                    servicoItem11: sheet.GetRow(new CellReference(cellReference[31]).Row).GetCell(new CellReference(cellReference[31]).Col).NumericCellValue.ToString(),
                    servicoItem12: sheet.GetRow(new CellReference(cellReference[32]).Row).GetCell(new CellReference(cellReference[32]).Col).NumericCellValue.ToString(),
                    servicoItem13: sheet.GetRow(new CellReference(cellReference[33]).Row).GetCell(new CellReference(cellReference[33]).Col).NumericCellValue.ToString(),
                    servicoItem14: sheet.GetRow(new CellReference(cellReference[34]).Row).GetCell(new CellReference(cellReference[34]).Col).NumericCellValue.ToString(),
                    servicoItem15: sheet.GetRow(new CellReference(cellReference[35]).Row).GetCell(new CellReference(cellReference[35]).Col).NumericCellValue.ToString(),
                    servicoItem16: sheet.GetRow(new CellReference(cellReference[36]).Row).GetCell(new CellReference(cellReference[36]).Col).NumericCellValue.ToString(),
                    servicoItem17: sheet.GetRow(new CellReference(cellReference[37]).Row).GetCell(new CellReference(cellReference[37]).Col).NumericCellValue.ToString(),
                    servicoItem18: sheet.GetRow(new CellReference(cellReference[38]).Row).GetCell(new CellReference(cellReference[38]).Col).NumericCellValue.ToString(),
                    servicoItem19: sheet.GetRow(new CellReference(cellReference[39]).Row).GetCell(new CellReference(cellReference[39]).Col).NumericCellValue.ToString(),
                    servicoItem20: sheet.GetRow(new CellReference(cellReference[40]).Row).GetCell(new CellReference(cellReference[40]).Col).NumericCellValue.ToString(),
                    cronogramaExecutado: sheet.GetRow(new CellReference(cellReference[41]).Row).GetCell(new CellReference(cellReference[41]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa1: sheet.GetRow(new CellReference(cellReference[42]).Row).GetCell(new CellReference(cellReference[42]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa2: sheet.GetRow(new CellReference(cellReference[43]).Row).GetCell(new CellReference(cellReference[43]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa3: sheet.GetRow(new CellReference(cellReference[44]).Row).GetCell(new CellReference(cellReference[44]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa4: sheet.GetRow(new CellReference(cellReference[45]).Row).GetCell(new CellReference(cellReference[45]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa5: sheet.GetRow(new CellReference(cellReference[46]).Row).GetCell(new CellReference(cellReference[46]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa6: sheet.GetRow(new CellReference(cellReference[47]).Row).GetCell(new CellReference(cellReference[47]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa7: sheet.GetRow(new CellReference(cellReference[48]).Row).GetCell(new CellReference(cellReference[48]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa8: sheet.GetRow(new CellReference(cellReference[49]).Row).GetCell(new CellReference(cellReference[49]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa9: sheet.GetRow(new CellReference(cellReference[50]).Row).GetCell(new CellReference(cellReference[50]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa10: sheet.GetRow(new CellReference(cellReference[51]).Row).GetCell(new CellReference(cellReference[52]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa11: sheet.GetRow(new CellReference(cellReference[52]).Row).GetCell(new CellReference(cellReference[52]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa12: sheet.GetRow(new CellReference(cellReference[53]).Row).GetCell(new CellReference(cellReference[53]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa13: sheet.GetRow(new CellReference(cellReference[54]).Row).GetCell(new CellReference(cellReference[54]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa14: sheet.GetRow(new CellReference(cellReference[55]).Row).GetCell(new CellReference(cellReference[55]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa15: sheet.GetRow(new CellReference(cellReference[56]).Row).GetCell(new CellReference(cellReference[56]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa16: sheet.GetRow(new CellReference(cellReference[57]).Row).GetCell(new CellReference(cellReference[57]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa17: sheet.GetRow(new CellReference(cellReference[58]).Row).GetCell(new CellReference(cellReference[58]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa18: sheet.GetRow(new CellReference(cellReference[59]).Row).GetCell(new CellReference(cellReference[59]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa19: sheet.GetRow(new CellReference(cellReference[60]).Row).GetCell(new CellReference(cellReference[60]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa20: sheet.GetRow(new CellReference(cellReference[61]).Row).GetCell(new CellReference(cellReference[61]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa21: sheet.GetRow(new CellReference(cellReference[62]).Row).GetCell(new CellReference(cellReference[62]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa22: sheet.GetRow(new CellReference(cellReference[63]).Row).GetCell(new CellReference(cellReference[63]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa23: sheet.GetRow(new CellReference(cellReference[64]).Row).GetCell(new CellReference(cellReference[64]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa24: sheet.GetRow(new CellReference(cellReference[65]).Row).GetCell(new CellReference(cellReference[65]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa25: sheet.GetRow(new CellReference(cellReference[66]).Row).GetCell(new CellReference(cellReference[66]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa26: sheet.GetRow(new CellReference(cellReference[67]).Row).GetCell(new CellReference(cellReference[67]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa27: sheet.GetRow(new CellReference(cellReference[68]).Row).GetCell(new CellReference(cellReference[68]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa28: sheet.GetRow(new CellReference(cellReference[69]).Row).GetCell(new CellReference(cellReference[69]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa29: sheet.GetRow(new CellReference(cellReference[70]).Row).GetCell(new CellReference(cellReference[70]).Col).NumericCellValue.ToString(),
                    cronogramaEtapa30: sheet.GetRow(new CellReference(cellReference[71]).Row).GetCell(new CellReference(cellReference[71]).Col).NumericCellValue.ToString()
                );
                return proposal;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

        }

    }
}
