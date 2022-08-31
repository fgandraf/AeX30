using System;
using System.IO;
using aeX30.Entities;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace aeX30.Model
{
    public class PropostaModel
    {

        internal static string GetFooter(string path)
        {
            FileStream arquivoXLS = new FileStream(path, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(arquivoXLS);

            ISheet sheet = wbook.GetSheet(wbook.GetSheetName(0));
            return sheet.Footer.Left;
        }

        internal static string GetType(string path)
        {
            FileStream arquivoXLS = new FileStream(path, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(arquivoXLS);

            if (wbook.GetSheetName(0) == "Proposta")
                return "PFUI";
            else if (wbook.GetSheetName(0) == "Proposta_Constr_Individual")
                return "PCI";
            else
                return null;
        }

        internal PropostaEnt GetPFUI(string path, string[] aeX)
        {
            FileStream arquivoXLS = new FileStream(path, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(arquivoXLS);
            ISheet ws = wbook.GetSheet("Proposta");

            PropostaEnt pfui = new PropostaEnt();
            //CABEÇALHO
            pfui.Tipo = "PFUI";
            pfui.Vigencia = ws.Footer.Left;
            pfui.Prop_Nome = ws.GetRow(new CellReference(aeX[0]).Row).GetCell(new CellReference(aeX[0]).Col).ToString();
            pfui.Prop_CPF = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[1]).Row).GetCell(new CellReference(aeX[1]).Col).ToString());
            pfui.Prop_DDD = ws.GetRow(new CellReference(aeX[2]).Row).GetCell(new CellReference(aeX[2]).Col).ToString();
            pfui.Prop_Telefone = Util.FormatedFone(ws.GetRow(new CellReference(aeX[3]).Row).GetCell(new CellReference(aeX[3]).Col).ToString());
            pfui.Rt_Nome = ws.GetRow(new CellReference(aeX[4]).Row).GetCell(new CellReference(aeX[4]).Col).ToString();
            pfui.Rt_CAU_CREA = ws.GetRow(new CellReference(aeX[5]).Row).GetCell(new CellReference(aeX[5]).Col).ToString();
            pfui.Rt_UF = ws.GetRow(new CellReference(aeX[6]).Row).GetCell(new CellReference(aeX[6]).Col).ToString();
            pfui.Rt_CPF = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[7]).Row).GetCell(new CellReference(aeX[7]).Col).ToString());
            pfui.Rt_DDD = ws.GetRow(new CellReference(aeX[8]).Row).GetCell(new CellReference(aeX[8]).Col).ToString();
            pfui.Rt_Telefone = Util.FormatedFone(ws.GetRow(new CellReference(aeX[9]).Row).GetCell(new CellReference(aeX[9]).Col).ToString());
            pfui.End_Endereço = ws.GetRow(new CellReference(aeX[10]).Row).GetCell(new CellReference(aeX[10]).Col).ToString();
            pfui.End_Complemento = ws.GetRow(new CellReference(aeX[11]).Row).GetCell(new CellReference(aeX[11]).Col).ToString();
            pfui.End_CEP = Util.FormatedCEP(ws.GetRow(new CellReference(aeX[12]).Row).GetCell(new CellReference(aeX[12]).Col).ToString());
            pfui.End_Bairro = ws.GetRow(new CellReference(aeX[13]).Row).GetCell(new CellReference(aeX[13]).Col).ToString();
            pfui.End_Municipio = ws.GetRow(new CellReference(aeX[14]).Row).GetCell(new CellReference(aeX[14]).Col).ToString();
            pfui.End_UF = ws.GetRow(new CellReference(aeX[15]).Row).GetCell(new CellReference(aeX[15]).Col).ToString();
            pfui.Imov_Valor_Terreno = string.Format("{0:0,0.00}", Convert.ToDouble(ws.GetRow(new CellReference(aeX[16]).Row).GetCell(new CellReference(aeX[16]).Col).ToString()));
            pfui.Imov_Matricula = ws.GetRow(new CellReference(aeX[17]).Row).GetCell(new CellReference(aeX[17]).Col).ToString();
            pfui.Imov_Oficio = ws.GetRow(new CellReference(aeX[18]).Row).GetCell(new CellReference(aeX[18]).Col).ToString();
            pfui.Imov_Comarca = ws.GetRow(new CellReference(aeX[19]).Row).GetCell(new CellReference(aeX[19]).Col).ToString();
            pfui.Imov_UF = ws.GetRow(new CellReference(aeX[20]).Row).GetCell(new CellReference(aeX[20]).Col).ToString();
            //ORÇAMENTO (PERCENTUAIS)
            pfui.Item_17_01 = ws.GetRow(new CellReference(aeX[21]).Row).GetCell(new CellReference(aeX[21]).Col).NumericCellValue.ToString();
            pfui.Item_17_02 = ws.GetRow(new CellReference(aeX[22]).Row).GetCell(new CellReference(aeX[22]).Col).NumericCellValue.ToString();
            pfui.Item_17_03 = ws.GetRow(new CellReference(aeX[23]).Row).GetCell(new CellReference(aeX[23]).Col).NumericCellValue.ToString();
            pfui.Item_17_04 = ws.GetRow(new CellReference(aeX[24]).Row).GetCell(new CellReference(aeX[24]).Col).NumericCellValue.ToString();
            pfui.Item_17_05 = ws.GetRow(new CellReference(aeX[25]).Row).GetCell(new CellReference(aeX[25]).Col).NumericCellValue.ToString();
            pfui.Item_17_06 = ws.GetRow(new CellReference(aeX[26]).Row).GetCell(new CellReference(aeX[26]).Col).NumericCellValue.ToString();
            pfui.Item_17_07 = ws.GetRow(new CellReference(aeX[27]).Row).GetCell(new CellReference(aeX[27]).Col).NumericCellValue.ToString();
            pfui.Item_17_08 = ws.GetRow(new CellReference(aeX[28]).Row).GetCell(new CellReference(aeX[28]).Col).NumericCellValue.ToString();
            pfui.Item_17_09 = ws.GetRow(new CellReference(aeX[29]).Row).GetCell(new CellReference(aeX[29]).Col).NumericCellValue.ToString();
            pfui.Item_17_10 = ws.GetRow(new CellReference(aeX[30]).Row).GetCell(new CellReference(aeX[30]).Col).NumericCellValue.ToString();
            pfui.Item_17_11 = ws.GetRow(new CellReference(aeX[31]).Row).GetCell(new CellReference(aeX[31]).Col).NumericCellValue.ToString();
            pfui.Item_17_12 = ws.GetRow(new CellReference(aeX[32]).Row).GetCell(new CellReference(aeX[32]).Col).NumericCellValue.ToString();
            pfui.Item_17_13 = ws.GetRow(new CellReference(aeX[33]).Row).GetCell(new CellReference(aeX[33]).Col).NumericCellValue.ToString();
            pfui.Item_17_14 = ws.GetRow(new CellReference(aeX[34]).Row).GetCell(new CellReference(aeX[34]).Col).NumericCellValue.ToString();
            pfui.Item_17_15 = ws.GetRow(new CellReference(aeX[35]).Row).GetCell(new CellReference(aeX[35]).Col).NumericCellValue.ToString();
            pfui.Item_17_16 = ws.GetRow(new CellReference(aeX[36]).Row).GetCell(new CellReference(aeX[36]).Col).NumericCellValue.ToString();
            pfui.Item_17_17 = ws.GetRow(new CellReference(aeX[37]).Row).GetCell(new CellReference(aeX[37]).Col).NumericCellValue.ToString();
            pfui.Item_17_18 = ws.GetRow(new CellReference(aeX[38]).Row).GetCell(new CellReference(aeX[38]).Col).NumericCellValue.ToString();
            pfui.Item_17_19 = ws.GetRow(new CellReference(aeX[39]).Row).GetCell(new CellReference(aeX[39]).Col).NumericCellValue.ToString();
            pfui.Item_17_20 = ws.GetRow(new CellReference(aeX[40]).Row).GetCell(new CellReference(aeX[40]).Col).NumericCellValue.ToString();
            //CRONOGRAMA
            ///Executado
            if (ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col) != null)
                pfui.Cron_Executado = ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col).NumericCellValue.ToString();
            ///Parcelas 1 a 30
            int arr = 42;
            var valor = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
            int parcelaNumero = 1;

            while (arr <= aeX.Length)
            {
                var properties = pfui.GetType().GetProperties();
                foreach (var property in properties)
                {
                    if (property.Name.Equals("Cron_Parc_" + parcelaNumero.ToString()))
                        if (valor != null)
                            property.SetValue(pfui, ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col).NumericCellValue.ToString());
                        else
                            property.SetValue(pfui, "0");
                }

                parcelaNumero++;
                arr++;
                if (arr < aeX.Length)
                    valor = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
                else
                    break;
            


            


            }


            return pfui;
        }

        internal PropostaEnt GetPCI(string path, string[] aeX)
        {
            FileStream arquivoXLS = new FileStream(path, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(arquivoXLS);
            ISheet ws = wbook.GetSheet("Proposta_Constr_Individual");

            PropostaEnt pci = new PropostaEnt();
            //CABEÇALHO
            pci.Tipo = "PCI";
            pci.Vigencia = ws.Footer.Left;
            pci.Prop_Nome = ws.GetRow(new CellReference(aeX[0]).Row).GetCell(new CellReference(aeX[0]).Col).ToString();
            pci.Prop_CPF = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[1]).Row).GetCell(new CellReference(aeX[1]).Col).ToString());
            pci.Prop_DDD = ws.GetRow(new CellReference(aeX[2]).Row).GetCell(new CellReference(aeX[2]).Col).ToString();
            pci.Prop_Telefone = Util.FormatedFone(ws.GetRow(new CellReference(aeX[3]).Row).GetCell(new CellReference(aeX[3]).Col).ToString());
            pci.Rt_Nome = ws.GetRow(new CellReference(aeX[4]).Row).GetCell(new CellReference(aeX[4]).Col).ToString();
            pci.Rt_CAU_CREA = ws.GetRow(new CellReference(aeX[5]).Row).GetCell(new CellReference(aeX[5]).Col).ToString();
            pci.Rt_UF = ws.GetRow(new CellReference(aeX[6]).Row).GetCell(new CellReference(aeX[6]).Col).ToString();
            pci.Rt_CPF = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[7]).Row).GetCell(new CellReference(aeX[7]).Col).ToString());
            pci.Rt_DDD = ws.GetRow(new CellReference(aeX[8]).Row).GetCell(new CellReference(aeX[8]).Col).ToString();
            pci.Rt_Telefone = Util.FormatedFone(ws.GetRow(new CellReference(aeX[9]).Row).GetCell(new CellReference(aeX[9]).Col).ToString());
            pci.End_Endereço = ws.GetRow(new CellReference(aeX[10]).Row).GetCell(new CellReference(aeX[10]).Col).ToString();
            pci.End_Complemento = ws.GetRow(new CellReference(aeX[11]).Row).GetCell(new CellReference(aeX[11]).Col).ToString();
            pci.End_CEP = Util.FormatedCEP(ws.GetRow(new CellReference(aeX[12]).Row).GetCell(new CellReference(aeX[12]).Col).ToString());
            pci.End_Bairro = ws.GetRow(new CellReference(aeX[13]).Row).GetCell(new CellReference(aeX[13]).Col).ToString();
            pci.End_Municipio = ws.GetRow(new CellReference(aeX[14]).Row).GetCell(new CellReference(aeX[14]).Col).ToString();
            pci.End_UF = ws.GetRow(new CellReference(aeX[15]).Row).GetCell(new CellReference(aeX[15]).Col).ToString();
            pci.Imov_Valor_Terreno = string.Format("{0:0,0.00}", Convert.ToDouble(ws.GetRow(new CellReference(aeX[16]).Row).GetCell(new CellReference(aeX[16]).Col).ToString()));
            pci.Imov_Matricula = ws.GetRow(new CellReference(aeX[17]).Row).GetCell(new CellReference(aeX[17]).Col).ToString();
            pci.Imov_Oficio = ws.GetRow(new CellReference(aeX[18]).Row).GetCell(new CellReference(aeX[18]).Col).ToString();
            pci.Imov_Comarca = null;
            pci.Imov_UF = null;
            //ORÇAMENTO (PERCENTUAIS)
            pci.Item_17_01 = ws.GetRow(new CellReference(aeX[21]).Row).GetCell(new CellReference(aeX[21]).Col).NumericCellValue.ToString();
            pci.Item_17_02 = ws.GetRow(new CellReference(aeX[22]).Row).GetCell(new CellReference(aeX[22]).Col).NumericCellValue.ToString();
            pci.Item_17_03 = ws.GetRow(new CellReference(aeX[23]).Row).GetCell(new CellReference(aeX[23]).Col).NumericCellValue.ToString();
            pci.Item_17_04 = ws.GetRow(new CellReference(aeX[24]).Row).GetCell(new CellReference(aeX[24]).Col).NumericCellValue.ToString();
            pci.Item_17_05 = ws.GetRow(new CellReference(aeX[25]).Row).GetCell(new CellReference(aeX[25]).Col).NumericCellValue.ToString();
            pci.Item_17_06 = ws.GetRow(new CellReference(aeX[26]).Row).GetCell(new CellReference(aeX[26]).Col).NumericCellValue.ToString();
            pci.Item_17_07 = ws.GetRow(new CellReference(aeX[27]).Row).GetCell(new CellReference(aeX[27]).Col).NumericCellValue.ToString();
            pci.Item_17_08 = ws.GetRow(new CellReference(aeX[28]).Row).GetCell(new CellReference(aeX[28]).Col).NumericCellValue.ToString();
            pci.Item_17_09 = ws.GetRow(new CellReference(aeX[29]).Row).GetCell(new CellReference(aeX[29]).Col).NumericCellValue.ToString();
            pci.Item_17_10 = ws.GetRow(new CellReference(aeX[30]).Row).GetCell(new CellReference(aeX[30]).Col).NumericCellValue.ToString();
            pci.Item_17_11 = ws.GetRow(new CellReference(aeX[31]).Row).GetCell(new CellReference(aeX[31]).Col).NumericCellValue.ToString();
            pci.Item_17_12 = ws.GetRow(new CellReference(aeX[32]).Row).GetCell(new CellReference(aeX[32]).Col).NumericCellValue.ToString();
            pci.Item_17_13 = ws.GetRow(new CellReference(aeX[33]).Row).GetCell(new CellReference(aeX[33]).Col).NumericCellValue.ToString();
            pci.Item_17_14 = ws.GetRow(new CellReference(aeX[34]).Row).GetCell(new CellReference(aeX[34]).Col).NumericCellValue.ToString();
            pci.Item_17_15 = ws.GetRow(new CellReference(aeX[35]).Row).GetCell(new CellReference(aeX[35]).Col).NumericCellValue.ToString();
            pci.Item_17_16 = ws.GetRow(new CellReference(aeX[36]).Row).GetCell(new CellReference(aeX[36]).Col).NumericCellValue.ToString();
            pci.Item_17_17 = ws.GetRow(new CellReference(aeX[37]).Row).GetCell(new CellReference(aeX[37]).Col).NumericCellValue.ToString();
            pci.Item_17_18 = ws.GetRow(new CellReference(aeX[38]).Row).GetCell(new CellReference(aeX[38]).Col).NumericCellValue.ToString();
            pci.Item_17_19 = ws.GetRow(new CellReference(aeX[39]).Row).GetCell(new CellReference(aeX[39]).Col).NumericCellValue.ToString();
            pci.Item_17_20 = ws.GetRow(new CellReference(aeX[40]).Row).GetCell(new CellReference(aeX[40]).Col).NumericCellValue.ToString();
            //CRONOGRAMA
            ///Executado
            if (ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col) != null)
                pci.Cron_Executado = ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col).NumericCellValue.ToString();
            ///Parcelas 1 a 30
            int arr = 42;
            var valor = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
            int parcelaNumero = 1;

            while (arr <= aeX.Length)
            {
                var properties = pci.GetType().GetProperties();
                foreach (var property in properties)
                {
                    if (property.Name.Equals("Cron_Parc_" + parcelaNumero.ToString()))
                        if (valor != null)
                            property.SetValue(pci, ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col).NumericCellValue.ToString());
                        else
                            property.SetValue(pci, "0");
                }
            
                parcelaNumero++;
                arr++;
                if (arr < aeX.Length)
                    valor = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
                else
                    break;
            
            }
            return pci;

        }

    }
}
