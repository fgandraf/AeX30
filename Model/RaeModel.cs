using System;
using System.IO;
using aeX30.Entities;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace aeX30.Model
{
    public class RaeModel
    {

        internal static string GetFooter(string path)
        {
            FileStream arquivoXLS = new FileStream(path, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(arquivoXLS);

            ISheet sheet = wbook.GetSheet(wbook.GetSheetName(0));
            return sheet.Footer.Left;
        }



        internal int SetRAE(string path_modelo, string path_destino, string[] aeX, RaeEnt rae)
        {
            try
            {
                FileStream arquivoModelo = new FileStream(path_modelo, FileMode.Open, FileAccess.Read);
                HSSFWorkbook wbook = new HSSFWorkbook(arquivoModelo);
                ISheet ws = wbook.GetSheet("RAE");
                
                //CABEÇALHO
                ws.GetRow(new CellReference(aeX[0]).Row).GetCell(new CellReference(aeX[0]).Col).SetCellValue(rae.Ref0);
                ws.GetRow(new CellReference(aeX[1]).Row).GetCell(new CellReference(aeX[1]).Col).SetCellValue(rae.Ref1);
                ws.GetRow(new CellReference(aeX[2]).Row).GetCell(new CellReference(aeX[2]).Col).SetCellValue(Convert.ToInt32(rae.Ref2));
                ws.GetRow(new CellReference(aeX[3]).Row).GetCell(new CellReference(aeX[3]).Col).SetCellValue(rae.Ref3);
                ws.GetRow(new CellReference(aeX[4]).Row).GetCell(new CellReference(aeX[4]).Col).SetCellValue(rae.Ref4);
                ws.GetRow(new CellReference(aeX[5]).Row).GetCell(new CellReference(aeX[5]).Col).SetCellValue(rae.Ref5);
                ws.GetRow(new CellReference(aeX[6]).Row).GetCell(new CellReference(aeX[6]).Col).SetCellValue(rae.Ref6);
                ws.GetRow(new CellReference(aeX[7]).Row).GetCell(new CellReference(aeX[7]).Col).SetCellValue(rae.Prop_Nome);
                ws.GetRow(new CellReference(aeX[8]).Row).GetCell(new CellReference(aeX[8]).Col).SetCellValue(rae.Prop_CPF);
                ws.GetRow(new CellReference(aeX[9]).Row).GetCell(new CellReference(aeX[9]).Col).SetCellValue(rae.Prop_DDD);
                ws.GetRow(new CellReference(aeX[10]).Row).GetCell(new CellReference(aeX[10]).Col).SetCellValue(rae.Prop_Telefone);
                ws.GetRow(new CellReference(aeX[11]).Row).GetCell(new CellReference(aeX[11]).Col).SetCellValue(rae.Rt_Nome);
                ws.GetRow(new CellReference(aeX[12]).Row).GetCell(new CellReference(aeX[12]).Col).SetCellValue(rae.Rt_CAU_CREA);
                ws.GetRow(new CellReference(aeX[13]).Row).GetCell(new CellReference(aeX[13]).Col).SetCellValue(rae.Rt_UF);
                ws.GetRow(new CellReference(aeX[14]).Row).GetCell(new CellReference(aeX[14]).Col).SetCellValue(rae.Rt_CPF);
                ws.GetRow(new CellReference(aeX[15]).Row).GetCell(new CellReference(aeX[15]).Col).SetCellValue(rae.Rt_DDD);
                ws.GetRow(new CellReference(aeX[16]).Row).GetCell(new CellReference(aeX[16]).Col).SetCellValue(rae.Rt_Telefone);
                ws.GetRow(new CellReference(aeX[17]).Row).GetCell(new CellReference(aeX[17]).Col).SetCellValue(rae.End_Endereco);
                ws.GetRow(new CellReference(aeX[18]).Row).GetCell(new CellReference(aeX[18]).Col).SetCellValue(rae.End_Complemento);
                ws.GetRow(new CellReference(aeX[19]).Row).GetCell(new CellReference(aeX[19]).Col).SetCellValue(rae.End_Bairro);
                ws.GetRow(new CellReference(aeX[20]).Row).GetCell(new CellReference(aeX[20]).Col).SetCellValue(rae.End_CEP);
                ws.GetRow(new CellReference(aeX[21]).Row).GetCell(new CellReference(aeX[21]).Col).SetCellValue(rae.End_Municipio);
                ws.GetRow(new CellReference(aeX[22]).Row).GetCell(new CellReference(aeX[22]).Col).SetCellValue(rae.End_UF);
                ws.GetRow(new CellReference(aeX[23]).Row).GetCell(new CellReference(aeX[23]).Col).SetCellValue(rae.Imov_Valor_Terreno);
                ws.GetRow(new CellReference(aeX[24]).Row).GetCell(new CellReference(aeX[24]).Col).SetCellValue(rae.Imov_Matricula);
                ws.GetRow(new CellReference(aeX[25]).Row).GetCell(new CellReference(aeX[25]).Col).SetCellValue(rae.Imov_Oficio);
                ws.GetRow(new CellReference(aeX[26]).Row).GetCell(new CellReference(aeX[26]).Col).SetCellValue(rae.Imov_Comarca);
                ws.GetRow(new CellReference(aeX[27]).Row).GetCell(new CellReference(aeX[27]).Col).SetCellValue(rae.Imov_UF);
                //ORÇAMENTO (PERCENTUAIS)
                ws.GetRow(new CellReference(aeX[28]).Row).GetCell(new CellReference(aeX[28]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_01));
                ws.GetRow(new CellReference(aeX[29]).Row).GetCell(new CellReference(aeX[29]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_02));
                ws.GetRow(new CellReference(aeX[30]).Row).GetCell(new CellReference(aeX[30]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_03));
                ws.GetRow(new CellReference(aeX[31]).Row).GetCell(new CellReference(aeX[31]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_04));
                ws.GetRow(new CellReference(aeX[32]).Row).GetCell(new CellReference(aeX[32]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_05));
                ws.GetRow(new CellReference(aeX[33]).Row).GetCell(new CellReference(aeX[33]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_06));
                ws.GetRow(new CellReference(aeX[34]).Row).GetCell(new CellReference(aeX[34]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_07));
                ws.GetRow(new CellReference(aeX[35]).Row).GetCell(new CellReference(aeX[35]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_08));
                ws.GetRow(new CellReference(aeX[36]).Row).GetCell(new CellReference(aeX[36]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_09));
                ws.GetRow(new CellReference(aeX[37]).Row).GetCell(new CellReference(aeX[37]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_10));
                ws.GetRow(new CellReference(aeX[38]).Row).GetCell(new CellReference(aeX[38]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_11));
                ws.GetRow(new CellReference(aeX[39]).Row).GetCell(new CellReference(aeX[39]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_12));
                ws.GetRow(new CellReference(aeX[40]).Row).GetCell(new CellReference(aeX[40]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_13));
                ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_14));
                ws.GetRow(new CellReference(aeX[42]).Row).GetCell(new CellReference(aeX[42]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_15));
                ws.GetRow(new CellReference(aeX[43]).Row).GetCell(new CellReference(aeX[43]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_16));
                ws.GetRow(new CellReference(aeX[44]).Row).GetCell(new CellReference(aeX[44]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_17));
                ws.GetRow(new CellReference(aeX[45]).Row).GetCell(new CellReference(aeX[45]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_18));
                ws.GetRow(new CellReference(aeX[46]).Row).GetCell(new CellReference(aeX[46]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_19));
                ws.GetRow(new CellReference(aeX[47]).Row).GetCell(new CellReference(aeX[47]).Col).SetCellValue(Convert.ToDouble(rae.Item_17_20));
                if (rae.MensuradoAcumulado != "")
                    ws.GetRow(new CellReference(aeX[48]).Row).GetCell(new CellReference(aeX[48]).Col).SetCellValue(Convert.ToDouble(rae.MensuradoAcumulado));
                //DADOS ADICIONAIS
                if (rae.ContratoInicio != "  /  /")
                    ws.GetRow(new CellReference(aeX[49]).Row).GetCell(new CellReference(aeX[49]).Col).SetCellValue(rae.ContratoInicio);
                if (rae.ContratoTermino != "  /  /")
                    ws.GetRow(new CellReference(aeX[50]).Row).GetCell(new CellReference(aeX[50]).Col).SetCellValue(rae.ContratoTermino);
                
                
                if (aeX[83] != "2021")
                    ws.GetRow(new CellReference(aeX[51]).Row).GetCell(new CellReference(aeX[51]).Col).SetCellValue(rae.EtapaPrevista);

                ////CRONOGRAMA
                ws.GetRow(new CellReference(aeX[52]).Row).GetCell(new CellReference(aeX[52]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Executado));
                ws.GetRow(new CellReference(aeX[53]).Row).GetCell(new CellReference(aeX[53]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_1));
                ws.GetRow(new CellReference(aeX[54]).Row).GetCell(new CellReference(aeX[54]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_2));
                ws.GetRow(new CellReference(aeX[55]).Row).GetCell(new CellReference(aeX[55]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_3));
                ws.GetRow(new CellReference(aeX[56]).Row).GetCell(new CellReference(aeX[56]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_4));
                ws.GetRow(new CellReference(aeX[57]).Row).GetCell(new CellReference(aeX[57]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_5));
                ws.GetRow(new CellReference(aeX[58]).Row).GetCell(new CellReference(aeX[58]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_6));
                ws.GetRow(new CellReference(aeX[59]).Row).GetCell(new CellReference(aeX[59]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_7));
                ws.GetRow(new CellReference(aeX[60]).Row).GetCell(new CellReference(aeX[60]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_8));
                ws.GetRow(new CellReference(aeX[61]).Row).GetCell(new CellReference(aeX[61]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_9));
                ws.GetRow(new CellReference(aeX[62]).Row).GetCell(new CellReference(aeX[62]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_10));
                ws.GetRow(new CellReference(aeX[63]).Row).GetCell(new CellReference(aeX[63]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_11));
                ws.GetRow(new CellReference(aeX[64]).Row).GetCell(new CellReference(aeX[64]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_12));
                ws.GetRow(new CellReference(aeX[65]).Row).GetCell(new CellReference(aeX[65]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_13));
                ws.GetRow(new CellReference(aeX[66]).Row).GetCell(new CellReference(aeX[66]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_14));
                ws.GetRow(new CellReference(aeX[67]).Row).GetCell(new CellReference(aeX[67]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_15));
                ws.GetRow(new CellReference(aeX[68]).Row).GetCell(new CellReference(aeX[68]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_16));
                ws.GetRow(new CellReference(aeX[69]).Row).GetCell(new CellReference(aeX[69]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_17));
                ws.GetRow(new CellReference(aeX[70]).Row).GetCell(new CellReference(aeX[70]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_18));
                ws.GetRow(new CellReference(aeX[71]).Row).GetCell(new CellReference(aeX[71]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_19));
                ws.GetRow(new CellReference(aeX[72]).Row).GetCell(new CellReference(aeX[72]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_20));
                ws.GetRow(new CellReference(aeX[73]).Row).GetCell(new CellReference(aeX[73]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_21));
                ws.GetRow(new CellReference(aeX[74]).Row).GetCell(new CellReference(aeX[74]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_22));
                ws.GetRow(new CellReference(aeX[75]).Row).GetCell(new CellReference(aeX[75]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_23));
                ws.GetRow(new CellReference(aeX[76]).Row).GetCell(new CellReference(aeX[76]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_24));
                ws.GetRow(new CellReference(aeX[77]).Row).GetCell(new CellReference(aeX[77]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_25));
                ws.GetRow(new CellReference(aeX[78]).Row).GetCell(new CellReference(aeX[78]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_26));
                ws.GetRow(new CellReference(aeX[79]).Row).GetCell(new CellReference(aeX[79]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_27));
                ws.GetRow(new CellReference(aeX[80]).Row).GetCell(new CellReference(aeX[80]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_28));
                ws.GetRow(new CellReference(aeX[81]).Row).GetCell(new CellReference(aeX[81]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_29));
                ws.GetRow(new CellReference(aeX[82]).Row).GetCell(new CellReference(aeX[82]).Col).SetCellValue(Convert.ToDouble(rae.Cron_Parc_30));

                using (FileStream arquivoRAE = new FileStream(path_destino, FileMode.Create, FileAccess.Write))
                {
                    wbook.ForceFormulaRecalculation = true;
                    wbook.Write(arquivoRAE);
                }
                return 1;

            }
            catch
            {
                return 0;
            }

        }

    }
}

