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

        internal static bool isValid(string path)
        {
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);
            
            string sheetName = wbook.GetSheetName(0);
            ISheet sheet = wbook.GetSheet(sheetName);
            string footer = sheet.Footer.Left;


            if ((sheetName == "Proposta" || sheetName == "Proposta_Constr_Individual") && (footer != "" || footer != null))
                return true;
            else
                return false;
        }


        internal PropostaEnt GetProposal(string path)
        {
            //Open File
            FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);
            ISheet ws = wbook.GetSheet(wbook.GetSheetName(0));
            string footer = ws.Footer.Left;


            //Set the right for the case array
            int version = Convert.ToInt32(Util.CleanInput(footer));
            string[] aeX = SetArray(version);


            //Create the object
            PropostaEnt proposal = new PropostaEnt();


            //Star populate the object
            proposal.Vigencia = footer;

            if (wbook.GetSheetName(0) == "Proposta")
            {
                proposal.Tipo = "PFUI";
                proposal.Imov_Comarca = ws.GetRow(new CellReference(aeX[19]).Row).GetCell(new CellReference(aeX[19]).Col).ToString();
                proposal.Imov_UF = ws.GetRow(new CellReference(aeX[20]).Row).GetCell(new CellReference(aeX[20]).Col).ToString();
            }
            else
            {
                proposal.Tipo = "PCI";
                proposal.Imov_Comarca = null;
                proposal.Imov_UF = null;
            }

            ///HEADER         
            proposal.Prop_Nome = ws.GetRow(new CellReference(aeX[0]).Row).GetCell(new CellReference(aeX[0]).Col).ToString();
            proposal.Prop_CPF = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[1]).Row).GetCell(new CellReference(aeX[1]).Col).ToString());
            proposal.Prop_DDD = ws.GetRow(new CellReference(aeX[2]).Row).GetCell(new CellReference(aeX[2]).Col).ToString();
            proposal.Prop_Telefone = Util.FormatedFone(ws.GetRow(new CellReference(aeX[3]).Row).GetCell(new CellReference(aeX[3]).Col).ToString());
            proposal.Rt_Nome = ws.GetRow(new CellReference(aeX[4]).Row).GetCell(new CellReference(aeX[4]).Col).ToString();
            proposal.Rt_CAU_CREA = ws.GetRow(new CellReference(aeX[5]).Row).GetCell(new CellReference(aeX[5]).Col).ToString();
            proposal.Rt_UF = ws.GetRow(new CellReference(aeX[6]).Row).GetCell(new CellReference(aeX[6]).Col).ToString();
            proposal.Rt_CPF = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[7]).Row).GetCell(new CellReference(aeX[7]).Col).ToString());
            proposal.Rt_DDD = ws.GetRow(new CellReference(aeX[8]).Row).GetCell(new CellReference(aeX[8]).Col).ToString();
            proposal.Rt_Telefone = Util.FormatedFone(ws.GetRow(new CellReference(aeX[9]).Row).GetCell(new CellReference(aeX[9]).Col).ToString());
            proposal.End_Endereço = ws.GetRow(new CellReference(aeX[10]).Row).GetCell(new CellReference(aeX[10]).Col).ToString();
            proposal.End_Complemento = ws.GetRow(new CellReference(aeX[11]).Row).GetCell(new CellReference(aeX[11]).Col).ToString();
            proposal.End_CEP = Util.FormatedCEP(ws.GetRow(new CellReference(aeX[12]).Row).GetCell(new CellReference(aeX[12]).Col).ToString());
            proposal.End_Bairro = ws.GetRow(new CellReference(aeX[13]).Row).GetCell(new CellReference(aeX[13]).Col).ToString();
            proposal.End_Municipio = ws.GetRow(new CellReference(aeX[14]).Row).GetCell(new CellReference(aeX[14]).Col).ToString();
            proposal.End_UF = ws.GetRow(new CellReference(aeX[15]).Row).GetCell(new CellReference(aeX[15]).Col).ToString();
            try
            {
                proposal.Imov_Valor_Terreno = string.Format("{0:0,0.00}", Convert.ToDouble(ws.GetRow(new CellReference(aeX[16]).Row).GetCell(new CellReference(aeX[16]).Col).ToString()));
            }
            catch
            {
                proposal.Imov_Valor_Terreno = null;
            }
            proposal.Imov_Matricula = ws.GetRow(new CellReference(aeX[17]).Row).GetCell(new CellReference(aeX[17]).Col).ToString();
            proposal.Imov_Oficio = ws.GetRow(new CellReference(aeX[18]).Row).GetCell(new CellReference(aeX[18]).Col).ToString();
            ///ORÇAMENTO (PERCENTUAIS)
            proposal.Item_17_01 = ws.GetRow(new CellReference(aeX[21]).Row).GetCell(new CellReference(aeX[21]).Col).NumericCellValue.ToString();
            proposal.Item_17_02 = ws.GetRow(new CellReference(aeX[22]).Row).GetCell(new CellReference(aeX[22]).Col).NumericCellValue.ToString();
            proposal.Item_17_03 = ws.GetRow(new CellReference(aeX[23]).Row).GetCell(new CellReference(aeX[23]).Col).NumericCellValue.ToString();
            proposal.Item_17_04 = ws.GetRow(new CellReference(aeX[24]).Row).GetCell(new CellReference(aeX[24]).Col).NumericCellValue.ToString();
            proposal.Item_17_05 = ws.GetRow(new CellReference(aeX[25]).Row).GetCell(new CellReference(aeX[25]).Col).NumericCellValue.ToString();
            proposal.Item_17_06 = ws.GetRow(new CellReference(aeX[26]).Row).GetCell(new CellReference(aeX[26]).Col).NumericCellValue.ToString();
            proposal.Item_17_07 = ws.GetRow(new CellReference(aeX[27]).Row).GetCell(new CellReference(aeX[27]).Col).NumericCellValue.ToString();
            proposal.Item_17_08 = ws.GetRow(new CellReference(aeX[28]).Row).GetCell(new CellReference(aeX[28]).Col).NumericCellValue.ToString();
            proposal.Item_17_09 = ws.GetRow(new CellReference(aeX[29]).Row).GetCell(new CellReference(aeX[29]).Col).NumericCellValue.ToString();
            proposal.Item_17_10 = ws.GetRow(new CellReference(aeX[30]).Row).GetCell(new CellReference(aeX[30]).Col).NumericCellValue.ToString();
            proposal.Item_17_11 = ws.GetRow(new CellReference(aeX[31]).Row).GetCell(new CellReference(aeX[31]).Col).NumericCellValue.ToString();
            proposal.Item_17_12 = ws.GetRow(new CellReference(aeX[32]).Row).GetCell(new CellReference(aeX[32]).Col).NumericCellValue.ToString();
            proposal.Item_17_13 = ws.GetRow(new CellReference(aeX[33]).Row).GetCell(new CellReference(aeX[33]).Col).NumericCellValue.ToString();
            proposal.Item_17_14 = ws.GetRow(new CellReference(aeX[34]).Row).GetCell(new CellReference(aeX[34]).Col).NumericCellValue.ToString();
            proposal.Item_17_15 = ws.GetRow(new CellReference(aeX[35]).Row).GetCell(new CellReference(aeX[35]).Col).NumericCellValue.ToString();
            proposal.Item_17_16 = ws.GetRow(new CellReference(aeX[36]).Row).GetCell(new CellReference(aeX[36]).Col).NumericCellValue.ToString();
            proposal.Item_17_17 = ws.GetRow(new CellReference(aeX[37]).Row).GetCell(new CellReference(aeX[37]).Col).NumericCellValue.ToString();
            proposal.Item_17_18 = ws.GetRow(new CellReference(aeX[38]).Row).GetCell(new CellReference(aeX[38]).Col).NumericCellValue.ToString();
            proposal.Item_17_19 = ws.GetRow(new CellReference(aeX[39]).Row).GetCell(new CellReference(aeX[39]).Col).NumericCellValue.ToString();
            proposal.Item_17_20 = ws.GetRow(new CellReference(aeX[40]).Row).GetCell(new CellReference(aeX[40]).Col).NumericCellValue.ToString();
            ///CRONOGRAMA
            ///Executado
            if (ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col) != null)
                proposal.Cron_Executado = ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col).NumericCellValue.ToString();
            ///Parcelas 1 a 30
            int arr = 42;
            var valor = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
            int parcelaNumero = 1;

            while (arr <= aeX.Length)
            {
                var properties = proposal.GetType().GetProperties();
                foreach (var property in properties)
                {
                    if (property.Name.Equals("Cron_Parc_" + parcelaNumero.ToString()))
                        if (valor != null)
                            property.SetValue(proposal, ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col).NumericCellValue.ToString());
                        else
                            property.SetValue(proposal, "0");
                }

                parcelaNumero++;
                arr++;
                if (arr < aeX.Length)
                    valor = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
                else
                    break;
            }
            
            

            //Return the populated object
            return proposal;
        }










        private static string[] SetArray(int foot)
        {
            switch (foot)
            {
                case 927112017:
                    return pfui_2017;

                case 1102018:
                    return pfui_2018;

                case 26022018:
                    return pfui_2018b;

                case 12042019:
                case 23052019:
                case 24052019:
                    return pfui_2019;

                case 13072020:
                case 23072020:
                    return pfui_2020a;

                case 24092020:
                case 26022021:
                case 5052021:
                    return pfui_2020b;

                case 4122020:
                    return pfui_2020c;

                case 1062021:
                case 5072021:
                case 14072021:
                case 6082021:
                    return pci_2021a;

                case 21102021:
                case 4112021:
                    return pci_2021b;

                default:
                    return null;
            }
        }
        private static string[] pfui_2017 = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR112", "AR114", "AR126", "AR133", "AR142", "AR152", "AR161", "AR168", "AR175", "AR187", "AR194", "AR205", "AR215", "AR226", "AR232", "AR243", "AR252", "AR260", "AR269", "AR271",
            "AJ309",
            "AL308", "AP308", "AT308", "AX308", "BB308", "BF308", "BJ308", "BN308",
            "AL348", "AP348", "AT348", "AX348", "BB348", "BF348", "BJ348", "BN348",
            "AL388", "AP388", "AT388", "AX388", "BB388", "BF388", "BJ388", "BN388",
            "AL428", "AP428", "AT428", "AX428", "BB428", "BF428"
        };
        private static string[] pfui_2018 = new string[]
                {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR116", "AR118", "AR130", "AR137", "AR146", "AR156", "AR165", "AR172", "AR179", "AR191", "AR198", "AR209", "AR219", "AR230", "AR236", "AR247", "AR256", "AR264", "AR273", "AR275",
            "AJ313",
            "AL312", "AP312", "AT312", "AX312", "BB312", "BF312", "BJ312", "BN312",
            "AL352", "AP352", "AT352", "AX352", "BB352", "BF352", "BJ352", "BN352",
            "AL392", "AP392", "AT392", "AX392", "BB392", "BF392", "BJ392", "BN392",
            "AL432", "AP432", "AT432", "AX432", "BB432", "BF432"
                };
        private static string[] pfui_2018b = new string[]
                {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR112", "AR114", "AR126", "AR133", "AR142", "AR152", "AR161", "AR168", "AR175", "AR187", "AR194", "AR205", "AR215", "AR226", "AR232", "AR243", "AR252", "AR260", "AR269", "AR271",
            "AJ309",
            "AL308", "AP308", "AT308", "AX308", "BB308", "BF308", "BJ308", "BN308",
            "AL348", "AP348", "AT348", "AX348", "BB348", "BF348", "BJ348", "BN348",
            "AL388", "AP388", "AT388", "AX388", "BB388", "BF388", "BJ388", "BN388",
            "AL428", "AP428", "AT428", "AX428", "BB428", "BF428"
                };
        private static string[] pfui_2019 = new string[]
                {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR116", "AR118", "AR130", "AR137", "AR146", "AR156", "AR165", "AR172", "AR179", "AR190", "AR197", "AR207", "AR217", "AR228", "AR234", "AR245", "AR254", "AR262", "AR271", "AR273",
            "AJ311",
            "AL310", "AP310", "AT310", "AX310", "BB310", "BF310", "BJ310", "BN310",
            "AL350", "AP350", "AT350", "AX350", "BB350", "BF350", "BJ350", "BN350",
            "AL390", "AP390", "AT390", "AX390", "BB390", "BF390", "BJ390", "BN390",
            "AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
                };
        private static string[] pfui_2020a = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR116", "AR118", "AR130", "AR137", "AR146", "AR156", "AR165", "AR172", "AR179", "AR190", "AR197", "AR207", "AR217", "AR228", "AR234", "AR245", "AR254", "AR262", "AR271", "AR273",
            "AJ312",
            "AL311", "AP311", "AT311", "AX311", "BB311", "BF311", "BJ311", "BN311",
            "AL351", "AP351", "AT351", "AX351", "BB351", "BF351", "BJ351", "BN351",
            "AL391", "AP391", "AT391", "AX391", "BB391", "BF391", "BJ391", "BN391",
            "AL431", "AP431", "AT431", "AX431", "BB431", "BF431"

        };
        private static string[] pfui_2020b = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR114", "AR116", "AR128", "AR135", "AR144", "AR154", "AR163", "AR170", "AR177", "AR188", "AR195", "AR205", "AR215", "AR226", "AR232", "AR243", "AR252", "AR260", "AR269", "AR271",
            "AJ310",
            "AL309", "AP309", "AT309", "AX309", "BB309", "BF309", "BJ309", "BN309",
            "AL349", "AP349", "AT349", "AX349", "BB349", "BF349", "BJ349", "BN349",
            "AL389", "AP389", "AT389", "AX389", "BB389", "BF389", "BJ389", "BN389",
            "AL429", "AP429", "AT429", "AX429", "BB429", "BF429"
        };
        private static string[] pfui_2020c = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR115", "AR117", "AR129", "AR136", "AR145", "AR155", "AR164", "AR171", "AR178", "AR189", "AR196", "AR206", "AR216", "AR227", "AR233", "AR244", "AR253", "AR261", "AR270", "AR272",
            "AJ311",
            "AL310", "AP310", "AT310", "AX310", "BB310", "BF310", "BJ310", "BN310",
            "AL350", "AP350", "AT350", "AX350", "BB350", "BF350", "BJ350", "BN350",
            "AL390", "AP390", "AT390", "AX390", "BB390", "BF390", "BJ390", "BN390",
            "AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
        };
        private static string[] pci_2021a = new string[]
{
            
            /*prop. NOME:----------*/    "G43",
            /*prop. CPF:-----------*/    "AK43",
            /*prop. DDD:-----------*/    "AQ43",
            /*prop. TELEFONE:------*/    "AS43",
            
            /*rt. NOME:------------*/    "G49",
            /*rt. CAU_CREA:--------*/    "AB49",
            /*rt. UF:--------------*/    "AI49",
            /*rt. CPF:-------------*/    "AK49",
            /*rt. DDD:-------------*/    "AQ49",
            /*rt. TELEFONE:--------*/    "AS49",

            /*end. ENDEREÇO:-------*/    "G53",
            /*end. COMPLEMENTO:----*/    "AJ53",
            /*end. CEP:------------*/    "V55",
            /*end. BAIRRO:---------*/    "G55",
            /*end. MUNICIPIO:------*/    "AA55",
            /*end. UF:-------------*/    "AV55",
            
            /*imov. VALOR TERRENO:-*/    "V70",
            /*imov. MATRICULA:-----*/    "G57",
            /*imov. OFICIO:--------*/    "M57",
            /*imov. COMARCA:-------*/    "-",
            /*imov. UF:------------*/    "-",
           
            /*17.01(%):------------*/    "X94",
            /*17.02(%):------------*/    "X95",
            /*17.03(%):------------*/    "X96",
            /*17.04(%):------------*/    "X97",
            /*17.05(%):------------*/    "X98",
            /*17.06(%):------------*/    "X99",
            /*17.07(%):------------*/    "X100",
            /*17.08(%):------------*/    "X101",
            /*17.09(%):------------*/    "X102",
            /*17.10(%):------------*/    "X103",
            /*17.11(%):------------*/    "X104",
            /*17.12(%):----PINTURA-*/    "X106",
            /*17.13(%):----PISO----*/    "X105",
            /*17.14(%):------------*/    "X107",
            /*17.15(%):------------*/    "X108",
            /*17.16(%):------------*/    "X109",
            /*17.17(%):------------*/    "X110",
            /*17.18(%):------------*/    "X111",
            /*17.19(%):------------*/    "X112",
            /*17.20(%):------------*/    "X113",

            /*cron. EXECUTADO:-----*/    "AM139",
            /*cron. PARC 1:--------*/    "AM140",
            /*cron. PARC 2:--------*/    "AM141",
            /*cron. PARC 3:--------*/    "AM142",
            /*cron. PARC 4:--------*/    "AM143",
            /*cron. PARC 5:--------*/    "AM144",
            /*cron. PARC 6:--------*/    "AM145",
            /*cron. PARC 7:--------*/    "AM146",
            /*cron. PARC 8:--------*/    "AM147",
            /*cron. PARC 9:--------*/    "AM148",
            /*cron. PARC 10:-------*/    "AM149",
            /*cron. PARC 11:-------*/    "AM150",
            /*cron. PARC 12:-------*/    "AM151",
            /*cron. PARC 13:-------*/    "AM152",
            /*cron. PARC 14:-------*/    "AM153",
            /*cron. PARC 15:-------*/    "AM154",
            /*cron. PARC 16:-------*/    "AM155",
            /*cron. PARC 17:-------*/    "AM156",
            /*cron. PARC 18:-------*/    "AM157",
            /*cron. PARC 19:-------*/    "AM158",
            /*cron. PARC 20:-------*/    "AM159",
            /*cron. PARC 21:-------*/    "AM160",
            /*cron. PARC 22:-------*/    "AM161",
            /*cron. PARC 23:-------*/    "AM162",
            /*cron. PARC 24:-------*/    "AM163",
};
        private static string[] pci_2021b = new string[]
        {
            
            /*prop. NOME:----------*/    "G43",
            /*prop. CPF:-----------*/    "AK43",
            /*prop. DDD:-----------*/    "AQ43",
            /*prop. TELEFONE:------*/    "AS43",
            
            /*rt. NOME:------------*/    "G49",
            /*rt. CAU_CREA:--------*/    "AB49",
            /*rt. UF:--------------*/    "AI49",
            /*rt. CPF:-------------*/    "AK49",
            /*rt. DDD:-------------*/    "AQ49",
            /*rt. TELEFONE:--------*/    "AS49",

            /*end. ENDEREÇO:-------*/    "G53",
            /*end. COMPLEMENTO:----*/    "AJ53",
            /*end. CEP:------------*/    "V55",
            /*end. BAIRRO:---------*/    "G55",
            /*end. MUNICIPIO:------*/    "AA55",
            /*end. UF:-------------*/    "AV55",
            
            /*imov. VALOR TERRENO:-*/    "V70",
            /*imov. MATRICULA:-----*/    "G57",
            /*imov. OFICIO:--------*/    "M57",
            /*imov. COMARCA:-------*/    "-",
            /*imov. UF:------------*/    "-",
           
            /*17.01(%):------------*/    "X95",
            /*17.02(%):------------*/    "X96",
            /*17.03(%):------------*/    "X97",
            /*17.04(%):------------*/    "X98",
            /*17.05(%):------------*/    "X99",
            /*17.06(%):------------*/    "X100",
            /*17.07(%):------------*/    "X101",
            /*17.08(%):------------*/    "X102",
            /*17.09(%):------------*/    "X103",
            /*17.10(%):------------*/    "X104",
            /*17.11(%):------------*/    "X105",
            /*17.12(%):----PINTURA-*/    "X106",
            /*17.13(%):----PISO----*/    "X107",
            /*17.14(%):------------*/    "X108",
            /*17.15(%):------------*/    "X109",
            /*17.16(%):------------*/    "X110",
            /*17.17(%):------------*/    "X111",
            /*17.18(%):------------*/    "X112",
            /*17.19(%):------------*/    "X113",
            /*17.20(%):------------*/    "X114",

            /*cron. EXECUTADO:-----*/    "AM140",
            /*cron. PARC 1:--------*/    "AM141",
            /*cron. PARC 2:--------*/    "AM142",
            /*cron. PARC 3:--------*/    "AM143",
            /*cron. PARC 4:--------*/    "AM144",
            /*cron. PARC 5:--------*/    "AM145",
            /*cron. PARC 6:--------*/    "AM146",
            /*cron. PARC 7:--------*/    "AM147",
            /*cron. PARC 8:--------*/    "AM148",
            /*cron. PARC 9:--------*/    "AM149",
            /*cron. PARC 10:-------*/    "AM150",
            /*cron. PARC 11:-------*/    "AM151",
            /*cron. PARC 12:-------*/    "AM152",
            /*cron. PARC 13:-------*/    "AM153",
            /*cron. PARC 14:-------*/    "AM154",
            /*cron. PARC 15:-------*/    "AM155",
            /*cron. PARC 16:-------*/    "AM156",
            /*cron. PARC 17:-------*/    "AM157",
            /*cron. PARC 18:-------*/    "AM158",
            /*cron. PARC 19:-------*/    "AM159",
            /*cron. PARC 20:-------*/    "AM160",
            /*cron. PARC 21:-------*/    "AM161",
            /*cron. PARC 22:-------*/    "AM162",
            /*cron. PARC 23:-------*/    "AM163",
            /*cron. PARC 24:-------*/    "AM164",
        };


    }
}
