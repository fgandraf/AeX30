using aeX30.Entities;
using aeX30.Model;
using System;

namespace aeX30.Controller
{
    internal class PropostaController
    {


        private static string[] pci14072021 = new string[]
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
        private static string[] pci21102021 = new string[]
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
        private static string[] pfuiv016 = new string[]
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
        private static string[] pfuiv17_18 = new string[]
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
        private static string[] pfuiv19_20 = new string[]
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
        private static string[] pfuiv21_25 = new string[]
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




        private static string[] SetArray(int foot, string type)
        {
            if (type == "PFUI")
            {
                if (foot == 1102018)
                    return pfuiv016;
                else if (foot == 12042019 || foot == 23052019)
                    return pfuiv17_18;
                else if (foot == 13072020 || foot == 13072020)
                    return pfuiv19_20;
                else if (foot == 24092020 || /*foot=planilha22 ||*/ foot == 4122020 || foot == 26022021 || foot == 5052021)
                    return pfuiv21_25;
                else
                    //MessageBox.Show("A versão da planilha PFUI inserida não foi testada.\r\nRedobre a atenção quanto aos valores importados!", "Versão da planilha não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return pfuiv21_25;
            }
            else
            {
                if (foot == 14072021)
                    return pci14072021;
                else if (foot == 21102021)
                    return pci21102021;
                else
                {
                    //MessageBox.Show("A versão da planilha PFUI inserida não foi testada.\r\nRedobre a atenção quanto aos valores importados!", "Versão da planilha não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return pci21102021;
                }
            }



        }




        internal PropostaEnt GetProposta(string path)
        {
            string type = PropostaModel.GetType(path);

            if (type == null)
                return null;
            else
            {
                
                string footer = PropostaModel.GetFooter(path);

                if (footer == "" || footer == null)
                    return null;
                else
                {
                    int version = Convert.ToInt32(Util.CleanInput(footer));
                    string[] aeX = SetArray(version, type);

                    if (aeX != null && type == "PFUI")
                        return new PropostaModel().GetPFUI(path, aeX);
                    else if (aeX != null && type == "PCI")
                        return new PropostaModel().GetPCI(path, aeX);
                    else
                        return null;
                }

            }

        }




    }


}
