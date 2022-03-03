using aeX30.Entities;
using aeX30.Model;
using System;

namespace aeX30.Controller
{
    internal class RaeController
    {

        private static string[] rae23072020 = new string[]
                {
            "AD35", "AF35", "AH35", "AL35", "AN35", "AO35", "AP35",
            "G43", "AJ43", "AP43", "AR43",
            "G46", "Z46", "AH46", "AJ46", "AP46", "AR46",
            "G49", "AJ49",
            "G51", "V51", "AA51", "AS51",
            "G53", "Q53", "AA53", "AJ53", "AS53",
            "S68","S69","S70","S71","S72","S73","S74","S75","S76","S77","S78","S79","S80","S81","S82","S83","S84","S85","S86","S87",
            "Y89",
            "AH63", "AS63",
            "AS74",
            "BC61","BC62","BC63","BC64","BC65","BC66","BC67","BC68","BC69","BC70","BC71","BC72","BC73","BC74","BC75","BC76","BC77", "BC78", "BC79", "BC80", "BC81", "BC82", "BC83", "BC84", "BC85", "BC86", "BC87", "BC88", "BC89", "BC90", "BC91",
               /*versão.......:-------*/    "2020"
                };
        private static string[] rae09092021 = new string[]
                {
            "AB35", "AD35", "AF35", "AJ35", "AL35", "AM35", "AN35",
            "G43", "AJ43", "AP43", "AR43",
            "G46", "Z46", "AH46", "AJ46", "AP46", "AR46",
            "G49", "AJ49",
            "G51", "V51", "AA51", "AS51",
            "G53", "Q53", "AA53", "AJ53", "AS53",
            "S68","S69","S70","S71","S72","S73","S74","S75","S76","S77","S78","S79","S80","S81","S82","S83","S84","S85","S86","S87",
            "Y89",
            "AH63", "AS63",
            "AS74",
            "BD62","BD63","BD64","BD65","BD66","BD67","BD68","BD69","BD70","BD71","BD72","BD73","BD74","BD75","BD76","BD77", "BD78", "BD79", "BD80", "BD81", "BD82", "BD83", "BD84", "BD85", "BD86", "BD87", "BD88", "BD89", "BD90", "BD91", "BD92",
                           /*versão.......:-------*/    "2021"
                };
        private static string[] rae22102021 = new string[]
                {
            /*auto. de serv.:------*/    "AB35", "AD35", "AF35", "AJ35", "AL35", "AM35", "AN35",
            
            /*prop. NOME:----------*/    "G43",
            /*prop. CPF:-----------*/    "AJ43",
            /*prop. DDD:-----------*/    "AP43",
            /*prop. TELEFONE:------*/    "AR43",
            
            /*rt. NOME:------------*/    "G46",
            /*rt. CAU_CREA:--------*/    "Z46",
            /*rt. UF:--------------*/    "AH46",
            /*rt. CPF:-------------*/    "AJ46",
            /*rt. DDD:-------------*/    "AP46",
            /*rt. TELEFONE:--------*/    "AR46",

            /*end. ENDEREÇO:-------*/    "G49",
            /*end. COMPLEMENTO:----*/    "AJ49",
            /*end. BAIRRO:---------*/    "G51",
            /*end. CEP:------------*/    "V51",
            /*end. MUNICIPIO:------*/    "AA51",
            /*end. UF:-------------*/    "AS51",
            
            /*imov. VALOR TERRENO:-*/    "G53",
            /*imov. MATRICULA:-----*/    "Q53",
            /*imov. OFICIO:--------*/    "AA53",
            /*imov. COMARCA:-------*/    "AJ53",
            /*imov. UF:------------*/    "AS53",
           
            /*17.01(%):------------*/    "S68",
            /*17.02(%):------------*/    "S69",
            /*17.03(%):------------*/    "S70",
            /*17.04(%):------------*/    "S71",
            /*17.05(%):------------*/    "S72",
            /*17.06(%):------------*/    "S73",
            /*17.07(%):------------*/    "S74",
            /*17.08(%):------------*/    "S75",
            /*17.09(%):------------*/    "S76",
            /*17.10(%):------------*/    "S77",
            /*17.11(%):------------*/    "S78",
            /*17.12(%):------------*/    "S79",
            /*17.13(%):------------*/    "S80",
            /*17.14(%):------------*/    "S81",
            /*17.15(%):------------*/    "S82",
            /*17.16(%):------------*/    "S83",
            /*17.17(%):------------*/    "S84",
            /*17.18(%):------------*/    "S85",
            /*17.19(%):------------*/    "S86",
            /*17.20(%):------------*/    "S87",
            /*ACUMULADO(%):--------*/    "Y89",

            /*contrato INICIO:-----*/    "AH63",
            /*contrato TERMINO:----*/    "AS63",

            /*ETAPA PREVISTA:------*/    "-",

            /*cron. EXECUTADO:-----*/    "AK71",
            /*cron. PARC 1:--------*/    "AK72",
            /*cron. PARC 2:--------*/    "AK73",
            /*cron. PARC 3:--------*/    "AK74",
            /*cron. PARC 4:--------*/    "AK75",
            /*cron. PARC 5:--------*/    "AK76",
            /*cron. PARC 6:--------*/    "AK77",
            /*cron. PARC 7:--------*/    "AK78",
            /*cron. PARC 8:--------*/    "AK79",
            /*cron. PARC 9:--------*/    "AK80",
            /*cron. PARC 10:-------*/    "AK81",
            /*cron. PARC 11:-------*/    "AK82",
            /*cron. PARC 12:-------*/    "AK83",
            /*cron. PARC 13:-------*/    "AK84",
            /*cron. PARC 14:-------*/    "AK85",
            /*cron. PARC 15:-------*/    "AK86",
            /*cron. PARC 16:-------*/    "AK87",
            /*cron. PARC 17:-------*/    "AK88",
            /*cron. PARC 18:-------*/    "AK89",
            /*cron. PARC 19:-------*/    "AK90",
            /*cron. PARC 20:-------*/    "AK91",
            /*cron. PARC 21:-------*/    "AK92",
            /*cron. PARC 22:-------*/    "AK93",
            /*cron. PARC 23:-------*/    "AK94",
            /*cron. PARC 24:-------*/    "AK95",
            /*cron. PARC 25:-------*/    "AK96",
            /*cron. PARC 26:-------*/    "AK97",
            /*cron. PARC 27:-------*/    "AK98",
            /*cron. PARC 28:-------*/    "AK99",
            /*cron. PARC 29:-------*/    "AK100",
            /*cron. PARC 30:-------*/    "AK101",
            /*versão.......:-------*/    "2021"
                };



        private static string[] SetArray(int version)
        {
            if (version == 23072020)
                return rae23072020;
            else if (version == 9092021)
                return rae09092021;
            else if (version == 22102021)
                return rae22102021;
            else
                //MessageBox.Show("A versão da planilha RAE inserida não foi testada.\r\nRedobre a atenção quanto aos valores importados!", "Versão da planilha não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
        }


        internal static int SetRAE(string path_modelo, string path_destino, RaeEnt rae)
        {

            string footer = RaeModel.GetFooter(path_modelo);

            if (footer == "" || footer == null)
                return 0;
            else
            {
                int version = Convert.ToInt32(Util.CleanInput(footer));
                string[] aeX = SetArray(version);

                if (aeX != null)
                    return new RaeModel().SetRAE(path_modelo, path_destino, aeX, rae);
                else
                    return 0;
            }

        }



    }
}







