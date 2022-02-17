using System.Windows.Forms;

namespace aeX30
{
    public class PCI
    {

        //ARRAY FOR DIFERENT SHEET VERSION
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


     //CONFERIR COMARCA E UF


        public static string[] SetArray(string version)
        {
            if (version == "14/07/2021")
                return pci14072021;
            //else if (version == "AE 130 017" || version == "AE 130 018")
            //    return ae130v017_018;
            //else if (version == "AE 130 019" || version == "AE 130 020")
            //    return ae130v019_020;
            //else if (version == "AE 130 021" || version == "AE 130 022" || version == "AE 130 023" || version == "AE 130 024" || version == "AE 130 025")
            //    return ae130v021_025;
            else
            {
                MessageBox.Show("A versão da planilha PCI inserida não foi testada.\r\nRedobre a atenção quanto aos valores importados!", "Versão da planilha não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return pci14072021;
            }
        }






    }
}
