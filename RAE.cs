using System.Windows.Forms;

namespace aeX30
{
    public class RAE
    {



        //ARRAY FOR DIFERENT SHEET VERSION
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
            "BC61","BC62","BC63","BC64","BC65","BC66","BC67","BC68","BC69","BC70","BC71","BC72","BC73","BC74","BC75","BC76","BC77", "BC78", "BC79", "BC80", "BC81", "BC82", "BC83", "BC84", "BC85", "BC86", "BC87", "BC88", "BC89", "BC90", "BC91"
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
            "BD62","BD63","BD64","BD65","BD66","BD67","BD68","BD69","BD70","BD71","BD72","BD73","BD74","BD75","BD76","BD77", "BD78", "BD79", "BD80", "BD81", "BD82", "BD83", "BD84", "BD85", "BD86", "BD87", "BD88", "BD89", "BD90", "BD91", "BD92"
                };


        public static string[] SetArray(string version)
        {
            if (version == "23/07/2020")
                return rae23072020;
            if (version == "09/09/2021")
                return rae09092021;


            else
            {
                MessageBox.Show("A versão da planilha RAE inserida não foi testada.\r\nRedobre a atenção quanto aos valores exportados!", "Versão da planilha não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return rae09092021;
            }
        }




    }
}
