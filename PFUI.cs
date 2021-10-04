using System.Windows.Forms;

namespace aeX30
{
    public class PFUI
    {



        //ARRAY FOR DIFERENT SHEET VERSION
        private static string[] ae130v016 = new string[]
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
        private static string[] ae130v017_018 = new string[]
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
        private static string[] ae130v019_020 = new string[]
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
        private static string[] ae130v021_025 = new string[]
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




        public static string[] SetArray(string version)
        {
            if (version == "AE 130 016")
                return ae130v016;
            else if (version == "AE 130 017" || version == "AE 130 018")
                return ae130v017_018;
            else if (version == "AE 130 019" || version == "AE 130 020")
                return ae130v019_020;
            else if (version == "AE 130 021" || version == "AE 130 022" || version == "AE 130 023" || version == "AE 130 024" || version == "AE 130 025")
                return ae130v021_025;
            else
            {
                MessageBox.Show("A versão da planilha PFUI inserida não foi testada.\r\nRedobre a atenção quanto aos valores importados!", "Versão da planilha não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return ae130v021_025;
            }
        }






    }
}
