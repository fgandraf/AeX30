using System;
using System.Text.RegularExpressions;

namespace aeX30
{
    internal class Util
    {

        internal static string CleanInput(string strIn)
        {
            strIn = strIn.Split(',')[0];
            var onlyNumber = new Regex(@"[^\d]");
            return onlyNumber.Replace(strIn, "");
        }

        internal static string FormatedCPF(string cpf)
        {
            string onlynumber = CleanInput(cpf);
            return Convert.ToUInt64(onlynumber).ToString(@"000\.000\.000\-00");
        }

        internal static string FormatedCEP(string cep)
        {
            string onlynumber = CleanInput(cep);
            return Convert.ToUInt64(onlynumber).ToString(@"00000\-000");
        }

        internal static string FormatedFone(string fone)
        {
            string onlynumber = CleanInput(fone);

            if (onlynumber.Length == 8)
                return Convert.ToUInt64(onlynumber).ToString(@"0000\-0000");
            else if (onlynumber.Length == 9)
                return Convert.ToUInt64(onlynumber).ToString(@"00000\-0000");
            else
                return onlynumber.Length.ToString();
        }


    }
}
