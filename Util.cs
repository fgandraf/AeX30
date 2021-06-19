using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace Raecef
{
    public class Util
    {
        public static string CleanInput (string strIn)
        {
            var onlyNumber = new Regex(@"[^\d]");
            return onlyNumber.Replace(strIn, "");
        }

        public static string FormatedCPF (string cpf)
        {
            string onlynumber = CleanInput(cpf);
            return Convert.ToUInt64(onlynumber).ToString(@"000\.000\.000\-00");
        }

        public static string FormatedCEP(string cep)
        {
            string onlynumber = CleanInput(cep);
            return Convert.ToUInt64(onlynumber).ToString(@"00000\-000");
        }

        public static string FormatedFone(string fone)
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
