using System;
using System.Text.RegularExpressions;

namespace AeX30
{
    public class FormatString
    {


        public static string CPF(string cpf)
        {
            return OnlyNumber(cpf).ToString(@"000\.000\.000\-00");
        }

        public static string CEP(string cep)
        {
            return OnlyNumber(cep).ToString(@"00000\-000");
        }

        public static string Fone(string fone)
        {
            int phoneNumber = OnlyNumber(fone);
            int lenght = phoneNumber.ToString().Length;

            if (lenght == 8)
                return phoneNumber.ToString(@"0000\-0000");
            else if (lenght == 9)
                return phoneNumber.ToString(@"00000\-0000");
            else
                return fone;
        }









        private static int OnlyNumber(string strIn)
        {
            strIn = strIn.Split(',')[0];
            var onlyNumber = new Regex(@"[^\d]");
            return Convert.ToInt32(onlyNumber.Replace(strIn, ""));
        }






    }
}
