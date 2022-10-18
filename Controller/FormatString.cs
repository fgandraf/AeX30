using System;
using System.Text.RegularExpressions;

namespace AeX30.Controller
{
    public class FormatString
    {


        public static string CPF(string cpf)
        {
            if (cpf == "")
                return cpf;
            else
                return OnlyNumber(cpf).ToString(@"000\.000\.000\-00");
        }

        public static string CEP(string cep)
        {
            if (cep == "")
                return cep;
            else
                return OnlyNumber(cep).ToString(@"00000\-000");
        }

        public static string Fone(string fone)
        {
            if (fone == "")
                return fone;
            else
            {
                long phoneNumber = OnlyNumber(fone);
                int lenght = phoneNumber.ToString().Length;

                if (lenght == 8)
                    return phoneNumber.ToString(@"0000\-0000");
                else if (lenght == 9)
                    return phoneNumber.ToString(@"00000\-0000");
                else
                    return fone;
            }
        }

        public static string ValorMonetário(string valor)
        {
            if (valor == "")
                return valor;
            else
            {
                long sonumero = OnlyNumber(valor);

                return $"{sonumero:N2}";

            }
        }







        private static long OnlyNumber(string strIn)
        {
            strIn = strIn.Split(',')[0];
            var onlyNumber = new Regex(@"[^\d]");
            return Convert.ToInt64(onlyNumber.Replace(strIn, ""));
        }






    }
}
