using System.Text.RegularExpressions;
using System;

namespace AeX30.Domain.ValueObject
{
    public class ZipCode
    {
        public ZipCode(string zip)
        {
            string formatedZipCode = string.Empty;

            if (!string.IsNullOrEmpty(zip))
            {
                zip = new Regex(@"[^\d]").Replace(zip, "");
                long zipNumber = Convert.ToInt64(zip);

                formatedZipCode = zipNumber.ToString(@"00000\-000");
            }
            Number = formatedZipCode;
        }

        public string Number { get; private set; }
    }
}
