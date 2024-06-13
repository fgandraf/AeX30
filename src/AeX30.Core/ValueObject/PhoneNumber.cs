using System.Text.RegularExpressions;
using System;

namespace AeX30.Core.ValueObject
{
    public class PhoneNumber
    {
        public PhoneNumber(string phone)
        {
            string formatedPhoneNumber = string.Empty;

            if (!string.IsNullOrEmpty(phone))
            {
                phone = new Regex(@"[^\d]").Replace(phone, "");
                long phoneNumber = Convert.ToInt64(phone);

                if (phoneNumber.ToString().Length == 8)
                    formatedPhoneNumber = phoneNumber.ToString(@"0000\-0000");
                else if (phoneNumber.ToString().Length == 9)
                    formatedPhoneNumber = phoneNumber.ToString(@"00000\-0000");
                else
                    formatedPhoneNumber = phone;
            }

            Number =  formatedPhoneNumber;
        }

        public string Number { get; private set; }
    }
}
