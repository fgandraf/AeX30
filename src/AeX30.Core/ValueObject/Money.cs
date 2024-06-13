using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace AeX30.Core.ValueObject
{
    public class Money
    {
        public Money(string currency)
        {
            if (!string.IsNullOrWhiteSpace(currency))
            {
                decimal value;

                if (!decimal.TryParse(currency, NumberStyles.Currency, CultureInfo.CurrentCulture, out value))
                    value = Convert.ToDecimal(Regex.Replace(currency, @"[^\d,]", "").Replace(",", "."), CultureInfo.InvariantCulture);
                else
                    value = Convert.ToDecimal(currency);

                Number = value.ToString("N2");
            }
        }

        public string Number { get; private set; }
    }
}
