
using System.Text.RegularExpressions;
using System;

namespace AeX30.Domain.ValueObject
{
    public class Document
    {
        public Document(string document) //CPF Only
        {
            string formatedDocument = string.Empty;

            if (!string.IsNullOrEmpty(document))
            {
                document = document.Split(',')[0];
                document = new Regex(@"[^\d]").Replace(document, "");
                long number = Convert.ToInt64(document);

                formatedDocument = number.ToString(@"000\.000\.000\-00");
            }
            Number = formatedDocument;
        }

        public string Number { get; private set; }
    }
}
