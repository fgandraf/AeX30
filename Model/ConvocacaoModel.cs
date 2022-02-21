using System;
using System.IO;
using System.Linq;

namespace aeX30.Model
{
    internal class ConvocacaoModel
    {
        internal string[] GetReferencia(string path)
        {
            var convocacao = File.ReadAllLines(path)
                               .Where(l => l.StartsWith("Refer"))
                               .Select(l => l.Substring(l.LastIndexOf("-") + 2))
                               .ToList();


            string referencia = Util.CleanInput(convocacao[0]);
            if (referencia.Substring(0, 1) == "0")
                referencia = referencia.Substring(1);


            string[] refe = new string[]
            {
            referencia.Substring(0, 4),
            referencia.Substring(4, 4),
            Convert.ToInt32(referencia.Substring(8, 9)).ToString(),
            referencia.Substring(17, 4),
            referencia.Substring(21, 2),
            referencia.Substring(23, 2),
            "01"
            };


            return refe;
        }

    }
}
