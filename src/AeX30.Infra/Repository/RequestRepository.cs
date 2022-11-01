using AeX30.Domain.Entities;
using System.IO;
using System.Linq;

namespace AeX30.Infra.Repository
{

    public class RequestRepository
    {
        // [0000].[0000].[00000000]/[0000].[00].[00]
        // [Ref1].[Ref2].[Ref3]/[Ref4].[Ref5].[Ref6]

        public Request GetRequestNumber(string path)
        {
            var line = File.ReadAllLines(path)
                               .Where(l => l.StartsWith("Refer"))
                               .Select(l => l.Substring(l.LastIndexOf("-") + 2))
                               .ToList();

            string fullNumber = line[0].TrimStart('0');

            Request requestReference = new Request();
            requestReference.Referencia[1] = fullNumber.Substring(0, 4);
            requestReference.Referencia[2] = fullNumber.Substring(5, 4);
            requestReference.Referencia[3] = fullNumber.Substring(10, 9).TrimStart('0');
            requestReference.Referencia[4] = fullNumber.Substring(20, 4);
            requestReference.Referencia[5] = fullNumber.Substring(25, 2);
            requestReference.Referencia[6] = fullNumber.Substring(28, 2);

            return requestReference;

        }

    }
}
