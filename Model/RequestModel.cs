using aeX30.Model.Entities;
using System.IO;
using System.Linq;

namespace aeX30.Model
{
    internal class RequestModel
    {
        internal Request GetRequestNumber(string path)
        {
            var line = File.ReadAllLines(path)
                               .Where(l => l.StartsWith("Refer"))
                               .Select(l => l.Substring(l.LastIndexOf("-") + 2))
                               .ToList();


            string fullNumber = line[0].TrimStart('0');
            
            return new Request
            {
                Ref1 = fullNumber.Substring(0, 4),
                Ref2 = fullNumber.Substring(5, 4),
                Ref3 = fullNumber.Substring(10, 9).TrimStart('0'),
                Ref4 = fullNumber.Substring(20, 4),
                Ref5 = fullNumber.Substring(25, 2),
                Ref6 = fullNumber.Substring(28, 2),
            };

        }

    }
}
