using System;
using aeX30.Entities;
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
                Part1 = fullNumber.Substring(0, 4),
                Part2 = fullNumber.Substring(5, 4),
                Part3 = fullNumber.Substring(10, 9).TrimStart('0'),
                Part4 = fullNumber.Substring(20, 4),
                Part5 = fullNumber.Substring(25, 2),
                Part6 = fullNumber.Substring(28, 2),
                Part7 = "01"
            };

        }

    }
}
