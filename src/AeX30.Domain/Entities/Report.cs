using System.IO;
using System;

namespace AeX30.Domain.Entities
{
    public class Report : Proposal
    {
        public Report()
        {
            Referencia = new string[7];
        }

        public string[] Referencia { get; set; }
        public string MensuradoAcumulado { get; set; }
        public string ContratoInicio { get; set; }
        public string ContratoTermino { get; set; }
    }
}