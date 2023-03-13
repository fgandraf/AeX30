using System;

namespace AeX30.Domain.Entities
{
    public class Report
    {
        public Report(Request referencia, Proposal proposal, string mensuradoAcumulado, string contratoInicio, string contratoTermino)
        {
            this.Request = referencia;
            this.Proposal = proposal;
            MensuradoAcumulado = mensuradoAcumulado;
            ContratoInicio = contratoInicio;
            ContratoTermino = contratoTermino;
        }

        public Request Request { get; private set; }
        public Proposal Proposal { get; private set; }
        public string MensuradoAcumulado { get; private set; }
        public string ContratoInicio { get; private set; }
        public string ContratoTermino { get; private set; }
    }
}