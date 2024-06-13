
namespace AeX30.Core.Entities
{
    public class Report
    {
        public Request Request { get; private set; }
        public Proposal Proposal { get; private set; }
        public string MensuradoAcumulado { get; private set; }
        public string ContratoInicio { get; private set; }
        public string ContratoTermino { get; private set; }


        public Report(Request referencia, Proposal proposal, string mensuradoAcumulado, string contratoInicio, string contratoTermino)
        {
            Request = referencia;
            Proposal = proposal;
            MensuradoAcumulado = mensuradoAcumulado;
            ContratoInicio = contratoInicio;
            ContratoTermino = contratoTermino;
        }

        public string SuggestedFileName()
        {
            string[] proponente = Proposal.ProponenteNome.ToLower().Split(' ');
            string nome = proponente[0].Substring(0, 1).ToUpper() + proponente[0].Substring(1);
            string sobrenome = proponente[proponente.Length - 1].Substring(0, 1).ToUpper() + proponente[proponente.Length - 1].Substring(1);

            return $"RAE_{nome}-{sobrenome}.xlsx";
        }

    }
}