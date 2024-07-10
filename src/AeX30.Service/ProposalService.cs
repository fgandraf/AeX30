using AeX30.Core.Entities;

namespace AeX30.Services
{
    public static class ProposalService
    {
        public static Proposal LoadFromFile(string filePath)
        {
            var tipo = Path.GetExtension(filePath).TrimStart('.');
            
            return tipo == "xlsx" 
                ? new XlsXProposal().LoadFromFile(filePath)
                : new XlsProposal().LoadFromFile(filePath);
        }

    }
}
