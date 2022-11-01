using AeX30.Domain.Entities;
using System.IO;
using AeX30.Infra.Repository;
using NPOI.SS.UserModel;

namespace AeX30.Services.Services
{
    public class ProposalService
    { 
     
        public Proposal GetProposal(string filePath)
        {
            if (IsValid(filePath))
            {
                string footer = ProposalRepository.GetLeftFooter(filePath);
                string[] cellReference = new ProposalCellReference().Get(footer);
                
                Proposal proposal = new ProposalRepository().GetProposal(filePath, cellReference);
                
                proposal.ProponenteCPF = FormatString.CPF(proposal.ProponenteCPF);
                proposal.ProponenteFone = FormatString.Fone(proposal.ProponenteFone);
                proposal.ResponsavelCPF = FormatString.CPF(proposal.ResponsavelCPF);
                proposal.ResponsavelFone = FormatString.Fone(proposal.ResponsavelFone);
                proposal.ImovelCep = FormatString.CEP(proposal.ImovelCep);
                proposal.ImovelValorTerreno = FormatString.ValorMonetario(proposal.ImovelValorTerreno);
                
                return proposal;
            }               
            else
                return null;
        }

        private bool IsValid(string filePath)
        {
            string footer = ProposalRepository.GetLeftFooter(filePath);
            string sheetName = ProposalRepository.GetSheetName(filePath);

            bool fileExists = File.Exists(filePath);
            bool sheetNameIsValid = sheetName == "Proposta" || sheetName == "Proposta_Constr_Individual";
            bool footerIsValid = footer != "" || footer != null;
            bool cellReferenceIsValid = new ProposalCellReference().Get(footer) != null;

            return fileExists && sheetNameIsValid && footerIsValid && cellReferenceIsValid;
        }

    }
}