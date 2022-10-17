using AeX30.Entities;
using System.IO;
using AeX30.Repository;

namespace AeX30.Controller
{
    public class ProposalController
    { 
     
        public Proposal GetProposal(string filePath)
        {
            if (IsValid(filePath))
                return new ProposalRepository().GetProposal(filePath);
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