using AeX30.Model;

namespace AeX30.Controller
{
    public class ProposalController
    { 
     
        public Proposal GetProposal(string filePath)
        {
            if (IsValid(filePath))
            {
                string footer = Proposal.GetLeftFooter(filePath);
                string[] cellReference = new ProposalCellReference().Get(footer);
                return new Proposal().GetProposal(filePath, cellReference);
            }
            else
                return null;
        }

        private bool IsValid(string filePath)
        {
            string sheetName = Proposal.GetSheetName(filePath);
            string footer = Proposal.GetLeftFooter(filePath);
           
            bool sheetNameIsValid = (sheetName == "Proposta" || sheetName == "Proposta_Constr_Individual" ? true : false);
            bool footerIsValid = (footer != "" || footer != null ? true : false);
            bool cellReferenceIsValid = (new ProposalCellReference().Get(footer) != null ? true : false);
            
            if ( (sheetNameIsValid) && (footerIsValid) && (cellReferenceIsValid) )
                return true;
            else
                return false;
        }

    }
}