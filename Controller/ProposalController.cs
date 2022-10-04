using AeX30.Model.Entities;
using AeX30.Model;

namespace AeX30.Controller
{
    public class ProposalController
    {
        

        public Proposal GetProposal(string filePath)
        {

            if (IsValid(filePath))
            {
                string footer = ProposalModel.GetFooter(filePath);
                string[] cellReference = new ProposalCellReference().Get(footer);
                return new ProposalModel().GetProposal(filePath, cellReference);
            }
            else
                return null;

            
        }



        private bool IsValid(string filePath)
        {
            string sheetName = ProposalModel.GetSheetName(filePath);
            string footer = ProposalModel.GetFooter(filePath);

            if ((sheetName == "Proposta" || sheetName == "Proposta_Constr_Individual") && (footer != "" || footer != null))
                return true;
            else
                return false;
        }



    }



}