using aeX30.Entities;
using aeX30.Model;

namespace aeX30.Controller
{
    internal class ProposalController
    {

        internal Proposal GetProposal(string filePath)
        {

            if (IsValid(filePath))
                return new ProposalModel().GetProposal(filePath);
            else
                return null;
        }




        private static bool IsValid(string filePath)
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