using aeX30.Entities;
using aeX30.Model;

namespace aeX30.Controller
{
    internal class ProposalController
    {

        internal Proposal GetProposal(string filePath)
        {

            if (IsValid(filePath))
            {
                Proposal proposal = new ProposalModel().GetProposal(filePath);
                return RegexObject(proposal);
            }
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



        private static Proposal RegexObject(Proposal prop)
        {
            prop.ProponenteCPF = Util.FormatedCPF(prop.ProponenteCPF);
            prop.ProponenteFone = Util.FormatedFone(prop.ProponenteFone);
            prop.ProponenteCPF = Util.FormatedCPF(prop.ProponenteCPF);
            prop.ResponsavelCPF = Util.FormatedCPF(prop.ResponsavelCPF);
            prop.ImovelCep = Util.FormatedCEP(prop.ImovelCep);

            return prop;
        }


    }



}