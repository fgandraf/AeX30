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
                Proposal proposal = new ProposalModel().GetProposal(filePath, cellReference);

                return FormatedProposal(proposal);
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



        private Proposal FormatedProposal(Proposal proposal)
        {
            proposal.ProponenteNome = proposal.ProponenteNome.TrimEnd();
            proposal.ProponenteCPF = FormatString.CPF(proposal.ProponenteCPF);
            proposal.ProponenteFone = FormatString.Fone(proposal.ProponenteFone);
            proposal.ResponsavelCPF = FormatString.CPF(proposal.ResponsavelCPF);
            proposal.ResponsavelFone = FormatString.Fone(proposal.ResponsavelFone);
            proposal.ImovelCep = FormatString.CEP(proposal.ImovelCep);

            if (proposal.Tipo == "Proposta")
                proposal.Tipo = "PFUI";
            else
            {
                proposal.Tipo = "PCI";
                proposal.ImovelComarca = "";
                proposal.ImovelComarcaUF = "";
            }

            return proposal;
        }


    }



}