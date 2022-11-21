using AeX30.Domain.Entities;
using System.IO;
using AeX30.Infra.Repository;

namespace AeX30.App.Services
{
    public class ProposalService
    {
        private string[] _cellReference;


        public Proposal GetProposal(string filePath)
        {
            if (IsValid(filePath))
                return new ProposalRepository().GetProposal(filePath, _cellReference);

            return null;
        }

        private bool IsValid(string filePath)
        {
            if (File.Exists(filePath))
            {
                string footer = FileProperties.GetLeftFooter(filePath);
                _cellReference = ProposalCellReference.Get(footer);
                string sheetName = FileProperties.GetSheetName(filePath);

                bool sheetNameIsValid = sheetName == "Proposta" || sheetName == "Proposta_Constr_Individual";
                bool footerIsValid = !string.IsNullOrEmpty(footer);
                bool cellReferenceIsValid = _cellReference != null;

                return sheetNameIsValid && footerIsValid && cellReferenceIsValid;
            }

            return false;
        }

    }
}