using AeX30.Domain.Entities;
using System.IO;
using AeX30.Infra.Repository;

namespace AeX30.App.Services
{
    public class ProposalService
    {
        
        public bool IsValid { get; private set; }
        private string[] _cellReference;
        private string _filePath;
        private ProposalRepository _proposalRepository;

        public ProposalService(string filePath)
        {
            _filePath = filePath;
            _proposalRepository = new ProposalRepository();
            Validate();
        }

        public void Validate()
        {
            if (File.Exists(_filePath))
            {
                string footer = _proposalRepository.GetLeftFooter(_filePath);
                _cellReference = ProposalCellReference.Get(footer);
                string sheetName = _proposalRepository.GetSheetName(_filePath);

                bool sheetNameIsValid = sheetName == "Proposta" || sheetName == "Proposta_Constr_Individual";
                bool footerIsValid = !string.IsNullOrEmpty(footer);
                bool cellReferenceIsValid = _cellReference != null;

                IsValid = sheetNameIsValid && footerIsValid && cellReferenceIsValid;
            }
        }

        public Proposal GetProposal()
        {
            if (IsValid)
                return _proposalRepository.GetProposal(_filePath, _cellReference);

            return null;
        }

        

    }
}