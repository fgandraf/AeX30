using AeX30.Core.Entities;

namespace AeX30.Core.Interfaces
{
    public interface IProposalService
    {
        Proposal LoadFromFile(string filePath);

    }
}
