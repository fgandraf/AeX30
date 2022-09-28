using aeX30.Entities;
using aeX30.Model;

namespace aeX30.Controller
{
    internal class ProposalController
    {



        internal Proposal GetProposal(string path)
        {

            if (ProposalModel.IsValid(path))
                return new ProposalModel().GetProposal(path);
            else
                return null;
        }



    }



}