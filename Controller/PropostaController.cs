using aeX30.Entities;
using aeX30.Model;

namespace aeX30.Controller
{
    internal class PropostaController
    {


        internal PropostaEnt GetProposal(string path)
        {
            if (PropostaModel.isValid(path))
                return new PropostaModel().GetProposal(path);
            else
                return null;
        }



    }
}
