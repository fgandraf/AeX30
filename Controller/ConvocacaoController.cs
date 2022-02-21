using System;

namespace aeX30.Controller
{
    internal class ConvocacaoController
    {

        internal string[] GetReferencia(string path)
        {
            try
            {
                return new Model.ConvocacaoModel().GetReferencia(path);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



    }
}
