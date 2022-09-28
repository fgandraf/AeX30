using System;

namespace aeX30.Controller
{
    internal class RequestController
    {

        internal string[] GetReferencia(string path)
        {
            try
            {
                return new Model.RequestModel().GetReferencia(path);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }



    }
}
