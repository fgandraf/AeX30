using AeX30.Model.Entities;
using AeX30.Model;

namespace AeX30.Controller
{
    internal class RequestController
    {

        internal Request GetRequestNumber(string path)
        {
            return new RequestModel().GetRequestNumber(path);
        }


    }
}
