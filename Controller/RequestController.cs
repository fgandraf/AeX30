using AeX30.Model;

namespace AeX30.Controller
{
    public class RequestController
    {

        public string[] GetRequestNumber(string path)
        {
            return new Request().GetRequestNumber(path);
        }


    }
}
