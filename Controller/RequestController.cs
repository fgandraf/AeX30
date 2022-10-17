using AeX30.Entities;
using AeX30.Repository;
using System.IO;

namespace AeX30.Controller
{
    public class RequestController
    {
        public Request GetRequestNumber(string filePath)
        {
            if (IsValid(filePath))
                return new RequestRepository().GetRequestNumber(filePath);
            else
                return null;
            
        }


        private bool IsValid(string filePath)
        {
            bool fileExists = File.Exists(filePath);
            bool fileNameValid = new FileInfo(filePath).Name == "CONVOCACAO PRESTACAO SERVICO.TXT";

            return fileExists && fileNameValid;
        }



    }
}
