using System;

namespace AeX30.Core.Entities
{
    public class Request
    {
        public Request(string[] reference)
        {
            Reference = reference;
        }

        public string[] Reference { get; private set; }
    }
}
