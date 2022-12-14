using System;

namespace AeX30.Domain.Entities
{
    public class Request
    {
        public Request(string[] referencia)
        {
            Referencia = referencia;
        }
        public string[] Referencia { get; private set; }
    }
}
