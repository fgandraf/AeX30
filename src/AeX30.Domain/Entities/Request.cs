using System;

namespace AeX30.Domain.Entities
{
    public class Request
    {
        public Request()
        {
            Referencia = new string[7];
        }
        public string[] Referencia { get; set; }
    }
}
