﻿using aeX30.Model.Entities;
using aeX30.Model;

namespace aeX30.Controller
{
    internal class RequestController
    {

        internal Request GetRequestNumber(string path)
        {
            return new RequestModel().GetRequestNumber(path);
        }


    }
}
