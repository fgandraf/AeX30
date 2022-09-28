using aeX30.Entities;
using aeX30.Model;

namespace aeX30.Controller
{
    internal class ReportController
    {



        internal static int SetReport(string pathTemplate, string pathDestin, Report report)
        {
            if (ReportModel.IsValid(pathTemplate))
                return new ReportModel().SetReport(pathTemplate, pathDestin, report);
            else
                return 0;

        }



    }
}







