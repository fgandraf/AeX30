using aeX30.Model.Entities;
using aeX30.Model;

namespace aeX30.Controller
{
    internal class ReportController
    {



        internal static int SetReport(string pathTemplate, string pathDestin, Report report)
        {
            if (IsValid(pathTemplate))
                return new ReportModel().SetReport(pathTemplate, pathDestin, report);
            else
                return 0;

        }



        private static bool IsValid(string filePath)
        {
            string sheetName = ReportModel.GetSheetName(filePath);
            string footer = ReportModel.GetFooter(filePath);

            if (sheetName == "RAE" && footer != "" || footer != null)
                return true;
            else
                return false;
        }



    }
}







