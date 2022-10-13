using AeX30.Model;

namespace AeX30.Controller
{
    public class ReportController
    {
        public static int SetReport(string pathTemplate, string pathDestin, Report report)
        {
            if (IsValid(pathTemplate))
                return new Report().SetReport(pathTemplate, pathDestin, report);
            else
                return 0;

        }


        private static bool IsValid(string filePath)
        {
            string sheetName = Report.GetSheetName(filePath);
            string footer = Report.GetLeftFooter(filePath);

            if (sheetName == "RAE" && footer != "" || footer != null)
                return true;
            else
                return false;
        }



    }
}







