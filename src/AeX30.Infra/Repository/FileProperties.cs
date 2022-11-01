using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace AeX30.Infra.Repository
{
    public class FileProperties
    {
        public static string GetSheetName(string filePath)
        {
            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);

            return wbook.GetSheetName(0);
        }

        public static string GetLeftFooter(string filePath)
        {
            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);
            ISheet sheet = wbook.GetSheet(wbook.GetSheetName(0));

            return sheet.Footer.Left;
        }

    }
}
