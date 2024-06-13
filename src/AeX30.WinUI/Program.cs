using System;
using System.Windows.Forms;
using AeX30.Core.Services;
using AeX30.WinUI.View;
using OfficeOpenXml;

namespace AeX30.WinUI

{
    static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            Application.Run(
                new FormMain(
                    new RequestService(), 
                    new ProposalService(), 
                    new ReportService()
                    )
                );
        }
    }
}
