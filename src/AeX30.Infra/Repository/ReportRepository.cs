using AeX30.Domain.Entities;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.IO;

namespace AeX30.Infra.Repository
{
    public class ReportRepository
    {
        public static void SetReport(string pathTemplate, string pathDestin, Report report)
        {
            string[] cellReference = new string[]
        {
            /*auto. de serv.:------*/    "AB35", "AD35", "AF35", "AJ35", "AL35", "AM35",
            
            /*prop. NOME:----------*/    "G43",
            /*prop. CPF:-----------*/    "AJ43",
            /*prop. DDD:-----------*/    "AP43",
            /*prop. TELEFONE:------*/    "AR43",
            
            /*rt. NOME:------------*/    "G46",
            /*rt. CAU_CREA:--------*/    "Z46",
            /*rt. UF:--------------*/    "AH46",
            /*rt. CPF:-------------*/    "AJ46",
            /*rt. DDD:-------------*/    "AP46",
            /*rt. TELEFONE:--------*/    "AR46",

            /*end. ENDEREÇO:-------*/    "G49",
            /*end. COMPLEMENTO:----*/    "AJ49",
            /*end. BAIRRO:---------*/    "G51",
            /*end. CEP:------------*/    "V51",
            /*end. MUNICIPIO:------*/    "AA51",
            /*end. UF:-------------*/    "AS51",
            
            /*imov. VALOR TERRENO:-*/    "G53",
            /*imov. MATRICULA:-----*/    "Q53",
            /*imov. OFICIO:--------*/    "AA53",
            /*imov. COMARCA:-------*/    "AJ53",
            /*imov. UF:------------*/    "AS53",
           
            /*17.01(%):------------*/    "S68",
            /*17.02(%):------------*/    "S69",
            /*17.03(%):------------*/    "S70",
            /*17.04(%):------------*/    "S71",
            /*17.05(%):------------*/    "S72",
            /*17.06(%):------------*/    "S73",
            /*17.07(%):------------*/    "S74",
            /*17.08(%):------------*/    "S75",
            /*17.09(%):------------*/    "S76",
            /*17.10(%):------------*/    "S77",
            /*17.11(%):------------*/    "S78",
            /*17.12(%):------------*/    "S79",
            /*17.13(%):------------*/    "S80",
            /*17.14(%):------------*/    "S81",
            /*17.15(%):------------*/    "S82",
            /*17.16(%):------------*/    "S83",
            /*17.17(%):------------*/    "S84",
            /*17.18(%):------------*/    "S85",
            /*17.19(%):------------*/    "S86",
            /*17.20(%):------------*/    "S87",
            /*ACUMULADO(%):--------*/    "Y89",

            /*contrato INICIO:-----*/    "AH63",
            /*contrato TERMINO:----*/    "AS63",

            /*cron. EXECUTADO:-----*/    "AK71",
            /*cron. PARC 1:--------*/    "AK72",
            /*cron. PARC 2:--------*/    "AK73",
            /*cron. PARC 3:--------*/    "AK74",
            /*cron. PARC 4:--------*/    "AK75",
            /*cron. PARC 5:--------*/    "AK76",
            /*cron. PARC 6:--------*/    "AK77",
            /*cron. PARC 7:--------*/    "AK78",
            /*cron. PARC 8:--------*/    "AK79",
            /*cron. PARC 9:--------*/    "AK80",
            /*cron. PARC 10:-------*/    "AK81",
            /*cron. PARC 11:-------*/    "AK82",
            /*cron. PARC 12:-------*/    "AK83",
            /*cron. PARC 13:-------*/    "AK84",
            /*cron. PARC 14:-------*/    "AK85",
            /*cron. PARC 15:-------*/    "AK86",
            /*cron. PARC 16:-------*/    "AK87",
            /*cron. PARC 17:-------*/    "AK88",
            /*cron. PARC 18:-------*/    "AK89",
            /*cron. PARC 19:-------*/    "AK90",
            /*cron. PARC 20:-------*/    "AK91",
            /*cron. PARC 21:-------*/    "AK92",
            /*cron. PARC 22:-------*/    "AK93",
            /*cron. PARC 23:-------*/    "AK94",
            /*cron. PARC 24:-------*/    "AK95",
            /*cron. PARC 25:-------*/    "AK96",
            /*cron. PARC 26:-------*/    "AK97",
            /*cron. PARC 27:-------*/    "AK98",
            /*cron. PARC 28:-------*/    "AK99",
            /*cron. PARC 29:-------*/    "AK100",
            /*cron. PARC 30:-------*/    "AK101",
            /*cron. PARC 31:-------*/    "AK102",
            /*cron. PARC 32:-------*/    "AK103",
            /*cron. PARC 33:-------*/    "AK104",
            /*cron. PARC 34:-------*/    "AK105",
            /*cron. PARC 35:-------*/    "AK106",
            /*cron. PARC 36:-------*/    "AK107",

        };
            dynamic[] values = report.Get();
            HSSFWorkbook wbook = new HSSFWorkbook();

            try
            {
                using (FileStream templateFile = new FileStream(pathTemplate, FileMode.Open, FileAccess.Read))
                {
                    wbook = new HSSFWorkbook(templateFile);
                    ISheet sheet = wbook.GetSheet("RAE");
                    for (int i = 0; i < values.Length; i++)
                        sheet.GetRow(new CellReference(cellReference[i]).Row).GetCell(new CellReference(cellReference[i]).Col).SetCellValue(values[i]);
                }

                using (FileStream reportFile = new FileStream(pathDestin, FileMode.Create, FileAccess.Write))
                {
                    wbook.ForceFormulaRecalculation = true;
                    wbook.Write(reportFile);
                }
            }
            catch (Exception ex)
            {
                throw ex.InnerException;
            }
        }
    }
}
