
namespace AeX30.Core.Services
{
    public abstract class ReportCellMap
    {
        public static string[] Get()
            => cellReference;

        private static string[] cellReference = new string[]
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
            /*end. COMPLEMENTO:----*/    "AL49",
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

            /*cron. PARC 1:--------*/    "AG71",
            /*cron. EXECUTADO:-----*/    "AG72",
            /*cron. PARC 2:--------*/    "AG73",
            /*cron. PARC 3:--------*/    "AG74",
            /*cron. PARC 4:--------*/    "AG75",
            /*cron. PARC 5:--------*/    "AG76",
            /*cron. PARC 6:--------*/    "AG77",
            /*cron. PARC 7:--------*/    "AG78",
            /*cron. PARC 8:--------*/    "AG79",
            /*cron. PARC 9:--------*/    "AG80",
            /*cron. PARC 10:-------*/    "AG81",
            /*cron. PARC 11:-------*/    "AG82",
            /*cron. PARC 12:-------*/    "AG83",
            /*cron. PARC 13:-------*/    "AG84",
            /*cron. PARC 14:-------*/    "AG85",
            /*cron. PARC 15:-------*/    "AG86",
            /*cron. PARC 16:-------*/    "AG87",
            /*cron. PARC 17:-------*/    "AG88",
            /*cron. PARC 18:-------*/    "AG89",
            /*cron. PARC 19:-------*/    "AG90",
            /*cron. PARC 20:-------*/    "AG91",
            /*cron. PARC 21:-------*/    "AG92",
            /*cron. PARC 22:-------*/    "AG93",
            /*cron. PARC 23:-------*/    "AG94",
            /*cron. PARC 24:-------*/    "AG95",
            /*cron. PARC 25:-------*/    "AG96",
            /*cron. PARC 26:-------*/    "AG97",
            /*cron. PARC 27:-------*/    "AG98",
            /*cron. PARC 28:-------*/    "AG99",
            /*cron. PARC 29:-------*/    "AG100",
            /*cron. PARC 30:-------*/    "AG101",
            /*cron. PARC 31:-------*/    "AG102",
            /*cron. PARC 32:-------*/    "AG103",
            /*cron. PARC 33:-------*/    "AG104",
            /*cron. PARC 34:-------*/    "AG105",
            /*cron. PARC 35:-------*/    "AG106",
            /*cron. PARC 36:-------*/    "AG107",

        };
    }
}
