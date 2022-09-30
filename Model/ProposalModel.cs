using System;
using System.IO;
using System.Text.RegularExpressions;
using aeX30.Entities;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace aeX30.Model
{
    public class ProposalModel
    {

        internal static string GetSheetName(string filePath)
        {
            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);

            return wbook.GetSheetName(0);
        }


        internal static string GetFooter(string filePath)
        {
            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);

            ISheet sheet = wbook.GetSheet(wbook.GetSheetName(0));

            return sheet.Footer.Left;
        }


        internal Proposal GetProposal(string filePath)
        {
            //Open File
            FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            HSSFWorkbook wbook = new HSSFWorkbook(file);
            ISheet sheet = wbook.GetSheetAt(0);


            //Set the right cell reference for the version
            string[] xy = SetCellReference(sheet.Footer.Left);


            //Create the object
            Proposal proposal = new Proposal
            {
                Vigencia          = sheet.Footer.Left,
                
                ProponenteNome    = sheet.GetRow(new CellReference(xy[0]).Row).GetCell(new CellReference(xy[0]).Col).ToString(),
                ProponenteCPF     = sheet.GetRow(new CellReference(xy[1]).Row).GetCell(new CellReference(xy[1]).Col).ToString(),
                ProponenteDDD     = sheet.GetRow(new CellReference(xy[2]).Row).GetCell(new CellReference(xy[2]).Col).ToString(),
                ProponenteFone    = sheet.GetRow(new CellReference(xy[3]).Row).GetCell(new CellReference(xy[3]).Col).ToString(),
                
                ResponsavelNome   = sheet.GetRow(new CellReference(xy[4]).Row).GetCell(new CellReference(xy[4]).Col).ToString(),
                ReponsavelCauCrea = sheet.GetRow(new CellReference(xy[5]).Row).GetCell(new CellReference(xy[5]).Col).ToString(),
                ResponsavelUF     = sheet.GetRow(new CellReference(xy[6]).Row).GetCell(new CellReference(xy[6]).Col).ToString(),
                ResponsavelCPF    = sheet.GetRow(new CellReference(xy[7]).Row).GetCell(new CellReference(xy[7]).Col).ToString(),
                ResponsavelDDD    = sheet.GetRow(new CellReference(xy[8]).Row).GetCell(new CellReference(xy[8]).Col).ToString(),
                ResponsavelFone   = sheet.GetRow(new CellReference(xy[9]).Row).GetCell(new CellReference(xy[9]).Col).ToString(),
                
                ImovelEndereco    = sheet.GetRow(new CellReference(xy[10]).Row).GetCell(new CellReference(xy[10]).Col).ToString(),
                ImovelComplemento = sheet.GetRow(new CellReference(xy[11]).Row).GetCell(new CellReference(xy[11]).Col).ToString(),
                ImovelCep         = sheet.GetRow(new CellReference(xy[12]).Row).GetCell(new CellReference(xy[12]).Col).ToString(),
                ImovelBairro      = sheet.GetRow(new CellReference(xy[13]).Row).GetCell(new CellReference(xy[13]).Col).ToString(),
                ImovelMunicipio   = sheet.GetRow(new CellReference(xy[14]).Row).GetCell(new CellReference(xy[14]).Col).ToString(),
                ImovelUF          = sheet.GetRow(new CellReference(xy[15]).Row).GetCell(new CellReference(xy[15]).Col).ToString(),
                ImovelMatricula   = sheet.GetRow(new CellReference(xy[17]).Row).GetCell(new CellReference(xy[17]).Col).ToString(),
                ImovelOficio      = sheet.GetRow(new CellReference(xy[18]).Row).GetCell(new CellReference(xy[18]).Col).ToString(),
                
                ServicoItem01     = sheet.GetRow(new CellReference(xy[21]).Row).GetCell(new CellReference(xy[21]).Col).NumericCellValue.ToString(),
                ServicoItem02     = sheet.GetRow(new CellReference(xy[22]).Row).GetCell(new CellReference(xy[22]).Col).NumericCellValue.ToString(),
                ServicoItem03     = sheet.GetRow(new CellReference(xy[23]).Row).GetCell(new CellReference(xy[23]).Col).NumericCellValue.ToString(),
                ServicoItem04     = sheet.GetRow(new CellReference(xy[24]).Row).GetCell(new CellReference(xy[24]).Col).NumericCellValue.ToString(),
                ServicoItem05     = sheet.GetRow(new CellReference(xy[25]).Row).GetCell(new CellReference(xy[25]).Col).NumericCellValue.ToString(),
                ServicoItem06     = sheet.GetRow(new CellReference(xy[26]).Row).GetCell(new CellReference(xy[26]).Col).NumericCellValue.ToString(),
                ServicoItem07     = sheet.GetRow(new CellReference(xy[27]).Row).GetCell(new CellReference(xy[27]).Col).NumericCellValue.ToString(),
                ServicoItem08     = sheet.GetRow(new CellReference(xy[28]).Row).GetCell(new CellReference(xy[28]).Col).NumericCellValue.ToString(),
                ServicoItem09     = sheet.GetRow(new CellReference(xy[29]).Row).GetCell(new CellReference(xy[29]).Col).NumericCellValue.ToString(),
                ServicoItem10     = sheet.GetRow(new CellReference(xy[30]).Row).GetCell(new CellReference(xy[30]).Col).NumericCellValue.ToString(),
                ServicoItem11     = sheet.GetRow(new CellReference(xy[31]).Row).GetCell(new CellReference(xy[31]).Col).NumericCellValue.ToString(),
                ServicoItem12     = sheet.GetRow(new CellReference(xy[32]).Row).GetCell(new CellReference(xy[32]).Col).NumericCellValue.ToString(),
                ServicoItem13     = sheet.GetRow(new CellReference(xy[33]).Row).GetCell(new CellReference(xy[33]).Col).NumericCellValue.ToString(),
                ServicoItem14     = sheet.GetRow(new CellReference(xy[34]).Row).GetCell(new CellReference(xy[34]).Col).NumericCellValue.ToString(),
                ServicoItem15     = sheet.GetRow(new CellReference(xy[35]).Row).GetCell(new CellReference(xy[35]).Col).NumericCellValue.ToString(),
                ServicoItem16     = sheet.GetRow(new CellReference(xy[36]).Row).GetCell(new CellReference(xy[36]).Col).NumericCellValue.ToString(),
                ServicoItem17     = sheet.GetRow(new CellReference(xy[37]).Row).GetCell(new CellReference(xy[37]).Col).NumericCellValue.ToString(),
                ServicoItem18     = sheet.GetRow(new CellReference(xy[38]).Row).GetCell(new CellReference(xy[38]).Col).NumericCellValue.ToString(),
                ServicoItem19     = sheet.GetRow(new CellReference(xy[39]).Row).GetCell(new CellReference(xy[39]).Col).NumericCellValue.ToString(),
                ServicoItem20     = sheet.GetRow(new CellReference(xy[40]).Row).GetCell(new CellReference(xy[40]).Col).NumericCellValue.ToString(),

                Cron_Executado    = sheet.GetRow(new CellReference(xy[41]).Row).GetCell(new CellReference(xy[41]).Col).NumericCellValue.ToString(),
                
                Cron_Parc_1       = sheet.GetRow(new CellReference(xy[42]).Row).GetCell(new CellReference(xy[42]).Col).NumericCellValue.ToString(),
                Cron_Parc_2       = sheet.GetRow(new CellReference(xy[43]).Row).GetCell(new CellReference(xy[43]).Col).NumericCellValue.ToString(),
                Cron_Parc_3       = sheet.GetRow(new CellReference(xy[44]).Row).GetCell(new CellReference(xy[44]).Col).NumericCellValue.ToString(),
                Cron_Parc_4       = sheet.GetRow(new CellReference(xy[45]).Row).GetCell(new CellReference(xy[45]).Col).NumericCellValue.ToString(),
                Cron_Parc_5       = sheet.GetRow(new CellReference(xy[46]).Row).GetCell(new CellReference(xy[46]).Col).NumericCellValue.ToString(),
                Cron_Parc_6       = sheet.GetRow(new CellReference(xy[47]).Row).GetCell(new CellReference(xy[47]).Col).NumericCellValue.ToString(),
                Cron_Parc_7       = sheet.GetRow(new CellReference(xy[48]).Row).GetCell(new CellReference(xy[48]).Col).NumericCellValue.ToString(),
                Cron_Parc_8       = sheet.GetRow(new CellReference(xy[49]).Row).GetCell(new CellReference(xy[49]).Col).NumericCellValue.ToString(),
                Cron_Parc_9       = sheet.GetRow(new CellReference(xy[50]).Row).GetCell(new CellReference(xy[50]).Col).NumericCellValue.ToString(),
                Cron_Parc_10      = sheet.GetRow(new CellReference(xy[51]).Row).GetCell(new CellReference(xy[52]).Col).NumericCellValue.ToString(),
                Cron_Parc_11      = sheet.GetRow(new CellReference(xy[52]).Row).GetCell(new CellReference(xy[52]).Col).NumericCellValue.ToString(),
                Cron_Parc_12      = sheet.GetRow(new CellReference(xy[53]).Row).GetCell(new CellReference(xy[53]).Col).NumericCellValue.ToString(),
                Cron_Parc_13      = sheet.GetRow(new CellReference(xy[54]).Row).GetCell(new CellReference(xy[54]).Col).NumericCellValue.ToString(),
                Cron_Parc_14      = sheet.GetRow(new CellReference(xy[55]).Row).GetCell(new CellReference(xy[55]).Col).NumericCellValue.ToString(),
                Cron_Parc_15      = sheet.GetRow(new CellReference(xy[56]).Row).GetCell(new CellReference(xy[56]).Col).NumericCellValue.ToString(),
                Cron_Parc_16      = sheet.GetRow(new CellReference(xy[57]).Row).GetCell(new CellReference(xy[57]).Col).NumericCellValue.ToString(),
                Cron_Parc_17      = sheet.GetRow(new CellReference(xy[58]).Row).GetCell(new CellReference(xy[58]).Col).NumericCellValue.ToString(),
                Cron_Parc_18      = sheet.GetRow(new CellReference(xy[59]).Row).GetCell(new CellReference(xy[59]).Col).NumericCellValue.ToString(),
                Cron_Parc_19      = sheet.GetRow(new CellReference(xy[60]).Row).GetCell(new CellReference(xy[60]).Col).NumericCellValue.ToString(),
                Cron_Parc_20      = sheet.GetRow(new CellReference(xy[61]).Row).GetCell(new CellReference(xy[61]).Col).NumericCellValue.ToString(),
                Cron_Parc_21      = sheet.GetRow(new CellReference(xy[62]).Row).GetCell(new CellReference(xy[62]).Col).NumericCellValue.ToString(),
                Cron_Parc_22      = sheet.GetRow(new CellReference(xy[63]).Row).GetCell(new CellReference(xy[63]).Col).NumericCellValue.ToString(),
                Cron_Parc_23      = sheet.GetRow(new CellReference(xy[64]).Row).GetCell(new CellReference(xy[64]).Col).NumericCellValue.ToString(),
                Cron_Parc_24      = sheet.GetRow(new CellReference(xy[65]).Row).GetCell(new CellReference(xy[65]).Col).NumericCellValue.ToString(),
                Cron_Parc_25      = sheet.GetRow(new CellReference(xy[66]).Row).GetCell(new CellReference(xy[66]).Col).NumericCellValue.ToString(),
                Cron_Parc_26      = sheet.GetRow(new CellReference(xy[67]).Row).GetCell(new CellReference(xy[67]).Col).NumericCellValue.ToString(),
                Cron_Parc_27      = sheet.GetRow(new CellReference(xy[68]).Row).GetCell(new CellReference(xy[68]).Col).NumericCellValue.ToString(),
                Cron_Parc_28      = sheet.GetRow(new CellReference(xy[69]).Row).GetCell(new CellReference(xy[69]).Col).NumericCellValue.ToString(),
                Cron_Parc_29      = sheet.GetRow(new CellReference(xy[70]).Row).GetCell(new CellReference(xy[70]).Col).NumericCellValue.ToString(),
                Cron_Parc_30      = sheet.GetRow(new CellReference(xy[71]).Row).GetCell(new CellReference(xy[71]).Col).NumericCellValue.ToString()
            };



            try
            {
                proposal.ImovelValorTerreno = string.Format("{0:0,0.00}", Convert.ToDouble(sheet.GetRow(new CellReference(xy[16]).Row).GetCell(new CellReference(xy[16]).Col).ToString()));
            }
            catch
            {
                proposal.ImovelValorTerreno = null;
            }
            

            if (wbook.GetSheetName(0) == "Proposta")
            {
                proposal.Tipo = "PFUI";
                proposal.ImovelComarca = sheet.GetRow(new CellReference(xy[19]).Row).GetCell(new CellReference(xy[19]).Col).ToString();
                proposal.ImovelComarcaUF = sheet.GetRow(new CellReference(xy[20]).Row).GetCell(new CellReference(xy[20]).Col).ToString();
            }
            else
                proposal.Tipo = "PCI";


            //Return the populated object
            return proposal;
        }


        private static string[] SetCellReference(string version)
        {
            var onlyNumber = new Regex(@"[^\d]");
            int convertedVersion = Convert.ToInt32(onlyNumber.Replace(version, ""));

            switch (convertedVersion)
            {
                case 927112017:
                    return pfui_2017;

                case 1102018:
                    return pfui_2018;

                case 26022018:
                    return pfui_2018b;

                case 12042019:
                case 23052019:
                case 24052019:
                    return pfui_2019;

                case 13072020:
                case 23072020:
                    return pfui_2020a;

                case 24092020:
                case 26022021:
                case 5052021:
                    return pfui_2020b;

                case 4122020:
                    return pfui_2020c;

                case 1062021:
                case 5072021:
                case 14072021:
                case 6082021:
                    return pci_2021a;

                case 21102021:
                case 4112021:
                    return pci_2021b;

                default:
                    return null;
            }
        }
 
        private static string[] pfui_2017 = new string[]
        {
            // PROPOSTA
            "G42",     // [0]  Proponente
            "AA42",    // [1]  CPF Prop.
            "AG42",    // [2]  Telefone Prop. (DDD)
            "AI42",    // [3]  Telefone Prop.
            "AN42",    // [4]  RT pela Execuçao da Obra - RTE"
            "AX42",    // [5]  Nº CAU/CREA/CFT - RTE
            "BD42",    // [6]  UF (RTE)
            "BF42",    // [7]  CPF - RTE
            "BL42",    // [8]  Telefone - RTE (DDD)
            "BN42",    // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G46",     // [10]  Endereço
            "AP46",    // [11]  Complemento
            "BA46",    // [12]  CEP
            "BF46",    // [13]  Bairro
            "G48",     // [14]  Município
            "AB48",    // [15]  UF
            "AG50",    // [16]  Terreno: valor proposto
            "AP50",    // [17]  Matrícula
            "AX50",    // [18]  ORI
            "BF50",    // [19]  Comarca
            "BN50",    // [20]  Comarca UF
           
            // SERVIÇOS
            "AR112",   // [21]  Item 17.1
            "AR114",   // [22]  Item 17.2
            "AR126",   // [23]  Item 17.3
            "AR133",   // [24]  Item 17.4
            "AR142",   // [25]  Item 17.5
            "AR152",   // [26]  Item 17.6
            "AR161",   // [27]  Item 17.7
            "AR168",   // [28]  Item 17.8
            "AR175",   // [29]  Item 17.9
            "AR187",   // [30]  Item 17.10
            "AR194",   // [31]  Item 17.11
            "AR205",   // [32]  Item 17.13 (Pintura)
            "AR215",   // [33]  Item 17.12 (Piso)
            "AR226",   // [34]  Item 17.14
            "AR232",   // [35]  Item 17.15
            "AR243",   // [36]  Item 17.16
            "AR252",   // [37]  Item 17.17
            "AR260",   // [38]  Item 17.18
            "AR269",   // [39]  Item 17.19
            "AR271",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ309",  // [41]  Pré-Exc.
            "AL308",  // [42]  Etapa 1
            "AP308",  // [43]  Etapa 2
            "AT308",  // [44]  Etapa 3
            "AX308",  // [45]  Etapa 4
            "BB308",  // [46]  Etapa 5
            "BF308",  // [47]  Etapa 6
            "BJ308",  // [48]  Etapa 7
            "BN308",  // [48]  Etapa 8
            "AL348",  // [50]  Etapa 9
            "AP348",  // [51]  Etapa 10
            "AT348",  // [52]  Etapa 11
            "AX348",  // [53]  Etapa 12
            "BB348",  // [54]  Etapa 13
            "BF348",  // [55]  Etapa 14
            "BJ348",  // [56]  Etapa 15
            "BN348",  // [57]  Etapa 16
            "AL388",  // [58]  Etapa 17
            "AP388",  // [59]  Etapa 18
            "AT388",  // [60]  Etapa 19
            "AX388",  // [61]  Etapa 20
            "BB388",  // [62]  Etapa 21
            "BF388",  // [63]  Etapa 22
            "BJ388",  // [64]  Etapa 23
            "BN388",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN388",  // [66]  Etapa 25
            "BN388",  // [67]  Etapa 26
            "BN388",  // [68]  Etapa 27
            "BN388",  // [69]  Etapa 28
            "BN388",  // [70]  Etapa 29
            "BN388"   // [71]  Etapa 30  
        };
        private static string[] pfui_2018 = new string[]
        {
            // PROPOSTA
            "G42",     // [0]  Proponente
            "AA42",    // [1]  CPF Prop.
            "AG42",    // [2]  Telefone Prop. (DDD)
            "AI42",    // [3]  Telefone Prop.
            "AN42",    // [4]  RT pela Execuçao da Obra - RTE"
            "AX42",    // [5]  Nº CAU/CREA/CFT - RTE
            "BD42",    // [6]  UF (RTE)
            "BF42",    // [7]  CPF - RTE
            "BM42",    // [8]  Telefone - RTE (DDD)
            "BN42",    // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G46",     // [10]  Endereço
            "AP46",    // [11]  Complemento
            "BA46",    // [12]  CEP
            "BF46",    // [13]  Bairro
            "G48",     // [14]  Município
            "AB48",    // [15]  UF
            "AG50",    // [16]  Terreno: valor proposto
            "AP50",    // [17]  Matrícula
            "AX50",    // [18]  ORI
            "BF50",    // [19]  Comarca
            "BN50",    // [20]  Comarca UF
           
            // SERVIÇOS
            "AR116",   // [21]  Item 17.1
            "AR118",   // [22]  Item 17.2
            "AR130",   // [23]  Item 17.3
            "AR137",   // [24]  Item 17.4
            "AR146",   // [25]  Item 17.5
            "AR156",   // [26]  Item 17.6
            "AR165",   // [27]  Item 17.7
            "AR172",   // [28]  Item 17.8
            "AR179",   // [29]  Item 17.9
            "AR191",   // [30]  Item 17.10
            "AR198",   // [31]  Item 17.11
            "AR209",   // [32]  Item 17.13 (Pintura)
            "AR219",   // [33]  Item 17.12 (Piso)
            "AR230",   // [34]  Item 17.14
            "AR236",   // [35]  Item 17.15
            "AR247",   // [36]  Item 17.16
            "AR256",   // [37]  Item 17.17
            "AR264",   // [38]  Item 17.18
            "AR273",   // [39]  Item 17.19
            "AR275",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ313",  // [41]  Pré-Exc.
            "AL312",  // [42]  Etapa 1
            "AP312",  // [43]  Etapa 2
            "AT312",  // [44]  Etapa 3
            "AX312",  // [45]  Etapa 4
            "BB312",  // [46]  Etapa 5
            "BF312",  // [47]  Etapa 6
            "BJ312",  // [48]  Etapa 7
            "BN312",  // [48]  Etapa 8
            "AL352",  // [50]  Etapa 9
            "AP352",  // [51]  Etapa 10
            "AT352",  // [52]  Etapa 11
            "AX352",  // [53]  Etapa 12
            "BB352",  // [54]  Etapa 13
            "BF352",  // [55]  Etapa 14
            "BJ352",  // [56]  Etapa 15
            "BN352",  // [57]  Etapa 16
            "AL392",  // [58]  Etapa 17
            "AP392",  // [59]  Etapa 18
            "AT392",  // [60]  Etapa 19
            "AX392",  // [61]  Etapa 20
            "BB392",  // [62]  Etapa 21
            "BF392",  // [63]  Etapa 22
            "BJ392",  // [64]  Etapa 23
            "BN392",  // [65]  Etapa 24
            ///"BN388", "BN388", "BN388", "BN388", "BN388", "BN388"
            "BN392",  // [66]  Etapa 25
            "BN392",  // [67]  Etapa 26
            "BN392",  // [68]  Etapa 27
            "BN392",  // [69]  Etapa 28
            "BN392",  // [70]  Etapa 29
            "BN392"   // [71]  Etapa 30   
        };
        private static string[] pfui_2018b = new string[]
        {
            // PROPOSTA
            "G42",     // [0]  Proponente
            "AA42",    // [1]  CPF Prop.
            "AG42",    // [2]  Telefone Prop. (DDD)
            "AI42",    // [3]  Telefone Prop.
            "AN42",    // [4]  RT pela Execuçao da Obra - RTE"
            "AX42",    // [5]  Nº CAU/CREA/CFT - RTE
            "BD42",    // [6]  UF (RTE)
            "BF42",    // [7]  CPF - RTE
            "BM42",    // [8]  Telefone - RTE (DDD)
            "BN42",    // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G46",     // [10]  Endereço
            "AP46",    // [11]  Complemento
            "BA46",    // [12]  CEP
            "BF46",    // [13]  Bairro
            "G48",     // [14]  Município
            "AB48",    // [15]  UF
            "AG50",    // [16]  Terreno: valor proposto
            "AP50",    // [17]  Matrícula
            "AX50",    // [18]  ORI
            "BF50",    // [19]  Comarca
            "BN50",    // [20]  Comarca UF
           
            // SERVIÇOS
            "AR112",   // [21]  Item 17.1
            "AR114",   // [22]  Item 17.2
            "AR126",   // [23]  Item 17.3
            "AR133",   // [24]  Item 17.4
            "AR142",   // [25]  Item 17.5
            "AR152",   // [26]  Item 17.6
            "AR161",   // [27]  Item 17.7
            "AR168",   // [28]  Item 17.8
            "AR175",   // [29]  Item 17.9
            "AR187",   // [30]  Item 17.10
            "AR194",   // [31]  Item 17.11
            "AR205",   // [32]  Item 17.13 (Pintura)
            "AR215",   // [33]  Item 17.12 (Piso)
            "AR226",   // [34]  Item 17.14
            "AR232",   // [35]  Item 17.15
            "AR243",   // [36]  Item 17.16
            "AR252",   // [37]  Item 17.17
            "AR260",   // [38]  Item 17.18
            "AR269",   // [39]  Item 17.19
            "AR271",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ309",  // [41]  Pré-Exc.
            "AL308",  // [42]  Etapa 1
            "AP308",  // [43]  Etapa 2
            "AT308",  // [44]  Etapa 3
            "AX308",  // [45]  Etapa 4
            "BB308",  // [46]  Etapa 5
            "BF308",  // [47]  Etapa 6
            "BJ308",  // [48]  Etapa 7
            "BN308",  // [48]  Etapa 8
            "AL348",  // [50]  Etapa 9
            "AP348",  // [51]  Etapa 10
            "AT348",  // [52]  Etapa 11
            "AX348",  // [53]  Etapa 12
            "BB348",  // [54]  Etapa 13
            "BF348",  // [55]  Etapa 14
            "BJ348",  // [56]  Etapa 15
            "BN348",  // [57]  Etapa 16
            "AL388",  // [58]  Etapa 17
            "AP388",  // [59]  Etapa 18
            "AT388",  // [60]  Etapa 19
            "AX388",  // [61]  Etapa 20
            "BB388",  // [62]  Etapa 21
            "BF388",  // [63]  Etapa 22
            "BJ388",  // [64]  Etapa 23
            "BN388",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN388",  // [66]  Etapa 25
            "BN388",  // [67]  Etapa 26
            "BN388",  // [68]  Etapa 27
            "BN388",  // [69]  Etapa 28
            "BN388",  // [70]  Etapa 29
            "BN388"   // [71]  Etapa 30  
        };
        private static string[] pfui_2019 = new string[]
        {
            // PROPOSTA
            "G42",     // [0]  Proponente
            "AA42",    // [1]  CPF Prop.
            "AG42",    // [2]  Telefone Prop. (DDD)
            "AI42",    // [3]  Telefone Prop.
            "AN42",    // [4]  RT pela Execuçao da Obra - RTE"
            "AX42",    // [5]  Nº CAU/CREA/CFT - RTE
            "BD42",    // [6]  UF (RTE)
            "BF42",    // [7]  CPF - RTE
            "BM42",    // [8]  Telefone - RTE (DDD)
            "BN42",    // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G46",     // [10]  Endereço
            "AP46",    // [11]  Complemento
            "BA46",    // [12]  CEP
            "BF46",    // [13]  Bairro
            "G48",     // [14]  Município
            "AB48",    // [15]  UF
            "AG50",    // [16]  Terreno: valor proposto
            "AP50",    // [17]  Matrícula
            "AX50",    // [18]  ORI
            "BF50",    // [19]  Comarca
            "BN50",    // [20]  Comarca UF
           
            // SERVIÇOS
            "AR116",   // [21]  Item 17.1
            "AR118",   // [22]  Item 17.2
            "AR130",   // [23]  Item 17.3
            "AR137",   // [24]  Item 17.4
            "AR146",   // [25]  Item 17.5
            "AR156",   // [26]  Item 17.6
            "AR165",   // [27]  Item 17.7
            "AR172",   // [28]  Item 17.8
            "AR179",   // [29]  Item 17.9
            "AR190",   // [30]  Item 17.10
            "AR197",   // [31]  Item 17.11
            "AR207",   // [32]  Item 17.13 (Pintura)
            "AR217",   // [33]  Item 17.12 (Piso)
            "AR228",   // [34]  Item 17.14
            "AR234",   // [35]  Item 17.15
            "AR245",   // [36]  Item 17.16
            "AR254",   // [37]  Item 17.17
            "AR262",   // [38]  Item 17.18
            "AR271",   // [39]  Item 17.19
            "AR273",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ311",  // [41]  Pré-Exc.
            "AL310",  // [42]  Etapa 1
            "AP310",  // [43]  Etapa 2
            "AT310",  // [44]  Etapa 3
            "AX310",  // [45]  Etapa 4
            "BB310",  // [46]  Etapa 5
            "BF310",  // [47]  Etapa 6
            "BJ310",  // [48]  Etapa 7
            "BN310",  // [48]  Etapa 8
            "AL350",  // [50]  Etapa 9
            "AP350",  // [51]  Etapa 10
            "AT350",  // [52]  Etapa 11
            "AX350",  // [53]  Etapa 12
            "BB350",  // [54]  Etapa 13
            "BF350",  // [55]  Etapa 14
            "BJ350",  // [56]  Etapa 15
            "BN350",  // [57]  Etapa 16
            "AL390",  // [58]  Etapa 17
            "AP390",  // [59]  Etapa 18
            "AT390",  // [60]  Etapa 19
            "AX390",  // [61]  Etapa 20
            "BB390",  // [62]  Etapa 21
            "BF390",  // [63]  Etapa 22
            "BJ390",  // [64]  Etapa 23
            "BN390",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN390",  // [66]  Etapa 25
            "BN390",  // [67]  Etapa 26
            "BN390",  // [68]  Etapa 27
            "BN390",  // [69]  Etapa 28
            "BN390",  // [70]  Etapa 29
            "BN390"   // [71]  Etapa 30  
  
        };
        private static string[] pfui_2020a = new string[]
        {

            // PROPOSTA
            "G42",     // [0]  Proponente
            "AA42",    // [1]  CPF Prop.
            "AG42",    // [2]  Telefone Prop. (DDD)
            "AI42",    // [3]  Telefone Prop.
            "AN42",    // [4]  RT pela Execuçao da Obra - RTE"
            "AX42",    // [5]  Nº CAU/CREA/CFT - RTE
            "BD42",    // [6]  UF (RTE)
            "BF42",    // [7]  CPF - RTE
            "BM42",    // [8]  Telefone - RTE (DDD)
            "BN42",    // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G46",     // [10]  Endereço
            "AP46",    // [11]  Complemento
            "BA46",    // [12]  CEP
            "BF46",    // [13]  Bairro
            "G48",     // [14]  Município
            "AB48",    // [15]  UF
            "AG50",    // [16]  Terreno: valor proposto
            "AP50",    // [17]  Matrícula
            "AX50",    // [18]  ORI
            "BF50",    // [19]  Comarca
            "BN50",    // [20]  Comarca UF
           
            // SERVIÇOS
            "AR116",   // [21]  Item 17.1
            "AR117",   // [22]  Item 17.2
            "AR130",   // [23]  Item 17.3
            "AR137",   // [24]  Item 17.4
            "AR146",   // [25]  Item 17.5
            "AR156",   // [26]  Item 17.6
            "AR165",   // [27]  Item 17.7
            "AR172",   // [28]  Item 17.8
            "AR179",   // [29]  Item 17.9
            "AR190",   // [30]  Item 17.10
            "AR197",   // [31]  Item 17.11
            "AR207",   // [32]  Item 17.13 (Pintura)
            "AR217",   // [33]  Item 17.12 (Piso)
            "AR228",   // [34]  Item 17.14
            "AR234",   // [35]  Item 17.15
            "AR245",   // [36]  Item 17.16
            "AR254",   // [37]  Item 17.17
            "AR262",   // [38]  Item 17.18
            "AR271",   // [39]  Item 17.19
            "AR273",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ312",  // [41]  Pré-Exc.
            "AL311",  // [42]  Etapa 1
            "AP311",  // [43]  Etapa 2
            "AT311",  // [44]  Etapa 3
            "AX311",  // [45]  Etapa 4
            "BB311",  // [46]  Etapa 5
            "BF311",  // [47]  Etapa 6
            "BJ311",  // [48]  Etapa 7
            "BN311",  // [48]  Etapa 8
            "AL351",  // [50]  Etapa 9
            "AP351",  // [51]  Etapa 10
            "AT351",  // [52]  Etapa 11
            "AX351",  // [53]  Etapa 12
            "BB351",  // [54]  Etapa 13
            "BF351",  // [55]  Etapa 14
            "BJ351",  // [56]  Etapa 15
            "BN351",  // [57]  Etapa 16
            "AL391",  // [58]  Etapa 17
            "AP391",  // [59]  Etapa 18
            "AT391",  // [60]  Etapa 19
            "AX391",  // [61]  Etapa 20
            "BB391",  // [62]  Etapa 21
            "BF391",  // [63]  Etapa 22
            "BJ391",  // [64]  Etapa 23
            "BN391",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN391",  // [66]  Etapa 25
            "BN391",  // [67]  Etapa 26
            "BN391",  // [68]  Etapa 27
            "BN391",  // [69]  Etapa 28
            "BN391",  // [70]  Etapa 29
            "BN391"   // [71]  Etapa 30  
        };
        private static string[] pfui_2020b = new string[]
        {
            // PROPOSTA
            "G42",     // [0]  Proponente
            "AA42",    // [1]  CPF Prop.
            "AG42",    // [2]  Telefone Prop. (DDD)
            "AI42",    // [3]  Telefone Prop.
            "AN42",    // [4]  RT pela Execuçao da Obra - RTE"
            "AX42",    // [5]  Nº CAU/CREA/CFT - RTE
            "BD42",    // [6]  UF (RTE)
            "BF42",    // [7]  CPF - RTE
            "BM42",    // [8]  Telefone - RTE (DDD)
            "BN42",    // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G46",     // [10]  Endereço
            "AP46",    // [11]  Complemento
            "BA46",    // [12]  CEP
            "BF46",    // [13]  Bairro
            "G48",     // [14]  Município
            "AB48",    // [15]  UF
            "AG50",    // [16]  Terreno: valor proposto
            "AP50",    // [17]  Matrícula
            "AX50",    // [18]  ORI
            "BF50",    // [19]  Comarca
            "BN50",    // [20]  Comarca UF
           
            // SERVIÇOS
            "AR114",   // [21]  Item 17.1
            "AR116",   // [22]  Item 17.2
            "AR128",   // [23]  Item 17.3
            "AR135",   // [24]  Item 17.4
            "AR144",   // [25]  Item 17.5
            "AR154",   // [26]  Item 17.6
            "AR163",   // [27]  Item 17.7
            "AR170",   // [28]  Item 17.8
            "AR177",   // [29]  Item 17.9
            "AR188",   // [30]  Item 17.10
            "AR195",   // [31]  Item 17.11
            "AR205",   // [32]  Item 17.13 (Pintura)
            "AR215",   // [33]  Item 17.12 (Piso)
            "AR226",   // [34]  Item 17.14
            "AR232",   // [35]  Item 17.15
            "AR243",   // [36]  Item 17.16
            "AR252",   // [37]  Item 17.17
            "AR260",   // [38]  Item 17.18
            "AR269",   // [39]  Item 17.19
            "AR271",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ310",  // [41]  Pré-Exc.
            "AL309",  // [42]  Etapa 1
            "AP309",  // [43]  Etapa 2
            "AT309",  // [44]  Etapa 3
            "AX309",  // [45]  Etapa 4
            "BB309",  // [46]  Etapa 5
            "BF309",  // [47]  Etapa 6
            "BJ309",  // [48]  Etapa 7
            "BN309",  // [48]  Etapa 8
            "AL349",  // [50]  Etapa 9
            "AP349",  // [51]  Etapa 10
            "AT349",  // [52]  Etapa 11
            "AX349",  // [53]  Etapa 12
            "BB349",  // [54]  Etapa 13
            "BF349",  // [55]  Etapa 14
            "BJ349",  // [56]  Etapa 15
            "BN349",  // [57]  Etapa 16
            "AL389",  // [58]  Etapa 17
            "AP389",  // [59]  Etapa 18
            "AT389",  // [60]  Etapa 19
            "AX389",  // [61]  Etapa 20
            "BB389",  // [62]  Etapa 21
            "BF389",  // [63]  Etapa 22
            "BJ389",  // [64]  Etapa 23
            "BN389",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN389",  // [66]  Etapa 25
            "BN389",  // [67]  Etapa 26
            "BN389",  // [68]  Etapa 27
            "BN389",  // [69]  Etapa 28
            "BN389",  // [70]  Etapa 29
            "BN389"   // [71]  Etapa 30  
        };
        private static string[] pfui_2020c = new string[]
        {
            // PROPOSTA
            "G42",     // [0]  Proponente
            "AA42",    // [1]  CPF Prop.
            "AG42",    // [2]  Telefone Prop. (DDD)
            "AI42",    // [3]  Telefone Prop.
            "AN42",    // [4]  RT pela Execuçao da Obra - RTE"
            "AX42",    // [5]  Nº CAU/CREA/CFT - RTE
            "BD42",    // [6]  UF (RTE)
            "BF42",    // [7]  CPF - RTE
            "BM42",    // [8]  Telefone - RTE (DDD)
            "BN42",    // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G46",     // [10]  Endereço
            "AP46",    // [11]  Complemento
            "BA46",    // [12]  CEP
            "BF46",    // [13]  Bairro
            "G48",     // [14]  Município
            "AB48",    // [15]  UF
            "AG50",    // [16]  Terreno: valor proposto
            "AP50",    // [17]  Matrícula
            "AX50",    // [18]  ORI
            "BF50",    // [19]  Comarca
            "BN50",    // [20]  Comarca UF
           
            // SERVIÇOS
            "AR115",   // [21]  Item 17.1
            "AR117",   // [22]  Item 17.2
            "AR129",   // [23]  Item 17.3
            "AR136",   // [24]  Item 17.4
            "AR145",   // [25]  Item 17.5
            "AR155",   // [26]  Item 17.6
            "AR164",   // [27]  Item 17.7
            "AR171",   // [28]  Item 17.8
            "AR178",   // [29]  Item 17.9
            "AR189",   // [30]  Item 17.10
            "AR196",   // [31]  Item 17.11
            "AR206",   // [32]  Item 17.13 (Pintura)
            "AR216",   // [33]  Item 17.12 (Piso)
            "AR227",   // [34]  Item 17.14
            "AR233",   // [35]  Item 17.15
            "AR244",   // [36]  Item 17.16
            "AR253",   // [37]  Item 17.17
            "AR261",   // [38]  Item 17.18
            "AR270",   // [39]  Item 17.19
            "AR272",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ311",  // [41]  Pré-Exc.
            "AL310",  // [42]  Etapa 1
            "AP310",  // [43]  Etapa 2
            "AT310",  // [44]  Etapa 3
            "AX310",  // [45]  Etapa 4
            "BB310",  // [46]  Etapa 5
            "BF310",  // [47]  Etapa 6
            "BJ310",  // [48]  Etapa 7
            "BN310",  // [48]  Etapa 8
            "AL350",  // [50]  Etapa 9
            "AP350",  // [51]  Etapa 10
            "AT350",  // [52]  Etapa 11
            "AX350",  // [53]  Etapa 12
            "BB350",  // [54]  Etapa 13
            "BF350",  // [55]  Etapa 14
            "BJ350",  // [56]  Etapa 15
            "BN350",  // [57]  Etapa 16
            "AL390",  // [58]  Etapa 17
            "AP390",  // [59]  Etapa 18
            "AT390",  // [60]  Etapa 19
            "AX390",  // [61]  Etapa 20
            "BB390",  // [62]  Etapa 21
            "BF390",  // [63]  Etapa 22
            "BJ390",  // [64]  Etapa 23
            "BN390",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "AL430",  // [66]  Etapa 25
            "AP430",  // [67]  Etapa 26
            "AT430",  // [68]  Etapa 27
            "AX430",  // [69]  Etapa 28
            "BB430",  // [70]  Etapa 29
            "BF430"   // [71]  Etapa 30  
        };
        private static string[] pci_2021a = new string[]
        {
            // IDENTIFICAÇÃO
            "G43",   // [0]  Proponente
            "AK43",  // [1]  CPF Prop.
            "AQ43",  // [2]  Telefone Prop. (DDD)
            "AS43",  // [3]  Telefone Prop.
            "G49",   // [4]  RT pela Execuçao da Obra - RTE"
            "AB49",  // [5]  Nº CAU/CREA/CFT - RTE
            "AI49",  // [6]  UF (RTE)
            "AK49",  // [7]  CPF - RTE
            "AQ49",  // [8]  Telefone - RTE (DDD)
            "AS49",  // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G53",   // [10]  Endereço
            "AJ53",  // [11]  Complemento
            "V55",   // [12]  CEP
            "G55",   // [13]  Bairro
            "AA55",  // [14]  Município
            "AV55",  // [15]  UF
            "V70",   // [16]  Valor do terreno
            "G57",   // [17]  Matrícula
            "M57",   // [18]  ORI
            "-",     // [19]  Comarca
            "-",     // [20]  Comarca UF
           
            // SERVIÇOS 
            "X94",    // [21]  Item 1
            "X95",    // [22]  Item 2
            "X96",    // [23]  Item 3
            "X97",    // [24]  Item 4
            "X98",    // [25]  Item 5
            "X99",    // [26]  Item 6
            "X100",   // [27]  Item 7
            "X101",   // [28]  Item 8
            "X102",   // [29]  Item 9
            "X103",   // [30]  Item 10
            "X104",   // [31]  Item 11
            "X106",   // [32]  Item 13 (Pintura)
            "X105",   // [33]  Item 12 (Piso)
            "X107",   // [34]  Item 14
            "X108",   // [35]  Item 15
            "X109",   // [36]  Item 16
            "X110",   // [37]  Item 17
            "X111",   // [38]  Item 18
            "X112",   // [39]  Item 19
            "X113",   // [40]  Item 20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AM139",  // [41]  Pré-Exc.
            "AM140",  // [42]  Etapa 1
            "AM141",  // [43]  Etapa 2
            "AM142",  // [44]  Etapa 3
            "AM143",  // [45]  Etapa 4
            "AM144",  // [46]  Etapa 5
            "AM145",  // [47]  Etapa 6
            "AM146",  // [48]  Etapa 7
            "AM147",  // [49]  Etapa 8
            "AM148",  // [50]  Etapa 9
            "AM149",  // [51]  Etapa 10
            "AM150",  // [52]  Etapa 11
            "AM151",  // [53]  Etapa 12
            "AM152",  // [54]  Etapa 13
            "AM153",  // [55]  Etapa 14
            "AM154",  // [56]  Etapa 15
            "AM155",  // [57]  Etapa 16
            "AM156",  // [58]  Etapa 17
            "AM157",  // [59]  Etapa 18
            "AM158",  // [60]  Etapa 19
            "AM159",  // [61]  Etapa 20
            "AM160",  // [62]  Etapa 21
            "AM161",  // [63]  Etapa 22
            "AM162",  // [64]  Etapa 23
            "AM163",  // [65]  Etapa 24

            "AM164",  // [66]  Etapa 25
            "AM164",  // [67]  Etapa 26
            "AM164",  // [68]  Etapa 27
            "AM164",  // [69]  Etapa 28
            "AM164",  // [70]  Etapa 29
            "AM164"   // [71]  Etapa 30
        };
        private static string[] pci_2021b = new string[]
        {
            // IDENTIFICAÇÃO
            "G43",   // [0]  Proponente
            "AK43",  // [1]  CPF Prop.
            "AQ43",  // [2]  Telefone Prop. (DDD)
            "AS43",  // [3]  Telefone Prop.
            "G49",   // [4]  RT pela Execuçao da Obra - RTE"
            "AB49",  // [5]  Nº CAU/CREA/CFT - RTE
            "AI49",  // [6]  UF (RTE)
            "AK49",  // [7]  CPF - RTE
            "AQ49",  // [8]  Telefone - RTE (DDD)
            "AS49",  // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G53",   // [10]  Endereço
            "AJ53",  // [11]  Complemento
            "V55",   // [12]  CEP
            "G55",   // [13]  Bairro
            "AA55",  // [14]  Município
            "AV55",  // [15]  UF
            "V70",   // [16]  Valor do terreno
            "G57",   // [17]  Matrícula
            "M57",   // [18]  ORI
            "-",     // [19]  Comarca
            "-",     // [20]  Comarca UF
           
            // SERVIÇOS
            "X95",    // [21]  Item 1
            "X96",    // [22]  Item 2
            "X97",    // [23]  Item 3
            "X98",    // [24]  Item 4
            "X99",    // [25]  Item 5
            "X100",   // [26]  Item 6
            "X101",   // [27]  Item 7
            "X102",   // [28]  Item 8
            "X103",   // [29]  Item 9
            "X104",   // [30]  Item 10
            "X105",   // [31]  Item 11
            "X106",   // [32]  Item 13 (Pintura)
            "X107",   // [33]  Item 12 (Piso)
            "X108",   // [34]  Item 14
            "X109",   // [35]  Item 15
            "X110",   // [36]  Item 16
            "X111",   // [37]  Item 17
            "X112",   // [38]  Item 18
            "X113",   // [39]  Item 19
            "X114",   // [40]  Item 20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AM140",  // [41]  Pré-Exc.
            "AM141",  // [42]  Etapa 1
            "AM142",  // [43]  Etapa 2
            "AM143",  // [44]  Etapa 3
            "AM144",  // [45]  Etapa 4
            "AM145",  // [46]  Etapa 5
            "AM146",  // [47]  Etapa 6
            "AM147",  // [48]  Etapa 7
            "AM148",  // [49]  Etapa 8
            "AM149",  // [50]  Etapa 9
            "AM150",  // [51]  Etapa 10
            "AM151",  // [52]  Etapa 11
            "AM152",  // [53]  Etapa 12
            "AM153",  // [54]  Etapa 13
            "AM154",  // [55]  Etapa 14
            "AM155",  // [56]  Etapa 15
            "AM156",  // [57]  Etapa 16
            "AM157",  // [58]  Etapa 17
            "AM158",  // [59]  Etapa 18
            "AM159",  // [60]  Etapa 19
            "AM160",  // [61]  Etapa 20
            "AM161",  // [62]  Etapa 21
            "AM162",  // [63]  Etapa 22
            "AM163",  // [64]  Etapa 23
            "AM164",  // [65]  Etapa 24

            "AM165",  // [66]  Etapa 25
            "AM165",  // [67]  Etapa 26
            "AM165",  // [68]  Etapa 27
            "AM165",  // [69]  Etapa 28
            "AM165",  // [70]  Etapa 29
            "AM165"   // [71]  Etapa 30
        };


    }
}
