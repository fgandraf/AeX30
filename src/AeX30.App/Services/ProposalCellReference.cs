
namespace AeX30.App.Services
{
    public abstract class ProposalCellReference
    {
        public static string[] Get(string footer)
        {
            switch (footer)
            {
                case "Vigência: 01/06/2021":
                case "Vigência: 05/07/2021":
                case "Vigência: 14/07/2021":
                case "Vigência: 06/08/2021":
                    return _pci2021a;

                case "Vigência: 21/10/2021":
                case "Vigência: 04/11/2021":
                case "Vigência: 28/03/2022":
                    return _pci2021b;

                case "Vigência: 04/05/2022":
                case "Vigência: 08/06/2022":
                case "Vigência: 28/06/2022":
                    return _pci2022;

                case "Vigência: 11/08/2023":
                case "Vigência: 10/11/2023":
                    return _pci2023;

                default:
                    return null;
            }
        }


        private static readonly string[] _pci2021a = new string[]
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
            "AO139",  // [41]  Pré-Exc.
            "AO140",  // [42]  Etapa 1
            "AO141",  // [43]  Etapa 2
            "AO142",  // [44]  Etapa 3
            "AO143",  // [45]  Etapa 4
            "AO144",  // [46]  Etapa 5
            "AO145",  // [47]  Etapa 6
            "AO146",  // [48]  Etapa 7
            "AO147",  // [49]  Etapa 8
            "AO148",  // [50]  Etapa 9
            "AO149",  // [51]  Etapa 10
            "AO150",  // [52]  Etapa 11
            "AO151",  // [53]  Etapa 12
            "AO152",  // [54]  Etapa 13
            "AO153",  // [55]  Etapa 14
            "AO154",  // [56]  Etapa 15
            "AO155",  // [57]  Etapa 16
            "AO156",  // [58]  Etapa 17
            "AO157",  // [59]  Etapa 18
            "AO158",  // [60]  Etapa 19
            "AO159",  // [61]  Etapa 20
            "AO160",  // [62]  Etapa 21
            "AO161",  // [63]  Etapa 22
            "AO162",  // [64]  Etapa 23
            "AO163",  // [65]  Etapa 24

            "AO164",  // [66]  Etapa 25
            "AO164",  // [67]  Etapa 26
            "AO164",  // [68]  Etapa 27
            "AO164",  // [69]  Etapa 28
            "AO164",  // [70]  Etapa 29
            "AO164"   // [71]  Etapa 30
        };

        public static readonly string[] _pci2021b = new string[]
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
            "AO140",  // [41]  Pré-Exc.
            "AO141",  // [42]  Etapa 1
            "AO142",  // [43]  Etapa 2
            "AO143",  // [44]  Etapa 3
            "AO144",  // [45]  Etapa 4
            "AO145",  // [46]  Etapa 5
            "AO146",  // [47]  Etapa 6
            "AO147",  // [48]  Etapa 7
            "AO148",  // [49]  Etapa 8
            "AO149",  // [50]  Etapa 9
            "AO150",  // [51]  Etapa 10
            "AO151",  // [52]  Etapa 11
            "AO152",  // [53]  Etapa 12
            "AO153",  // [54]  Etapa 13
            "AO154",  // [55]  Etapa 14
            "AO155",  // [56]  Etapa 15
            "AO156",  // [57]  Etapa 16
            "AO157",  // [58]  Etapa 17
            "AO158",  // [59]  Etapa 18
            "AO159",  // [60]  Etapa 19
            "AO160",  // [61]  Etapa 20
            "AO161",  // [62]  Etapa 21
            "AO162",  // [63]  Etapa 22
            "AO163",  // [64]  Etapa 23
            "AO164",  // [65]  Etapa 24

            "AO165",  // [66]  Etapa 25
            "AO165",  // [67]  Etapa 26
            "AO165",  // [68]  Etapa 27
            "AO165",  // [69]  Etapa 28
            "AO165",  // [70]  Etapa 29
            "AO165"   // [71]  Etapa 30
        };

        public static readonly string[] _pci2022 = new string[]
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
            "AR71",  // [16]  Valor do terreno
            "G57",   // [17]  Matrícula
            "M57",   // [18]  ORI
            "-",     // [19]  Comarca
            "-",     // [20]  Comarca UF
           
            // SERVIÇOS
            "X98",    // [21]  Item 1
            "X99",    // [22]  Item 2
            "X100",   // [23]  Item 3
            "X101",   // [24]  Item 4
            "X102",   // [25]  Item 5
            "X103",   // [26]  Item 6
            "X104",   // [27]  Item 7
            "X105",   // [28]  Item 8
            "X106",   // [29]  Item 9
            "X107",   // [30]  Item 10
            "X108",   // [31]  Item 11
            "X109",   // [32]  Item 13 (Pintura)
            "X110",   // [33]  Item 12 (Piso)
            "X111",   // [34]  Item 14
            "X112",   // [35]  Item 15
            "X113",   // [36]  Item 16
            "X114",   // [37]  Item 17
            "X115",   // [38]  Item 18
            "X116",   // [39]  Item 19
            "X117",   // [40]  Item 20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AO143",  // [41]  Pré-Exc.
            "AO144",  // [42]  Etapa 1
            "AO145",  // [43]  Etapa 2
            "AO146",  // [44]  Etapa 3
            "AO147",  // [45]  Etapa 4
            "AO148",  // [46]  Etapa 5
            "AO149",  // [47]  Etapa 6
            "AO150",  // [48]  Etapa 7
            "AO151",  // [49]  Etapa 8
            "AO152",  // [50]  Etapa 9
            "AO153",  // [51]  Etapa 10
            "AO154",  // [52]  Etapa 11
            "AO155",  // [53]  Etapa 12
            "AO156",  // [54]  Etapa 13
            "AO157",  // [55]  Etapa 14
            "AO158",  // [56]  Etapa 15
            "AO159",  // [57]  Etapa 16
            "AO160",  // [58]  Etapa 17
            "AO161",  // [59]  Etapa 18
            "AO162",  // [60]  Etapa 19
            "AO163",  // [61]  Etapa 20
            "AO164",  // [62]  Etapa 21
            "AO165",  // [63]  Etapa 22
            "AO166",  // [64]  Etapa 23
            "AO167",  // [65]  Etapa 24

            "AO167",  // [66]  Etapa 25
            "AO167",  // [67]  Etapa 26
            "AO167",  // [68]  Etapa 27
            "AO167",  // [69]  Etapa 28
            "AO167",  // [70]  Etapa 29
            "AO167"   // [71]  Etapa 30
        };

        public static readonly string[] _pci2023 = new string[]
        {
            // IDENTIFICAÇÃO
            "G44",   // [0]  Proponente
            "AK44",  // [1]  CPF Prop.
            "AQ44",  // [2]  Telefone Prop. (DDD)
            "AS44",  // [3]  Telefone Prop.
            "G50",   // [4]  RT pela Execuçao da Obra - RTE"
            "AB50",  // [5]  Nº CAU/CREA/CFT - RTE
            "AI50",  // [6]  UF (RTE)
            "AK50",  // [7]  CPF - RTE
            "AQ50",  // [8]  Telefone - RTE (DDD)
            "AS50",  // [9]  Telefone - RTE

            // IDENTIFICAÇÃO DO IMÓVEL PROPOSTO
            "G54",   // [10]  Endereço
            "AJ54",  // [11]  Complemento
            "V56",   // [12]  CEP
            "G56",   // [13]  Bairro
            "AA56",  // [14]  Município
            "AV56",  // [15]  UF
            "AR72",  // [16]  Valor do terreno
            "G58",   // [17]  Matrícula
            "M58",   // [18]  ORI
            "-",     // [19]  Comarca
            "-",     // [20]  Comarca UF
           
            // SERVIÇOS
            "X99",    // [21]  Item 1
            "X100",   // [22]  Item 2
            "X101",   // [23]  Item 3
            "X102",   // [24]  Item 4
            "X103",   // [25]  Item 5
            "X104",   // [26]  Item 6
            "X105",   // [27]  Item 7
            "X106",   // [28]  Item 8
            "X107",   // [29]  Item 9
            "X108",   // [30]  Item 10
            "X109",   // [31]  Item 11
            "X110",   // [32]  Item 13 (Pintura)
            "X111",   // [33]  Item 12 (Piso)
            "X112",   // [34]  Item 14
            "X113",   // [35]  Item 15
            "X114",   // [36]  Item 16
            "X115",   // [37]  Item 17
            "X116",   // [38]  Item 18
            "X117",   // [39]  Item 19
            "X118",   // [40]  Item 20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AO144",  // [41]  Pré-Exc.
            "AO145",  // [42]  Etapa 1
            "AO146",  // [43]  Etapa 2
            "AO147",  // [44]  Etapa 3
            "AO148",  // [45]  Etapa 4
            "AO149",  // [46]  Etapa 5
            "AO150",  // [47]  Etapa 6
            "AO151",  // [48]  Etapa 7
            "AO152",  // [49]  Etapa 8
            "AO153",  // [50]  Etapa 9
            "AO154",  // [51]  Etapa 10
            "AO155",  // [52]  Etapa 11
            "AO156",  // [53]  Etapa 12
            "AO157",  // [54]  Etapa 13
            "AO158",  // [55]  Etapa 14
            "AO159",  // [56]  Etapa 15
            "AO160",  // [57]  Etapa 16
            "AO161",  // [58]  Etapa 17
            "AO162",  // [59]  Etapa 18
            "AO163",  // [60]  Etapa 19
            "AO164",  // [61]  Etapa 20
            "AO165",  // [62]  Etapa 21
            "AO166",  // [63]  Etapa 22
            "AO167",  // [64]  Etapa 23
            "AO168",  // [65]  Etapa 24
              
            "AO168",  // [66]  Etapa 25
            "AO168",  // [67]  Etapa 26
            "AO168",  // [68]  Etapa 27
            "AO168",  // [69]  Etapa 28
            "AO168",  // [70]  Etapa 29
            "AO168"   // [71]  Etapa 30
        };

    }
}
