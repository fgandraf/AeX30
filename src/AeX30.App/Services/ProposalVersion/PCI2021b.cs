using System;

namespace AeX30.App.Services.ProposalVersion
{
    public abstract class PCI2021b
    {
        public static readonly string[] References = new string[]
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
    }
}
