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
