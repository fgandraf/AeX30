using System;

namespace AeX30.App.Services.ProposalVersion
{
    public abstract class PCI2023
    {
        public static readonly string[] References = new string[]
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
            "X100",    // [22]  Item 2
            "X101",    // [23]  Item 3
            "X102",    // [24]  Item 4
            "X103",    // [25]  Item 5
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
            "AM144",  // [41]  Pré-Exc.
            "AM145",  // [42]  Etapa 1
            "AM146",  // [43]  Etapa 2
            "AM147",  // [44]  Etapa 3
            "AM148",  // [45]  Etapa 4
            "AM149",  // [46]  Etapa 5
            "AM150",  // [47]  Etapa 6
            "AM151",  // [48]  Etapa 7
            "AM152",  // [49]  Etapa 8
            "AM153",  // [50]  Etapa 9
            "AM154",  // [51]  Etapa 10
            "AM155",  // [52]  Etapa 11
            "AM156",  // [53]  Etapa 12
            "AM157",  // [54]  Etapa 13
            "AM158",  // [55]  Etapa 14
            "AM159",  // [56]  Etapa 15
            "AM160",  // [57]  Etapa 16
            "AM161",  // [58]  Etapa 17
            "AM162",  // [59]  Etapa 18
            "AM163",  // [60]  Etapa 19
            "AM164",  // [61]  Etapa 20
            "AM165",  // [62]  Etapa 21
            "AM166",  // [63]  Etapa 22
            "AM167",  // [64]  Etapa 23
            "AM168",  // [65]  Etapa 24

            "AM168",  // [66]  Etapa 25
            "AM168",  // [67]  Etapa 26
            "AM168",  // [68]  Etapa 27
            "AM168",  // [69]  Etapa 28
            "AM168",  // [70]  Etapa 29
            "AM168"   // [71]  Etapa 30
        };
    }
}
