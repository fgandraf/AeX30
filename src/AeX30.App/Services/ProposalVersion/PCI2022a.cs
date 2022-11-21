using System;

namespace AeX30.App.Services.ProposalVersion
{
    public abstract class PCI2022a
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
            "AR71",  // [16]  Valor do terreno
            "G57",   // [17]  Matrícula
            "M57",   // [18]  ORI
            "-",     // [19]  Comarca
            "-",     // [20]  Comarca UF
           
            // SERVIÇOS
            "X98",    // [21]  Item 1
            "X99",    // [22]  Item 2
            "X100",    // [23]  Item 3
            "X101",    // [24]  Item 4
            "X102",    // [25]  Item 5
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
            "AM143",  // [41]  Pré-Exc.
            "AM144",  // [42]  Etapa 1
            "AM145",  // [43]  Etapa 2
            "AM146",  // [44]  Etapa 3
            "AM147",  // [45]  Etapa 4
            "AM148",  // [46]  Etapa 5
            "AM149",  // [47]  Etapa 6
            "AM150",  // [48]  Etapa 7
            "AM151",  // [49]  Etapa 8
            "AM152",  // [50]  Etapa 9
            "AM153",  // [51]  Etapa 10
            "AM154",  // [52]  Etapa 11
            "AM155",  // [53]  Etapa 12
            "AM156",  // [54]  Etapa 13
            "AM157",  // [55]  Etapa 14
            "AM158",  // [56]  Etapa 15
            "AM159",  // [57]  Etapa 16
            "AM160",  // [58]  Etapa 17
            "AM161",  // [59]  Etapa 18
            "AM162",  // [60]  Etapa 19
            "AM163",  // [61]  Etapa 20
            "AM164",  // [62]  Etapa 21
            "AM165",  // [63]  Etapa 22
            "AM166",  // [64]  Etapa 23
            "AM167",  // [65]  Etapa 24

            "AM167",  // [66]  Etapa 25
            "AM167",  // [67]  Etapa 26
            "AM167",  // [68]  Etapa 27
            "AM167",  // [69]  Etapa 28
            "AM167",  // [70]  Etapa 29
            "AM167"   // [71]  Etapa 30
        };
    }
}
