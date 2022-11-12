using System;

namespace AeX30.Services.ProposalVersion
{
    public abstract class PCI2021a
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
            "-",   // [19]  Comarca
            "-",   // [20]  Comarca UF
           
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
    }
}
