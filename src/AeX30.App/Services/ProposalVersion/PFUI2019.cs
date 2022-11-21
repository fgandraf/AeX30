using System;

namespace AeX30.App.Services.ProposalVersion
{
    public abstract class PFUI2019
    {
        public static readonly string[] References = new string[]
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
            "BN390",  // [66]  Etapa 25
            "BN390",  // [67]  Etapa 26
            "BN390",  // [68]  Etapa 27
            "BN390",  // [69]  Etapa 28
            "BN390",  // [70]  Etapa 29
            "BN390"   // [71]  Etapa 30  
        };
    }
}
