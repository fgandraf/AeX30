using System;

namespace AeX30.Services.ProposalVersion
{
    public abstract class PFUI2020a
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
            "BN391",  // [66]  Etapa 25
            "BN391",  // [67]  Etapa 26
            "BN391",  // [68]  Etapa 27
            "BN391",  // [69]  Etapa 28
            "BN391",  // [70]  Etapa 29
            "BN391"   // [71]  Etapa 30  
        };
    }
}
