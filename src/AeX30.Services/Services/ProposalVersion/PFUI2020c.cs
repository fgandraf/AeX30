using System;

namespace AeX30.Services.ProposalVersion
{
    public abstract class PFUI2020c
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
            "BN390",  // [66]  Etapa 25
            "BN390",  // [67]  Etapa 26
            "BN390",  // [68]  Etapa 27
            "BN390",  // [69]  Etapa 28
            "BN390",  // [70]  Etapa 29
            "BN390"   // [71]  Etapa 30  
        };
    }
}
