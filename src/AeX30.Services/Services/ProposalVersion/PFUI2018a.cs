using System;

namespace AeX30.Services.ProposalVersion
{
    public abstract class PFUI2018a
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
            "BN392",  // [66]  Etapa 25
            "BN392",  // [67]  Etapa 26
            "BN392",  // [68]  Etapa 27
            "BN392",  // [69]  Etapa 28
            "BN392",  // [70]  Etapa 29
            "BN392"   // [71]  Etapa 30   
        };
    }
}
