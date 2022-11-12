using System;

namespace AeX30.Services.ProposalVersion
{
    public abstract class PFUI2020b
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
            "BL42",    // [8]  Telefone - RTE (DDD)
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
            "AR114",   // [21]  Item 17.1
            "AR116",   // [22]  Item 17.2
            "AR128",   // [23]  Item 17.3
            "AR135",   // [24]  Item 17.4
            "AR144",   // [25]  Item 17.5
            "AR154",   // [26]  Item 17.6
            "AR163",   // [27]  Item 17.7
            "AR170",   // [28]  Item 17.8
            "AR177",   // [29]  Item 17.9
            "AR188",   // [30]  Item 17.10
            "AR195",   // [31]  Item 17.11
            "AR205",   // [32]  Item 17.13 (Pintura)
            "AR215",   // [33]  Item 17.12 (Piso)
            "AR226",   // [34]  Item 17.14
            "AR232",   // [35]  Item 17.15
            "AR243",   // [36]  Item 17.16
            "AR252",   // [37]  Item 17.17
            "AR260",   // [38]  Item 17.18
            "AR269",   // [39]  Item 17.19
            "AR271",   // [40]  Item 17.20

            // CRONOGRAMA FÍSICO FINANCEIRO
            "AJ310",  // [41]  Pré-Exc.
            "AL309",  // [42]  Etapa 1
            "AP309",  // [43]  Etapa 2
            "AT309",  // [44]  Etapa 3
            "AX309",  // [45]  Etapa 4
            "BB309",  // [46]  Etapa 5
            "BF309",  // [47]  Etapa 6
            "BJ309",  // [48]  Etapa 7
            "BN309",  // [48]  Etapa 8
            "AL349",  // [50]  Etapa 9
            "AP349",  // [51]  Etapa 10
            "AT349",  // [52]  Etapa 11
            "AX349",  // [53]  Etapa 12
            "BB349",  // [54]  Etapa 13
            "BF349",  // [55]  Etapa 14
            "BJ349",  // [56]  Etapa 15
            "BN349",  // [57]  Etapa 16
            "AL389",  // [58]  Etapa 17
            "AP389",  // [59]  Etapa 18
            "AT389",  // [60]  Etapa 19
            "AX389",  // [61]  Etapa 20
            "BB389",  // [62]  Etapa 21
            "BF389",  // [63]  Etapa 22
            "BJ389",  // [64]  Etapa 23
            "BN389",  // [65]  Etapa 24
            "BN389",  // [66]  Etapa 25
            "BN389",  // [67]  Etapa 26
            "BN389",  // [68]  Etapa 27
            "BN389",  // [69]  Etapa 28
            "BN389",  // [70]  Etapa 29
            "BN389"   // [71]  Etapa 30  
        };
    }
}
