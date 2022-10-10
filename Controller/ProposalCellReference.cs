using System;
using System.Text.RegularExpressions;

namespace AeX30.Controller
{
    public class ProposalCellReference
    {
        private readonly string[] PFUI2017 = new string[]
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
            "AR112",   // [21]  Item 17.1
            "AR114",   // [22]  Item 17.2
            "AR126",   // [23]  Item 17.3
            "AR133",   // [24]  Item 17.4
            "AR142",   // [25]  Item 17.5
            "AR152",   // [26]  Item 17.6
            "AR161",   // [27]  Item 17.7
            "AR168",   // [28]  Item 17.8
            "AR175",   // [29]  Item 17.9
            "AR187",   // [30]  Item 17.10
            "AR194",   // [31]  Item 17.11
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
            "AJ309",  // [41]  Pré-Exc.
            "AL308",  // [42]  Etapa 1
            "AP308",  // [43]  Etapa 2
            "AT308",  // [44]  Etapa 3
            "AX308",  // [45]  Etapa 4
            "BB308",  // [46]  Etapa 5
            "BF308",  // [47]  Etapa 6
            "BJ308",  // [48]  Etapa 7
            "BN308",  // [48]  Etapa 8
            "AL348",  // [50]  Etapa 9
            "AP348",  // [51]  Etapa 10
            "AT348",  // [52]  Etapa 11
            "AX348",  // [53]  Etapa 12
            "BB348",  // [54]  Etapa 13
            "BF348",  // [55]  Etapa 14
            "BJ348",  // [56]  Etapa 15
            "BN348",  // [57]  Etapa 16
            "AL388",  // [58]  Etapa 17
            "AP388",  // [59]  Etapa 18
            "AT388",  // [60]  Etapa 19
            "AX388",  // [61]  Etapa 20
            "BB388",  // [62]  Etapa 21
            "BF388",  // [63]  Etapa 22
            "BJ388",  // [64]  Etapa 23
            "BN388",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN388",  // [66]  Etapa 25
            "BN388",  // [67]  Etapa 26
            "BN388",  // [68]  Etapa 27
            "BN388",  // [69]  Etapa 28
            "BN388",  // [70]  Etapa 29
            "BN388"   // [71]  Etapa 30  
           };
        private readonly string[] PFUI2018a = new string[]
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
            ///"BN388", "BN388", "BN388", "BN388", "BN388", "BN388"
            "BN392",  // [66]  Etapa 25
            "BN392",  // [67]  Etapa 26
            "BN392",  // [68]  Etapa 27
            "BN392",  // [69]  Etapa 28
            "BN392",  // [70]  Etapa 29
            "BN392"   // [71]  Etapa 30   
        };
        private readonly string[] PFUI2018b = new string[]
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
            "AR112",   // [21]  Item 17.1
            "AR114",   // [22]  Item 17.2
            "AR126",   // [23]  Item 17.3
            "AR133",   // [24]  Item 17.4
            "AR142",   // [25]  Item 17.5
            "AR152",   // [26]  Item 17.6
            "AR161",   // [27]  Item 17.7
            "AR168",   // [28]  Item 17.8
            "AR175",   // [29]  Item 17.9
            "AR187",   // [30]  Item 17.10
            "AR194",   // [31]  Item 17.11
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
            "AJ309",  // [41]  Pré-Exc.
            "AL308",  // [42]  Etapa 1
            "AP308",  // [43]  Etapa 2
            "AT308",  // [44]  Etapa 3
            "AX308",  // [45]  Etapa 4
            "BB308",  // [46]  Etapa 5
            "BF308",  // [47]  Etapa 6
            "BJ308",  // [48]  Etapa 7
            "BN308",  // [48]  Etapa 8
            "AL348",  // [50]  Etapa 9
            "AP348",  // [51]  Etapa 10
            "AT348",  // [52]  Etapa 11
            "AX348",  // [53]  Etapa 12
            "BB348",  // [54]  Etapa 13
            "BF348",  // [55]  Etapa 14
            "BJ348",  // [56]  Etapa 15
            "BN348",  // [57]  Etapa 16
            "AL388",  // [58]  Etapa 17
            "AP388",  // [59]  Etapa 18
            "AT388",  // [60]  Etapa 19
            "AX388",  // [61]  Etapa 20
            "BB388",  // [62]  Etapa 21
            "BF388",  // [63]  Etapa 22
            "BJ388",  // [64]  Etapa 23
            "BN388",  // [65]  Etapa 24
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN388",  // [66]  Etapa 25
            "BN388",  // [67]  Etapa 26
            "BN388",  // [68]  Etapa 27
            "BN388",  // [69]  Etapa 28
            "BN388",  // [70]  Etapa 29
            "BN388"   // [71]  Etapa 30  
        };
        private readonly string[] PFUI2019 = new string[]
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
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN390",  // [66]  Etapa 25
            "BN390",  // [67]  Etapa 26
            "BN390",  // [68]  Etapa 27
            "BN390",  // [69]  Etapa 28
            "BN390",  // [70]  Etapa 29
            "BN390"   // [71]  Etapa 30  
        };
        private readonly string[] PFUI2020a = new string[]
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
                ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
                "BN391",  // [66]  Etapa 25
                "BN391",  // [67]  Etapa 26
                "BN391",  // [68]  Etapa 27
                "BN391",  // [69]  Etapa 28
                "BN391",  // [70]  Etapa 29
                "BN391"   // [71]  Etapa 30  
            };
        private readonly string[] PFUI2020b = new string[]
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
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "BN389",  // [66]  Etapa 25
            "BN389",  // [67]  Etapa 26
            "BN389",  // [68]  Etapa 27
            "BN389",  // [69]  Etapa 28
            "BN389",  // [70]  Etapa 29
            "BN389"   // [71]  Etapa 30  
        };
        private readonly string[] PFUI2020c = new string[]
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
            ///"AL430", "AP430", "AT430", "AX430", "BB430", "BF430"
            "AL430",  // [66]  Etapa 25
            "AP430",  // [67]  Etapa 26
            "AT430",  // [68]  Etapa 27
            "AX430",  // [69]  Etapa 28
            "BB430",  // [70]  Etapa 29
            "BF430"   // [71]  Etapa 30  
        };
        private readonly string[] PCI2021a = new string[]
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
        private readonly string[] PCI2021b = new string[]
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



        public string[] Get(string footer)
        {
            var onlyNumber = new Regex(@"[^\d]");
            int convertedVersion = Convert.ToInt32(onlyNumber.Replace(footer, ""));

            switch (convertedVersion)
            {
                case 27112017:
                    return PFUI2017;

                case 1102018:
                    return PFUI2018a;

                case 26022018:
                    return PFUI2018b;

                case 12042019:
                case 23052019:
                case 24052019:
                    return PFUI2019;

                case 13072020:
                case 23072020:
                    return PFUI2020a;

                case 24092020:
                case 26022021:
                case 5052021:
                    return PFUI2020b;

                case 4122020:
                    return PFUI2020c;

                case 1062021:
                case 5072021:
                case 14072021:
                case 6082021:
                    return PCI2021a;

                case 21102021:
                case 4112021:
                    return PCI2021b;

                default:
                    return null;
            }
        }


    }
}
