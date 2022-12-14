using System.IO;
using System;

namespace AeX30.Domain.Entities
{
    public class Report : Proposal
    {
        public Report(string[] referencia, string mensuradoAcumulado, string contratoInicio, string contratoTermino, string tipo, string vigencia, string proponenteNome, string proponenteCPF, string proponenteDDD, string proponenteFone, string responsavelNome, string responsavelCauCrea, string responsavelUF, string responsavelCPF, string responsavelDDD, string responsavelFone, string imovelEndereco, string imovelComplemento, string imovelCep, string imovelBairro, string imovelMunicipio, string imovelUF, string imovelValorTerreno, string imovelMatricula, string imovelOficio, string imovelComarca, string imovelComarcaUF, string servicoItem01, string servicoItem02, string servicoItem03, string servicoItem04, string servicoItem05, string servicoItem06, string servicoItem07, string servicoItem08, string servicoItem09, string servicoItem10, string servicoItem11, string servicoItem12, string servicoItem13, string servicoItem14, string servicoItem15, string servicoItem16, string servicoItem17, string servicoItem18, string servicoItem19, string servicoItem20, string cronogramaExecutado, string cronogramaEtapa1, string cronogramaEtapa2, string cronogramaEtapa3, string cronogramaEtapa4, string cronogramaEtapa5, string cronogramaEtapa6, string cronogramaEtapa7, string cronogramaEtapa8, string cronogramaEtapa9, string cronogramaEtapa10, string cronogramaEtapa11, string cronogramaEtapa12, string cronogramaEtapa13, string cronogramaEtapa14, string cronogramaEtapa15, string cronogramaEtapa16, string cronogramaEtapa17, string cronogramaEtapa18, string cronogramaEtapa19, string cronogramaEtapa20, string cronogramaEtapa21, string cronogramaEtapa22, string cronogramaEtapa23, string cronogramaEtapa24, string cronogramaEtapa25, string cronogramaEtapa26, string cronogramaEtapa27, string cronogramaEtapa28, string cronogramaEtapa29, string cronogramaEtapa30) : base(tipo, vigencia, proponenteNome, proponenteCPF, proponenteDDD, proponenteFone, responsavelNome, responsavelCauCrea, responsavelUF, responsavelCPF, responsavelDDD, responsavelFone, imovelEndereco, imovelComplemento, imovelCep, imovelBairro, imovelMunicipio, imovelUF, imovelValorTerreno, imovelMatricula, imovelOficio, imovelComarca, imovelComarcaUF, servicoItem01, servicoItem02, servicoItem03, servicoItem04, servicoItem05, servicoItem06, servicoItem07, servicoItem08, servicoItem09, servicoItem10, servicoItem11, servicoItem12, servicoItem13, servicoItem14, servicoItem15, servicoItem16, servicoItem17, servicoItem18, servicoItem19, servicoItem20, cronogramaExecutado, cronogramaEtapa1, cronogramaEtapa2, cronogramaEtapa3, cronogramaEtapa4, cronogramaEtapa5, cronogramaEtapa6, cronogramaEtapa7, cronogramaEtapa8, cronogramaEtapa9, cronogramaEtapa10, cronogramaEtapa11, cronogramaEtapa12, cronogramaEtapa13, cronogramaEtapa14, cronogramaEtapa15, cronogramaEtapa16, cronogramaEtapa17, cronogramaEtapa18, cronogramaEtapa19, cronogramaEtapa20, cronogramaEtapa21, cronogramaEtapa22, cronogramaEtapa23, cronogramaEtapa24, cronogramaEtapa25, cronogramaEtapa26, cronogramaEtapa27, cronogramaEtapa28, cronogramaEtapa29, cronogramaEtapa30)
        {
            Referencia = referencia;
            MensuradoAcumulado = mensuradoAcumulado;
            ContratoInicio = contratoInicio;
            ContratoTermino = contratoTermino;
            Tipo = tipo;
            Vigencia = vigencia;
            ProponenteNome = proponenteNome;
            ProponenteCPF = FormatedCPF(proponenteCPF);
            ProponenteDDD = proponenteDDD;
            ProponenteFone = FormatedPhoneNumber(proponenteFone);
            ResponsavelNome = responsavelNome;
            ResponsavelCauCrea = responsavelCauCrea;
            ResponsavelUF = responsavelUF;
            ResponsavelCPF = FormatedCPF(responsavelCPF);
            ResponsavelDDD = responsavelDDD;
            ResponsavelFone = FormatedPhoneNumber(responsavelFone);
            ImovelEndereco = imovelEndereco;
            ImovelComplemento = imovelComplemento;
            ImovelCep = FormatedZipCode(imovelCep);
            ImovelBairro = imovelBairro;
            ImovelMunicipio = imovelMunicipio;
            ImovelUF = imovelUF;
            ImovelValorTerreno = FormatedCurrencyValue(imovelValorTerreno);
            ImovelMatricula = imovelMatricula;
            ImovelOficio = imovelOficio;
            ImovelComarca = imovelComarca;
            ImovelComarcaUF = imovelComarcaUF;
            ServicoItem01 = servicoItem01;
            ServicoItem02 = servicoItem02;
            ServicoItem03 = servicoItem03;
            ServicoItem04 = servicoItem04;
            ServicoItem05 = servicoItem05;
            ServicoItem06 = servicoItem06;
            ServicoItem07 = servicoItem07;
            ServicoItem08 = servicoItem08;
            ServicoItem09 = servicoItem09;
            ServicoItem10 = servicoItem10;
            ServicoItem11 = servicoItem11;
            ServicoItem12 = servicoItem12;
            ServicoItem13 = servicoItem13;
            ServicoItem14 = servicoItem14;
            ServicoItem15 = servicoItem15;
            ServicoItem16 = servicoItem16;
            ServicoItem17 = servicoItem17;
            ServicoItem18 = servicoItem18;
            ServicoItem19 = servicoItem19;
            ServicoItem20 = servicoItem20;
            CronogramaExecutado = cronogramaExecutado;
            CronogramaEtapa1 = cronogramaEtapa1;
            CronogramaEtapa2 = cronogramaEtapa2;
            CronogramaEtapa3 = cronogramaEtapa3;
            CronogramaEtapa4 = cronogramaEtapa4;
            CronogramaEtapa5 = cronogramaEtapa5;
            CronogramaEtapa6 = cronogramaEtapa6;
            CronogramaEtapa7 = cronogramaEtapa7;
            CronogramaEtapa8 = cronogramaEtapa8;
            CronogramaEtapa9 = cronogramaEtapa9;
            CronogramaEtapa10 = cronogramaEtapa10;
            CronogramaEtapa11 = cronogramaEtapa11;
            CronogramaEtapa12 = cronogramaEtapa12;
            CronogramaEtapa13 = cronogramaEtapa13;
            CronogramaEtapa14 = cronogramaEtapa14;
            CronogramaEtapa15 = cronogramaEtapa15;
            CronogramaEtapa16 = cronogramaEtapa16;
            CronogramaEtapa17 = cronogramaEtapa17;
            CronogramaEtapa18 = cronogramaEtapa18;
            CronogramaEtapa19 = cronogramaEtapa19;
            CronogramaEtapa20 = cronogramaEtapa20;
            CronogramaEtapa21 = cronogramaEtapa21;
            CronogramaEtapa22 = cronogramaEtapa22;
            CronogramaEtapa23 = cronogramaEtapa23;
            CronogramaEtapa24 = cronogramaEtapa24;
            CronogramaEtapa25 = cronogramaEtapa25;
            CronogramaEtapa26 = cronogramaEtapa26;
            CronogramaEtapa27 = cronogramaEtapa27;
            CronogramaEtapa28 = cronogramaEtapa28;
            CronogramaEtapa29 = cronogramaEtapa29;
            CronogramaEtapa30 = cronogramaEtapa30;
        }

        public string[] Referencia { get; private set; }
        public string MensuradoAcumulado { get; private set; }
        public string ContratoInicio { get; private set; }
        public string ContratoTermino { get; private set; }
    }
}