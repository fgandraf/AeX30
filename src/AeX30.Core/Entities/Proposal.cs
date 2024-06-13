using AeX30.Core.ValueObject;



namespace AeX30.Core.Entities
{
    public class Proposal
    {
        public Proposal(string tipo, string vigencia, string proponenteNome, Cpf proponenteCPF, string proponenteDDD, PhoneNumber proponenteFone, string responsavelNome, string responsavelCauCrea, string responsavelUF, Cpf responsavelCPF, string responsavelDDD, PhoneNumber responsavelFone, string imovelEndereco, string imovelComplemento, ZipCode imovelCep, string imovelBairro, string imovelMunicipio, string imovelUF, Money imovelValorTerreno, string imovelMatricula, string imovelOficio, string imovelComarca, string imovelComarcaUF, string servicoItem01, string servicoItem02, string servicoItem03, string servicoItem04, string servicoItem05, string servicoItem06, string servicoItem07, string servicoItem08, string servicoItem09, string servicoItem10, string servicoItem11, string servicoItem12, string servicoItem13, string servicoItem14, string servicoItem15, string servicoItem16, string servicoItem17, string servicoItem18, string servicoItem19, string servicoItem20, string cronogramaExecutado, string cronogramaEtapa1, string cronogramaEtapa2, string cronogramaEtapa3, string cronogramaEtapa4, string cronogramaEtapa5, string cronogramaEtapa6, string cronogramaEtapa7, string cronogramaEtapa8, string cronogramaEtapa9, string cronogramaEtapa10, string cronogramaEtapa11, string cronogramaEtapa12, string cronogramaEtapa13, string cronogramaEtapa14, string cronogramaEtapa15, string cronogramaEtapa16, string cronogramaEtapa17, string cronogramaEtapa18, string cronogramaEtapa19, string cronogramaEtapa20, string cronogramaEtapa21, string cronogramaEtapa22, string cronogramaEtapa23, string cronogramaEtapa24, string cronogramaEtapa25, string cronogramaEtapa26, string cronogramaEtapa27, string cronogramaEtapa28, string cronogramaEtapa29, string cronogramaEtapa30)
        {
            Tipo = tipo == "Proposta" ? "PFUI" : "PCI";
            Vigencia = vigencia;
            ProponenteNome = proponenteNome;
            ProponenteCPF = proponenteCPF;
            ProponenteDDD = proponenteDDD;
            ProponenteFone = proponenteFone;
            ResponsavelNome = responsavelNome;
            ResponsavelCauCrea = responsavelCauCrea;
            ResponsavelUF = responsavelUF;
            ResponsavelCPF = responsavelCPF;
            ResponsavelDDD = responsavelDDD;
            ResponsavelFone = responsavelFone;
            ImovelEndereco = imovelEndereco;
            ImovelComplemento = imovelComplemento;
            ImovelCep = imovelCep;
            ImovelBairro = imovelBairro;
            ImovelMunicipio = imovelMunicipio;
            ImovelUF = imovelUF;
            ImovelValorTerreno = imovelValorTerreno;
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


        public string Tipo { get; private set; }
        public string Vigencia { get; private set; }
        public string ProponenteNome { get; private set; }
        public Cpf ProponenteCPF { get; private set; }
        public string ProponenteDDD { get; private set; }
        public PhoneNumber ProponenteFone { get; private set; }
        public string ResponsavelNome { get; private set; }
        public string ResponsavelCauCrea { get; private set; }
        public string ResponsavelUF { get; private set; }
        public Cpf ResponsavelCPF { get; private set; }
        public string ResponsavelDDD { get; private set; }
        public PhoneNumber ResponsavelFone { get; private set; }
        public string ImovelEndereco { get; private set; }
        public string ImovelComplemento { get; private set; }
        public ZipCode ImovelCep { get; private set; }
        public string ImovelBairro { get; private set; }
        public string ImovelMunicipio { get; private set; }
        public string ImovelUF { get; private set; }
        public Money ImovelValorTerreno { get; private set; }
        public string ImovelMatricula { get; private set; }
        public string ImovelOficio { get; private set; }
        public string ImovelComarca { get; private set; }
        public string ImovelComarcaUF { get; private set; }
        public string ServicoItem01 { get; private set; }
        public string ServicoItem02 { get; private set; }
        public string ServicoItem03 { get; private set; }
        public string ServicoItem04 { get; private set; }
        public string ServicoItem05 { get; private set; }
        public string ServicoItem06 { get; private set; }
        public string ServicoItem07 { get; private set; }
        public string ServicoItem08 { get; private set; }
        public string ServicoItem09 { get; private set; }
        public string ServicoItem10 { get; private set; }
        public string ServicoItem11 { get; private set; }
        public string ServicoItem12 { get; private set; }
        public string ServicoItem13 { get; private set; }
        public string ServicoItem14 { get; private set; }
        public string ServicoItem15 { get; private set; }
        public string ServicoItem16 { get; private set; }
        public string ServicoItem17 { get; private set; }
        public string ServicoItem18 { get; private set; }
        public string ServicoItem19 { get; private set; }
        public string ServicoItem20 { get; private set; }
        public string CronogramaExecutado { get; private set; }
        public string CronogramaEtapa1 { get; private set; }
        public string CronogramaEtapa2 { get; private set; }
        public string CronogramaEtapa3 { get; private set; }
        public string CronogramaEtapa4 { get; private set; }
        public string CronogramaEtapa5 { get; private set; }
        public string CronogramaEtapa6 { get; private set; }
        public string CronogramaEtapa7 { get; private set; }
        public string CronogramaEtapa8 { get; private set; }
        public string CronogramaEtapa9 { get; private set; }
        public string CronogramaEtapa10 { get; private set; }
        public string CronogramaEtapa11 { get; private set; }
        public string CronogramaEtapa12 { get; private set; }
        public string CronogramaEtapa13 { get; private set; }
        public string CronogramaEtapa14 { get; private set; }
        public string CronogramaEtapa15 { get; private set; }
        public string CronogramaEtapa16 { get; private set; }
        public string CronogramaEtapa17 { get; private set; }
        public string CronogramaEtapa18 { get; private set; }
        public string CronogramaEtapa19 { get; private set; }
        public string CronogramaEtapa20 { get; private set; }
        public string CronogramaEtapa21 { get; private set; }
        public string CronogramaEtapa22 { get; private set; }
        public string CronogramaEtapa23 { get; private set; }
        public string CronogramaEtapa24 { get; private set; }
        public string CronogramaEtapa25 { get; private set; }
        public string CronogramaEtapa26 { get; private set; }
        public string CronogramaEtapa27 { get; private set; }
        public string CronogramaEtapa28 { get; private set; }
        public string CronogramaEtapa29 { get; private set; }
        public string CronogramaEtapa30 { get; private set; }

    }
}
