using System;
using System.Text.RegularExpressions;



namespace AeX30.Domain.Entities
{
    public class Proposal
    {

        public string Vigencia { get; set; }
        public string ProponenteNome { get; set; }
        public string ProponenteDDD { get; set; }
        public string ResponsavelNome { get; set; }
        public string ReponsavelCauCrea { get; set; }
        public string ResponsavelUF { get; set; }
        public string ResponsavelDDD { get; set; }
        public string ImovelEndereco { get; set; }
        public string ImovelComplemento { get; set; }
        public string ImovelBairro { get; set; }
        public string ImovelMunicipio { get; set; }
        public string ImovelUF { get; set; }
        public string ImovelMatricula { get; set; }
        public string ImovelOficio { get; set; }
        public string ImovelComarca { get; set; }
        public string ImovelComarcaUF { get; set; }
        public string ServicoItem01 { get; set; }
        public string ServicoItem02 { get; set; }
        public string ServicoItem03 { get; set; }
        public string ServicoItem04 { get; set; }
        public string ServicoItem05 { get; set; }
        public string ServicoItem06 { get; set; }
        public string ServicoItem07 { get; set; }
        public string ServicoItem08 { get; set; }
        public string ServicoItem09 { get; set; }
        public string ServicoItem10 { get; set; }
        public string ServicoItem11 { get; set; }
        public string ServicoItem12 { get; set; }
        public string ServicoItem13 { get; set; }
        public string ServicoItem14 { get; set; }
        public string ServicoItem15 { get; set; }
        public string ServicoItem16 { get; set; }
        public string ServicoItem17 { get; set; }
        public string ServicoItem18 { get; set; }
        public string ServicoItem19 { get; set; }
        public string ServicoItem20 { get; set; }
        public string CronogramaExecutado { get; set; }
        public string CronogramaEtapa1 { get; set; }
        public string CronogramaEtapa2 { get; set; }
        public string CronogramaEtapa3 { get; set; }
        public string CronogramaEtapa4 { get; set; }
        public string CronogramaEtapa5 { get; set; }
        public string CronogramaEtapa6 { get; set; }
        public string CronogramaEtapa7 { get; set; }
        public string CronogramaEtapa8 { get; set; }
        public string CronogramaEtapa9 { get; set; }
        public string CronogramaEtapa10 { get; set; }
        public string CronogramaEtapa11 { get; set; }
        public string CronogramaEtapa12 { get; set; }
        public string CronogramaEtapa13 { get; set; }
        public string CronogramaEtapa14 { get; set; }
        public string CronogramaEtapa15 { get; set; }
        public string CronogramaEtapa16 { get; set; }
        public string CronogramaEtapa17 { get; set; }
        public string CronogramaEtapa18 { get; set; }
        public string CronogramaEtapa19 { get; set; }
        public string CronogramaEtapa20 { get; set; }
        public string CronogramaEtapa21 { get; set; }
        public string CronogramaEtapa22 { get; set; }
        public string CronogramaEtapa23 { get; set; }
        public string CronogramaEtapa24 { get; set; }
        public string CronogramaEtapa25 { get; set; }
        public string CronogramaEtapa26 { get; set; }
        public string CronogramaEtapa27 { get; set; }
        public string CronogramaEtapa28 { get; set; }
        public string CronogramaEtapa29 { get; set; }
        public string CronogramaEtapa30 { get; set; }


        private string _tipo;
        public string Tipo
        {
            get { return _tipo; }
            set { _tipo = value == "Proposta" ? "PFUI" : "PCI"; }
        }


        private string _proponenteCPF;
        public string ProponenteCPF
        {
            get { return _proponenteCPF; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    value = value.Split(',')[0];
                    value = new Regex(@"[^\d]").Replace(value, "");
                    long number = Convert.ToInt64(value);

                    _proponenteCPF = number.ToString(@"000\.000\.000\-00");
                }
            }
        }


        private string _proponenteFone;
        public string ProponenteFone
        {
            get { return _proponenteFone; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    value = new Regex(@"[^\d]").Replace(value, "");
                    long phoneNumber = Convert.ToInt64(value);

                    if (phoneNumber.ToString().Length == 8)
                        _proponenteFone = phoneNumber.ToString(@"0000\-0000");
                    else if (phoneNumber.ToString().Length == 9)
                        _proponenteFone = phoneNumber.ToString(@"00000\-0000");
                    else
                        _proponenteFone = value;
                }
            }
        }


        private string _responsavelCPF;
        public string ResponsavelCPF
        {
            get { return _responsavelCPF; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    value = value.Split(',')[0];
                    value = new Regex(@"[^\d]").Replace(value, "");
                    long number = Convert.ToInt64(value);

                    _responsavelCPF = number.ToString(@"000\.000\.000\-00");
                }
            }
        }


        private string _responsavelFone;
        public string ResponsavelFone
        {
            get { return _responsavelFone; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    value = new Regex(@"[^\d]").Replace(value, "");
                    long phoneNumber = Convert.ToInt64(value);

                    if (phoneNumber.ToString().Length == 8)
                        _responsavelFone = phoneNumber.ToString(@"0000\-0000");
                    else if (phoneNumber.ToString().Length == 9)
                        _responsavelFone = phoneNumber.ToString(@"00000\-0000");
                    else
                        _responsavelFone = value;
                }
            }
        }


        private string _imovelCep;
        public string ImovelCep
        {
            get { return _imovelCep; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    value = new Regex(@"[^\d]").Replace(value, "");
                    long cepNumber = Convert.ToInt64(value);

                    _imovelCep = cepNumber.ToString(@"00000\-000");
                }
            }
        }


        private string _valorTerreno;
        public string ImovelValorTerreno
        {
            get { return _valorTerreno; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    value = new Regex(@"[^\d]").Replace(value, "");
                    long number = Convert.ToInt64(value);

                    _valorTerreno = $"{number:N2}";
                }

            }
        }


    }
}
