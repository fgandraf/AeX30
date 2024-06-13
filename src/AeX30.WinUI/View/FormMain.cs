using AeX30.Core.Entities;
using AeX30.Core.ValueObject;
using System;
using System.Windows.Forms;
using AeX30.Core.Interfaces;


namespace AeX30.WinUI.View
{
    public partial class FormMain : Form
    {
        private string _templatePath;
        private IRequestService _requestService;
        private IProposalService _proposalService;
        private IReportService _reportService;


        public FormMain(IRequestService requestService, IProposalService proposalService, IReportService reportService)
        {
            _requestService = requestService;
            _proposalService = proposalService;
            _reportService = reportService;
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
            => tabControl.ItemSize = new System.Drawing.Size(0, 1);

        private void NextTabControl(object sender, EventArgs e)
            => tabControl.SelectTab(tabControl.SelectedIndex + 1);

        private void BackTabControl(object sender, EventArgs e)
            => tabControl.SelectTab(tabControl.SelectedIndex - 1);

        private void btnImportarConvocacao_Click(object sender, EventArgs e)
        {
            if (openText.ShowDialog() == DialogResult.OK)
            {
                pnlMainConvocacao.Show();
                btnStartNext.Show();
                txtRef1.Focus();

                var request = _requestService.LoadFromFile(openText.FileName);

                if (request is null)
                {
                    MessageBox.Show("Não foi possível ler o arquivo de convocação.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ShowRequestOnScreen(request);
            }
        }

        private void btnImportarProposta_Click(object sender, EventArgs e)
        {
            if (openExcel.ShowDialog() == DialogResult.OK)
            {
                pnlMainPfui.Show();
                btnStartNext.Show();
                txtPropNome.Focus();

                var proposal = _proposalService.LoadFromFile(openExcel.FileName);

                if (proposal is null)
                {
                    MessageBox.Show("Arquivo incompatível ou versão não suportada!\r\n\r\n", "Planilha PCI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ShowProposalOnScreen(proposal);
            }
        }

        private void btnModeloPadrao_Click(object sender, EventArgs e)
        {
            if (openExcel.ShowDialog() == DialogResult.OK)
            {
                _templatePath = openExcel.FileName;

                txtLogFinalizar.AppendText("Caminho do modelo padrão:\r\n\r\n");
                txtLogFinalizar.AppendText(_templatePath + "\r\n\r\n");
                txtLogFinalizar.AppendText("Pronto para gravar.\r\n\r\n");
                txtLogFinalizar.AppendText("Aguardando confirmação do usuário...\r\n");

                pnlMainFinalizar.Show();
                btnSalvarComo.Show();
            }
        }

        private void btnSalvarComo_Click(object sender, EventArgs e)
        {
            var report = BuildReport();

            saveExcel.FileName = report.SuggestedFileName();
            if (saveExcel.ShowDialog() == DialogResult.OK)
            {
                var success = _reportService.SaveReport(_templatePath, saveExcel.FileName, report);

                if (!success)
                {
                    txtLogFinalizar.Text += "\r\n--------------------------------\r\n\r\nNão foi possível escrever o relatório!";
                    return;
                }

                txtLogFinalizar.Text += "\r\n--------------------------------\r\n\r\nConcluído!";
                btnNew.Show();                    
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Controls.Clear();
            InitializeComponent();
            FormMain_Load(null, EventArgs.Empty);
        }


        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(tabControl.SelectedIndex)
            {
                case 0:
                    {
                        btnStartNext.Text = "Iniciar ❯❯";
                        btnBack.Hide();
                        btnStartNext.Show();
                        break;
                    }
                case 1:
                    {
                        btnBack.Show();
                        btnStartNext.Hide();

                        if (btnStartNext.Text == "Iniciar ❯❯")
                            btnStartNext.Text = "Próximo ❯❯";
                        break;
                    }
                case 3:
                    {
                        btnStartNext.Show();
                        break;
                    }
                default:
                    {
                        btnBack.Show();
                        btnStartNext.Hide();
                        break;
                    }
            }

            if ((tabControl.SelectedIndex == 1 && pnlMainConvocacao.Visible == true) ||
                (tabControl.SelectedIndex == 2 && pnlMainPfui.Visible == true))
            {
                btnStartNext.Show();
            }

        }

        private void ShowRequestOnScreen(Request request)
        {
            txtRef1.Text = request.Reference[1];
            txtRef2.Text = request.Reference[2];
            txtRef3.Text = request.Reference[3];
            txtRef4.Text = request.Reference[4];
            txtRef5.Text = request.Reference[5];
            txtRef6.Text = request.Reference[6];
        }

        private void ShowProposalOnScreen(Proposal proposal)
        {
            lblVigencia.Text = $"Ꙩ PCI | {proposal.Vigencia}";
            lblVigencia.Show();

            txtPropNome.Text = proposal.ProponenteNome.ToUpper();
            txtPropCPF.Text = proposal.ProponenteCPF.Number;
            txtPropDDD.Text = proposal.ProponenteDDD.ToUpper();
            txtPropTelefone.Text = proposal.ProponenteFone.Number;
            txtRTNome.Text = proposal.ResponsavelNome.ToUpper();
            txtRTCauCrea.Text = proposal.ResponsavelCauCrea.ToUpper();
            txtRTUF.Text = proposal.ResponsavelUF.ToUpper();
            txtRTCPF.Text = proposal.ResponsavelCPF.Number;
            txtRTDDD.Text = proposal.ResponsavelDDD.ToUpper();
            txtRTTelefone.Text = proposal.ResponsavelFone.Number;
            txtIdEndereco.Text = proposal.ImovelEndereco.ToUpper();
            txtIdComplemento.Text = proposal.ImovelComplemento.ToUpper();
            txtIdCEP.Text = proposal.ImovelCep.Number;
            txtIdBairro.Text = proposal.ImovelBairro.ToUpper();
            txtIdMunicipio.Text = proposal.ImovelMunicipio.ToUpper();
            txtIdUF.Text = proposal.ImovelUF.ToUpper();
            txtTerrenoValorProposto.Text = proposal.ImovelValorTerreno.Number;
            txtTerrenoMatricula.Text = proposal.ImovelMatricula.ToUpper();
            txtTerrenoOficio.Text = proposal.ImovelOficio.ToUpper();
            txtTerrenoComarca.Text = proposal.ImovelComarca.ToUpper();
            txtTerrenoUF.Text = proposal.ImovelComarcaUF.ToUpper();
            txt1701.Text = proposal.ServicoItem01;
            txt1702.Text = proposal.ServicoItem02;
            txt1703.Text = proposal.ServicoItem03;
            txt1704.Text = proposal.ServicoItem04;
            txt1705.Text = proposal.ServicoItem05;
            txt1706.Text = proposal.ServicoItem06;
            txt1707.Text = proposal.ServicoItem07;
            txt1708.Text = proposal.ServicoItem08;
            txt1709.Text = proposal.ServicoItem09;
            txt1710.Text = proposal.ServicoItem10;
            txt1711.Text = proposal.ServicoItem11;
            txt1712.Text = proposal.ServicoItem12;
            txt1713.Text = proposal.ServicoItem13;
            txt1714.Text = proposal.ServicoItem14;
            txt1715.Text = proposal.ServicoItem15;
            txt1716.Text = proposal.ServicoItem16;
            txt1717.Text = proposal.ServicoItem17;
            txt1718.Text = proposal.ServicoItem18;
            txt1719.Text = proposal.ServicoItem19;
            txt1720.Text = proposal.ServicoItem20;
            txtExecutado.Text = proposal.CronogramaExecutado;
            txtParcela1.Text = proposal.CronogramaEtapa1;
            txtParcela2.Text = proposal.CronogramaEtapa2;
            txtParcela3.Text = proposal.CronogramaEtapa3;
            txtParcela4.Text = proposal.CronogramaEtapa4;
            txtParcela5.Text = proposal.CronogramaEtapa5;
            txtParcela6.Text = proposal.CronogramaEtapa6;
            txtParcela7.Text = proposal.CronogramaEtapa7;
            txtParcela8.Text = proposal.CronogramaEtapa8;
            txtParcela9.Text = proposal.CronogramaEtapa9;
            txtParcela10.Text = proposal.CronogramaEtapa10;
            txtParcela11.Text = proposal.CronogramaEtapa11;
            txtParcela12.Text = proposal.CronogramaEtapa12;
            txtParcela13.Text = proposal.CronogramaEtapa13;
            txtParcela14.Text = proposal.CronogramaEtapa14;
            txtParcela15.Text = proposal.CronogramaEtapa15;
            txtParcela16.Text = proposal.CronogramaEtapa16;
            txtParcela17.Text = proposal.CronogramaEtapa17;
            txtParcela18.Text = proposal.CronogramaEtapa18;
            txtParcela19.Text = proposal.CronogramaEtapa19;
            txtParcela20.Text = proposal.CronogramaEtapa20;
            txtParcela21.Text = proposal.CronogramaEtapa21;
            txtParcela22.Text = proposal.CronogramaEtapa22;
            txtParcela23.Text = proposal.CronogramaEtapa23;
            txtParcela24.Text = proposal.CronogramaEtapa24;
            txtParcela25.Text = proposal.CronogramaEtapa25;
            txtParcela26.Text = proposal.CronogramaEtapa26;
            txtParcela27.Text = proposal.CronogramaEtapa27;
            txtParcela28.Text = proposal.CronogramaEtapa28;
            txtParcela29.Text = proposal.CronogramaEtapa29;
            txtParcela30.Text = proposal.CronogramaEtapa30;
        }

        private Report BuildReport()
        {
            var refer = new string[7];
            refer[1] = txtRef1.Text;
            refer[2] = txtRef2.Text;
            refer[3] = txtRef3.Text;
            refer[4] = txtRef4.Text;
            refer[5] = txtRef5.Text;
            refer[6] = txtRef6.Text;
            var request = new Request(refer);

            var proposal = new Proposal(
                tipo: string.Empty,
                vigencia: string.Empty,
                proponenteNome: txtPropNome.Text,
                proponenteCPF: new Cpf(txtPropCPF.Text),
                proponenteDDD: txtPropDDD.Text,
                proponenteFone: new PhoneNumber(txtPropTelefone.Text),
                responsavelNome: txtRTNome.Text,
                responsavelCauCrea: txtRTCauCrea.Text,
                responsavelUF: txtRTUF.Text,
                responsavelCPF: new Cpf(txtRTCPF.Text),
                responsavelDDD: txtRTDDD.Text,
                responsavelFone: new PhoneNumber(txtRTTelefone.Text),
                imovelEndereco: txtIdEndereco.Text,
                imovelComplemento: txtIdComplemento.Text,
                imovelCep: new ZipCode(txtIdCEP.Text),
                imovelBairro: txtIdBairro.Text,
                imovelMunicipio: txtIdMunicipio.Text,
                imovelUF: txtIdUF.Text,
                imovelValorTerreno: new Money(txtTerrenoValorProposto.Text),
                imovelMatricula: txtTerrenoMatricula.Text,
                imovelOficio: txtTerrenoOficio.Text,
                imovelComarca: txtTerrenoComarca.Text,
                imovelComarcaUF: txtTerrenoUF.Text,
                servicoItem01: txt1701.Text,
                servicoItem02: txt1702.Text,
                servicoItem03: txt1703.Text,
                servicoItem04: txt1704.Text,
                servicoItem05: txt1705.Text,
                servicoItem06: txt1706.Text,
                servicoItem07: txt1707.Text,
                servicoItem08: txt1708.Text,
                servicoItem09: txt1709.Text,
                servicoItem10: txt1710.Text,
                servicoItem11: txt1711.Text,
                servicoItem12: txt1712.Text,
                servicoItem13: txt1713.Text,
                servicoItem14: txt1714.Text,
                servicoItem15: txt1715.Text,
                servicoItem16: txt1716.Text,
                servicoItem17: txt1717.Text,
                servicoItem18: txt1718.Text,
                servicoItem19: txt1719.Text,
                servicoItem20: txt1720.Text,
                cronogramaExecutado: txtExecutado.Text,
                cronogramaEtapa1: txtParcela1.Text,
                cronogramaEtapa2: txtParcela2.Text,
                cronogramaEtapa3: txtParcela3.Text,
                cronogramaEtapa4: txtParcela4.Text,
                cronogramaEtapa5: txtParcela5.Text,
                cronogramaEtapa6: txtParcela6.Text,
                cronogramaEtapa7: txtParcela7.Text,
                cronogramaEtapa8: txtParcela8.Text,
                cronogramaEtapa9: txtParcela9.Text,
                cronogramaEtapa10: txtParcela10.Text,
                cronogramaEtapa11: txtParcela11.Text,
                cronogramaEtapa12: txtParcela12.Text,
                cronogramaEtapa13: txtParcela13.Text,
                cronogramaEtapa14: txtParcela14.Text,
                cronogramaEtapa15: txtParcela15.Text,
                cronogramaEtapa16: txtParcela16.Text,
                cronogramaEtapa17: txtParcela17.Text,
                cronogramaEtapa18: txtParcela18.Text,
                cronogramaEtapa19: txtParcela19.Text,
                cronogramaEtapa20: txtParcela20.Text,
                cronogramaEtapa21: txtParcela21.Text,
                cronogramaEtapa22: txtParcela22.Text,
                cronogramaEtapa23: txtParcela23.Text,
                cronogramaEtapa24: txtParcela24.Text,
                cronogramaEtapa25: txtParcela25.Text,
                cronogramaEtapa26: txtParcela26.Text,
                cronogramaEtapa27: txtParcela27.Text,
                cronogramaEtapa28: txtParcela28.Text,
                cronogramaEtapa29: txtParcela29.Text,
                cronogramaEtapa30: txtParcela30.Text
                );

            var report = new Report(
                request,
                proposal,
                txtMensuradoAcumulado.Text,
                txtContratoInicio.Text,
                txtContratoTermino.Text
                );

            return report;
        }
    }
}
