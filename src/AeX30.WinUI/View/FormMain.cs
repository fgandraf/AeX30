using System;
using System.Windows.Forms;
using AeX30.Domain.Entities;
using AeX30.App.Services;


namespace AeX30.WinUI.View
{
    public partial class FormMain : Form
    {
        private string _templatePath;

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            tabControl.ItemSize = new System.Drawing.Size(0, 1);
        }


        private void btnImportarConvocacao_Click(object sender, EventArgs e)
        {

            if (openText.ShowDialog() == DialogResult.OK)
            {
                Request request = new RequestService().GetRequestNumber(openText.FileName);

                if (request is null)
                    MessageBox.Show("Não foi possível ler o arquivo de convocação.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    PopulateFromRequest(request);

                pnlMainConvocacao.Show();
                btnStartNext.Show();
                txtRef1.Focus();
            }
        }

        private void btnImportarProposta_Click(object sender, EventArgs e)
        {
            if (openExcel.ShowDialog() == DialogResult.OK)
            {
                Proposal proposal = new ProposalService().GetProposal(openExcel.FileName);

                if (proposal is null)
                    MessageBox.Show("Arquivo incompatível ou versão não suportada!\r\n\r\n", "Planilha PFUI/PCI", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                else
                    PopulateFromProposal(proposal);

                pnlMainPfui.Show();
                btnStartNext.Show();
                txtPropNome.Focus();
            }
        }

        private void btnModeloPadrao_Click(object sender, EventArgs e)
        {
            openExcel.Title = "Abrir planilha modelo de RAE";
            if (openExcel.ShowDialog() == DialogResult.OK)
            {
                _templatePath = openExcel.FileName;

                txtLogFinalizar.AppendText("Caminho do modelo padrão:\r\n");
                txtLogFinalizar.AppendText(_templatePath+ "\r\n\r\n");
                txtLogFinalizar.AppendText("Pronto para gravar.\r\n\r\n");
                txtLogFinalizar.AppendText("Aguardando confirmação do usuário...\r\n");

                pnlMainFinalizar.Show();
                btnSalvarComo.Show();
            }
        }

        private void btnSalvarComo_Click(object sender, EventArgs e)
        {
            saveExcel.FileName = SuggestedFileName();

            if (saveExcel.ShowDialog() == DialogResult.OK)
            {
                if (ReportService.SetReport(_templatePath, saveExcel.FileName, PopulateReport()))
                {
                    txtLogFinalizar.Text += "\r\n--------------------------------\r\n\r\nConcluído!";
                    btnNew.Show();
                }
                else
                {
                    txtLogFinalizar.Text += "\r\n--------------------------------\r\n\r\nNão foi possível escrever o relatório!";
                }
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
            if (tabControl.SelectedIndex == 0)
            {
                btnStartNext.Text = "Iniciar ❯❯";
                btnBack.Hide();
                btnStartNext.Show();
            }
            else if (tabControl.SelectedIndex == 1)
            {
                btnBack.Show();
                btnStartNext.Hide();

                if (btnStartNext.Text == "Iniciar ❯❯")
                    btnStartNext.Text = "Próximo ❯❯";
            }
            else if (tabControl.SelectedIndex == 3)
            {
                btnStartNext.Show();
            }
            else
            {
                btnBack.Show();
                btnStartNext.Hide();
            }


            if ((tabControl.SelectedIndex == 1 && pnlMainConvocacao.Visible == true) ||
                (tabControl.SelectedIndex == 2 && pnlMainPfui.Visible == true))
            {
                btnStartNext.Show();
            }
                

        }




        private void NextTabControl(object sender, EventArgs e)
        {
            tabControl.SelectTab(tabControl.SelectedIndex + 1);
        }

        private void BackTabControl(object sender, EventArgs e)
        {
            tabControl.SelectTab(tabControl.SelectedIndex - 1);
        }

        private void PopulateFromRequest(Request request)
        {
            txtRef1.Text = request.Referencia[1];
            txtRef2.Text = request.Referencia[2];
            txtRef3.Text = request.Referencia[3];
            txtRef4.Text = request.Referencia[4];
            txtRef5.Text = request.Referencia[5];
            txtRef6.Text = request.Referencia[6];
        }

        private void PopulateFromProposal(Proposal proposal)
        {
            if (proposal.Tipo == "PFUI")
                lblVigencia.Text = $"Ꙩ PFUI | {proposal.Vigencia}";
            else
                lblVigencia.Text = $"Ꙩ PCI | {proposal.Vigencia}";
            lblVigencia.Show();

            txtPropNome.Text = proposal.ProponenteNome.ToUpper();
            txtPropCPF.Text = proposal.ProponenteCPF.ToUpper();
            txtPropDDD.Text = proposal.ProponenteDDD.ToUpper();
            txtPropTelefone.Text = proposal.ProponenteFone.ToUpper();
            txtRTNome.Text = proposal.ResponsavelNome.ToUpper();
            txtRTCauCrea.Text = proposal.ReponsavelCauCrea.ToUpper();
            txtRTUF.Text = proposal.ResponsavelUF.ToUpper();
            txtRTCPF.Text = proposal.ResponsavelCPF.ToUpper();
            txtRTDDD.Text = proposal.ResponsavelDDD.ToUpper();
            txtRTTelefone.Text = proposal.ResponsavelFone.ToUpper();
            txtIdEndereco.Text = proposal.ImovelEndereco.ToUpper();
            txtIdComplemento.Text = proposal.ImovelComplemento.ToUpper();
            txtIdCEP.Text = proposal.ImovelCep.ToUpper();
            txtIdBairro.Text = proposal.ImovelBairro.ToUpper();
            txtIdMunicipio.Text = proposal.ImovelMunicipio.ToUpper();
            txtIdUF.Text = proposal.ImovelUF.ToUpper();
            txtTerrenoValorProposto.Text = proposal.ImovelValorTerreno.ToUpper();
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

        private string SuggestedFileName()
        {
            string[] proponente = txtPropNome.Text.ToLower().Split(' ');
            string nome = proponente[0].Substring(0, 1).ToUpper() + proponente[0].Substring(1);
            string sobrenome = proponente[proponente.Length - 1].Substring(0, 1).ToUpper() + proponente[proponente.Length - 1].Substring(1);

            return $"RAE_{nome}-{sobrenome}.xls";
        }

        private Report PopulateReport()
        {
            Report report = new Report();

            report.Referencia[1] = txtRef1.Text;
            report.Referencia[2] = txtRef2.Text;
            report.Referencia[3] = txtRef3.Text;
            report.Referencia[4] = txtRef4.Text;
            report.Referencia[5] = txtRef5.Text;
            report.Referencia[6] = txtRef6.Text;
            report.ProponenteNome = txtPropNome.Text;
            report.ProponenteCPF = txtPropCPF.Text;
            report.ProponenteDDD = txtPropDDD.Text;
            report.ProponenteFone = txtPropTelefone.Text;
            report.ResponsavelNome = txtRTNome.Text;
            report.ReponsavelCauCrea = txtRTCauCrea.Text;
            report.ResponsavelUF = txtRTUF.Text;
            report.ResponsavelCPF = txtRTCPF.Text;
            report.ResponsavelDDD = txtRTDDD.Text;
            report.ResponsavelFone = txtRTTelefone.Text;
            report.ImovelEndereco = txtIdEndereco.Text;
            report.ImovelComplemento = txtIdComplemento.Text;
            report.ImovelBairro = txtIdBairro.Text;
            report.ImovelCep = txtIdCEP.Text;
            report.ImovelMunicipio = txtIdMunicipio.Text;
            report.ImovelUF = txtIdUF.Text;
            report.ImovelValorTerreno = txtTerrenoValorProposto.Text;
            report.ImovelMatricula = txtTerrenoMatricula.Text;
            report.ImovelOficio = txtTerrenoOficio.Text;
            report.ImovelComarca = txtTerrenoComarca.Text;
            report.ImovelComarcaUF = txtTerrenoUF.Text;
            report.ServicoItem01 = txt1701.Text;
            report.ServicoItem02 = txt1702.Text;
            report.ServicoItem03 = txt1703.Text;
            report.ServicoItem04 = txt1704.Text;
            report.ServicoItem05 = txt1705.Text;
            report.ServicoItem06 = txt1706.Text;
            report.ServicoItem07 = txt1707.Text;
            report.ServicoItem08 = txt1708.Text;
            report.ServicoItem09 = txt1709.Text;
            report.ServicoItem10 = txt1710.Text;
            report.ServicoItem11 = txt1711.Text;
            report.ServicoItem12 = txt1712.Text;
            report.ServicoItem13 = txt1713.Text;
            report.ServicoItem14 = txt1714.Text;
            report.ServicoItem15 = txt1715.Text;
            report.ServicoItem16 = txt1716.Text;
            report.ServicoItem17 = txt1717.Text;
            report.ServicoItem18 = txt1718.Text;
            report.ServicoItem19 = txt1719.Text;
            report.ServicoItem20 = txt1720.Text;
            report.MensuradoAcumulado = txtMensuradoAcumulado.Text;
            report.ContratoInicio = txtContratoInicio.Text;
            report.ContratoTermino = txtContratoTermino.Text;
            report.CronogramaExecutado = txtExecutado.Text;
            report.CronogramaEtapa1 = txtParcela1.Text;
            report.CronogramaEtapa2 = txtParcela2.Text;
            report.CronogramaEtapa3 = txtParcela3.Text;
            report.CronogramaEtapa4 = txtParcela4.Text;
            report.CronogramaEtapa5 = txtParcela5.Text;
            report.CronogramaEtapa6 = txtParcela6.Text;
            report.CronogramaEtapa7 = txtParcela7.Text;
            report.CronogramaEtapa8 = txtParcela8.Text;
            report.CronogramaEtapa9 = txtParcela9.Text;
            report.CronogramaEtapa10 = txtParcela10.Text;
            report.CronogramaEtapa11 = txtParcela11.Text;
            report.CronogramaEtapa12 = txtParcela12.Text;
            report.CronogramaEtapa13 = txtParcela13.Text;
            report.CronogramaEtapa14 = txtParcela14.Text;
            report.CronogramaEtapa15 = txtParcela15.Text;
            report.CronogramaEtapa16 = txtParcela16.Text;
            report.CronogramaEtapa17 = txtParcela17.Text;
            report.CronogramaEtapa18 = txtParcela18.Text;
            report.CronogramaEtapa19 = txtParcela19.Text;
            report.CronogramaEtapa20 = txtParcela20.Text;
            report.CronogramaEtapa21 = txtParcela21.Text;
            report.CronogramaEtapa22 = txtParcela22.Text;
            report.CronogramaEtapa23 = txtParcela23.Text;
            report.CronogramaEtapa24 = txtParcela24.Text;
            report.CronogramaEtapa25 = txtParcela25.Text;
            report.CronogramaEtapa26 = txtParcela26.Text;
            report.CronogramaEtapa27 = txtParcela27.Text;
            report.CronogramaEtapa28 = txtParcela28.Text;
            report.CronogramaEtapa29 = txtParcela29.Text;
            report.CronogramaEtapa30 = txtParcela30.Text;

            return report;
        }


    }
}
