using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using AeX30.Controller;
using AeX30.Model.Entities;

namespace AeX30
{
    public partial class FormMain : Form
    {


        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);



        private string _caminhoModelo;


        //:METODS
        private Report PopulateReport()
        {
            Report report = new Report();

            report.Ref1 = txtRef1.Text;
            report.Ref2 = txtRef2.Text;
            report.Ref3 = txtRef3.Text;
            report.Ref4 = txtRef4.Text;
            report.Ref5 = txtRef5.Text;


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
            report.ImovelCEP = txtIdCEP.Text;
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
        private void PopulateFromProposta(Proposal prop)
        {
            if (prop.Tipo == "PFUI")
                  lblVigencia.Text = $"Ꙩ PFUI | {prop.Vigencia}";
            else
                lblVigencia.Text = $"Ꙩ PCI | {prop.Vigencia}";
            lblVigencia.Show();

            ///CABEÇALHO
            txtPropNome.Text = prop.ProponenteNome.ToUpper();
            txtPropCPF.Text = prop.ProponenteCPF.ToUpper();
            txtPropDDD.Text = prop.ProponenteDDD.ToUpper();
            txtPropTelefone.Text = prop.ProponenteFone.ToUpper();
            txtRTNome.Text = prop.ResponsavelNome.ToUpper();
            txtRTCauCrea.Text = prop.ReponsavelCauCrea.ToUpper();
            txtRTUF.Text = prop.ResponsavelUF.ToUpper();
            txtRTCPF.Text = prop.ResponsavelCPF.ToUpper();
            txtRTDDD.Text = prop.ResponsavelDDD.ToUpper();
            txtRTTelefone.Text = prop.ResponsavelFone.ToUpper();
            txtIdEndereco.Text = prop.ImovelEndereco.ToUpper();
            txtIdComplemento.Text = prop.ImovelComplemento.ToUpper();
            txtIdCEP.Text = prop.ImovelCep.ToUpper();
            txtIdBairro.Text = prop.ImovelBairro.ToUpper();
            txtIdMunicipio.Text = prop.ImovelMunicipio.ToUpper();
            txtIdUF.Text = prop.ImovelUF.ToUpper();
            txtTerrenoValorProposto.Text = prop.ImovelValorTerreno.ToUpper();
            txtTerrenoMatricula.Text = prop.ImovelMatricula.ToUpper();
            txtTerrenoOficio.Text = prop.ImovelOficio.ToUpper();
            txtTerrenoComarca.Text = prop.ImovelComarca.ToUpper();
            txtTerrenoUF.Text = prop.ImovelComarcaUF.ToUpper();
            ///ORÇAMENTO (PERCENTUAIS)
            txt1701.Text = prop.ServicoItem01;
            txt1702.Text = prop.ServicoItem02;
            txt1703.Text = prop.ServicoItem03;
            txt1704.Text = prop.ServicoItem04;
            txt1705.Text = prop.ServicoItem05;
            txt1706.Text = prop.ServicoItem06;
            txt1707.Text = prop.ServicoItem07;
            txt1708.Text = prop.ServicoItem08;
            txt1709.Text = prop.ServicoItem09;
            txt1710.Text = prop.ServicoItem10;
            txt1711.Text = prop.ServicoItem11;
            txt1712.Text = prop.ServicoItem12;
            txt1713.Text = prop.ServicoItem13;
            txt1714.Text = prop.ServicoItem14;
            txt1715.Text = prop.ServicoItem15;
            txt1716.Text = prop.ServicoItem16;
            txt1717.Text = prop.ServicoItem17;
            txt1718.Text = prop.ServicoItem18;
            txt1719.Text = prop.ServicoItem19;
            txt1720.Text = prop.ServicoItem20;
            ///CRONOGRAMA
            txtExecutado.Text = prop.CronogramaExecutado;
            txtParcela1.Text = prop.CronogramaEtapa1;
            txtParcela2.Text = prop.CronogramaEtapa2;
            txtParcela3.Text = prop.CronogramaEtapa3;
            txtParcela4.Text = prop.CronogramaEtapa4;
            txtParcela5.Text = prop.CronogramaEtapa5;
            txtParcela6.Text = prop.CronogramaEtapa6;
            txtParcela7.Text = prop.CronogramaEtapa7;
            txtParcela8.Text = prop.CronogramaEtapa8;
            txtParcela9.Text = prop.CronogramaEtapa9;
            txtParcela10.Text = prop.CronogramaEtapa10;
            txtParcela11.Text = prop.CronogramaEtapa11;
            txtParcela12.Text = prop.CronogramaEtapa12;
            txtParcela13.Text = prop.CronogramaEtapa13;
            txtParcela14.Text = prop.CronogramaEtapa14;
            txtParcela15.Text = prop.CronogramaEtapa15;
            txtParcela16.Text = prop.CronogramaEtapa16;
            txtParcela17.Text = prop.CronogramaEtapa17;
            txtParcela18.Text = prop.CronogramaEtapa18;
            txtParcela19.Text = prop.CronogramaEtapa19;
            txtParcela20.Text = prop.CronogramaEtapa20;
            txtParcela21.Text = prop.CronogramaEtapa21;
            txtParcela22.Text = prop.CronogramaEtapa22;
            txtParcela23.Text = prop.CronogramaEtapa23;
            txtParcela24.Text = prop.CronogramaEtapa24;
            txtParcela25.Text = prop.CronogramaEtapa25;
            txtParcela26.Text = prop.CronogramaEtapa29;
            txtParcela27.Text = prop.CronogramaEtapa27;
            txtParcela28.Text = prop.CronogramaEtapa28;
            txtParcela29.Text = prop.CronogramaEtapa29;
            txtParcela30.Text = prop.CronogramaEtapa30;           
        }







        //:SHARED EVENT
        private void NextTabControl(object sender, EventArgs e)
        {
            tabControl.SelectTab(tabControl.SelectedIndex + 1);
        }







        //:EVENTS
        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            tabControl.ItemSize = new System.Drawing.Size(0, 1);
        }

        private void btnAppClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void pnlAppTopPanel_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (tabControl.SelectedIndex == 1)
            {
                tabControl.SelectTab(0);
                btnBack.Hide();
            }
            else
            {
                tabControl.SelectTab(tabControl.SelectedIndex - 1);
            }

        }

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            NextTabControl(sender, e);
            btnBack.Show();
        }

        private void btnImportarConvocacao_Click(object sender, EventArgs e)
        {
            try
            {
                if (openText.ShowDialog() == DialogResult.OK)
                {
                    Request requestNumber = new RequestController().GetRequestNumber(openText.FileName);

                    //Populate
                    txtRef1.Text = requestNumber.Ref1;
                    txtRef2.Text = requestNumber.Ref2;
                    txtRef3.Text = requestNumber.Ref3;
                    txtRef4.Text = requestNumber.Ref4;
                    txtRef5.Text = requestNumber.Ref5;
                    txtRef6.Text = requestNumber.Ref6;

                    pnlMainConvocacao.Show();
                    btnProximoConvocacao.Show();
                }
            }
            catch
            {

                if (MessageBox.Show("Não foi possível ler o arquivo de convocação.\r\n \r\n Deseja inserir o número manualmente?", "Informação", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    pnlMainConvocacao.Show();
                    btnProximoConvocacao.Show();
                    txtRef1.Focus();
                }

                return;
            }
        }

        private void btnImportarProposta_Click(object sender, EventArgs e)
        {
            try
            {
                if (openExcel.ShowDialog() == DialogResult.OK)
                {

                    Proposal prop = new ProposalController().GetProposal(openExcel.FileName);
                    if (prop == null)
                    {
                        MessageBox.Show("Arquivo incompatível ou versão não suportada!\r\n\r\n", "Planilha PFUI/PCI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        PopulateFromProposta(prop);                        

                        pnlMainPfui.Show();
                        btnProximoPfui.Show();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensagem de erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void btnModeloPadrao_Click(object sender, EventArgs e)
        {
            try
            {

                if (openExcel.ShowDialog() == DialogResult.OK)
                {
                    _caminhoModelo = openExcel.FileName;

                    txtLogFinalizar.Text = "Caminho do modelo padrão:\r\n" + openExcel.FileName + "\r\n";

                    txtLogFinalizar.Text += "\r\nPronto para gravar.\r\n\r\nAguardando confirmação do usuário...\r\n";
                    pnlMainFinalizar.Show();
                    btnSalvarComo.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensagem de erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void btnSalvarComo_Click(object sender, EventArgs e)
        {
            try
            {
                //CREATING FILENAME FROM "txtPropNome" USING UPPERCASE FOR THE FIRST CHARACTERE
                string[] proponente = txtPropNome.Text.ToLower().Split(' ');
                proponente[0] = proponente[0].Substring(0, 1).ToUpper() + proponente[0].Substring(1);
                proponente[proponente.Length - 1] = proponente[proponente.Length - 1].Substring(0, 1).ToUpper() + proponente[proponente.Length - 1].Substring(1);
                saveExcel.FileName = "RAE_" + proponente[0] + "-" + proponente[proponente.Length - 1] + ".xls";


                if (saveExcel.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(_caminhoModelo))
                    {

                        if (ReportController.SetReport(_caminhoModelo, saveExcel.FileName, PopulateReport()) == 1)
                        {
                            txtLogFinalizar.Text += "\r\n--------------------------------\r\n\r\nConcluído!";
                            btnNew.Show();
                        }
                        else
                        {
                            txtLogFinalizar.Text += "Não foi possível escrever no arquivo destino!\r\n\r\nresult=0";
                            return;
                        }
                    }
                    else
                    {
                        txtLogFinalizar.Text += "Arquivo modelo não existe ou foi movido!\r\nA gravação não foi executada!";
                        return;
                    }
                }


            }
            catch (Exception ex)
            {
                txtLogFinalizar.Text += "Erro: \r\n\r\n" + ex.Message + "\r\n\r\nProcesso encerrado!";
                return;
            }
        }

        private void btnAppMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            Controls.Clear();
            InitializeComponent();
            FormMain_Load(null, EventArgs.Empty);
        }

    }
}
