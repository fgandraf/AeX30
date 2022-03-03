using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using aeX30.Controller;
using aeX30.Entities;


namespace aeX30
{
    public partial class FormMain : Form
    {


        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);



        private string _caminhoModelo;


        //:METODS
        private RaeEnt PopulateToRAE()
        {
            RaeEnt rae = new RaeEnt();

            rae.Ref0 = txtRef0.Text;
            rae.Ref1 = txtRef1.Text;
            rae.Ref2 = txtRef2.Text;
            rae.Ref3 = txtRef3.Text;
            rae.Ref4 = txtRef4.Text;
            rae.Ref5 = txtRef5.Text;
            rae.Ref6 = txtRef6.Text;

            rae.Prop_Nome = txtPropNome.Text;
            rae.Prop_CPF = txtPropCPF.Text;
            rae.Prop_DDD = txtPropDDD.Text;
            rae.Prop_Telefone = txtPropTelefone.Text;

            rae.Rt_Nome = txtRTNome.Text;
            rae.Rt_CAU_CREA = txtRTCauCrea.Text;
            rae.Rt_UF = txtRTUF.Text;
            rae.Rt_CPF = txtRTCPF.Text;
            rae.Rt_DDD = txtRTDDD.Text;
            rae.Rt_Telefone = txtRTTelefone.Text;

            rae.End_Endereco = txtIdEndereco.Text;
            rae.End_Complemento = txtIdComplemento.Text;
            rae.End_Bairro = txtIdBairro.Text;
            rae.End_CEP = txtIdCEP.Text;
            rae.End_Municipio = txtIdMunicipio.Text;
            rae.End_UF = txtIdUF.Text;

            rae.Imov_Valor_Terreno = txtTerrenoValorProposto.Text;
            rae.Imov_Matricula = txtTerrenoMatricula.Text;
            rae.Imov_Oficio = txtTerrenoOficio.Text;
            rae.Imov_Comarca = txtTerrenoComarca.Text;
            rae.Imov_UF = txtTerrenoUF.Text;

            rae.Item_17_01 = txt1701.Text;
            rae.Item_17_02 = txt1702.Text;
            rae.Item_17_03 = txt1703.Text;
            rae.Item_17_04 = txt1704.Text;
            rae.Item_17_05 = txt1705.Text;
            rae.Item_17_06 = txt1706.Text;
            rae.Item_17_07 = txt1707.Text;
            rae.Item_17_08 = txt1708.Text;
            rae.Item_17_09 = txt1709.Text;
            rae.Item_17_10 = txt1710.Text;
            rae.Item_17_11 = txt1711.Text;
            rae.Item_17_12 = txt1712.Text;
            rae.Item_17_13 = txt1713.Text;
            rae.Item_17_14 = txt1714.Text;
            rae.Item_17_15 = txt1715.Text;
            rae.Item_17_16 = txt1716.Text;
            rae.Item_17_17 = txt1717.Text;
            rae.Item_17_18 = txt1718.Text;
            rae.Item_17_19 = txt1719.Text;
            rae.Item_17_20 = txt1720.Text;

            rae.MensuradoAcumulado = txtMensuradoAcumulado.Text;
            rae.ContratoInicio = txtContratoInicio.Text;
            rae.ContratoTermino = txtContratoTermino.Text;
            rae.EtapaPrevista = txtEtapaPrevista.Text;

            rae.Cron_Executado = txtExecutado.Text;
            rae.Cron_Parc_1 = txtParcela1.Text;
            rae.Cron_Parc_2 = txtParcela2.Text;
            rae.Cron_Parc_3 = txtParcela3.Text;
            rae.Cron_Parc_4 = txtParcela4.Text;
            rae.Cron_Parc_5 = txtParcela5.Text;
            rae.Cron_Parc_6 = txtParcela6.Text;
            rae.Cron_Parc_7 = txtParcela7.Text;
            rae.Cron_Parc_8 = txtParcela8.Text;
            rae.Cron_Parc_9 = txtParcela9.Text;
            rae.Cron_Parc_10 = txtParcela10.Text;
            rae.Cron_Parc_11 = txtParcela11.Text;
            rae.Cron_Parc_12 = txtParcela12.Text;
            rae.Cron_Parc_13 = txtParcela13.Text;
            rae.Cron_Parc_14 = txtParcela14.Text;
            rae.Cron_Parc_15 = txtParcela15.Text;
            rae.Cron_Parc_16 = txtParcela16.Text;
            rae.Cron_Parc_17 = txtParcela17.Text;
            rae.Cron_Parc_18 = txtParcela18.Text;
            rae.Cron_Parc_19 = txtParcela19.Text;
            rae.Cron_Parc_20 = txtParcela20.Text;
            rae.Cron_Parc_21 = txtParcela21.Text;
            rae.Cron_Parc_22 = txtParcela22.Text;
            rae.Cron_Parc_23 = txtParcela23.Text;
            rae.Cron_Parc_24 = txtParcela24.Text;
            rae.Cron_Parc_25 = txtParcela25.Text;
            rae.Cron_Parc_26 = txtParcela26.Text;
            rae.Cron_Parc_27 = txtParcela27.Text;
            rae.Cron_Parc_28 = txtParcela28.Text;
            rae.Cron_Parc_29 = txtParcela29.Text;
            rae.Cron_Parc_30 = txtParcela30.Text;

            return rae;
        }
        private void PopulateFromProposta(PropostaEnt prop)
        {
            if (prop.Tipo == "PFUI")
            {
                rbtPFUI.Show();
                rbtPFUI.Checked = true;
                rbtPCI.Hide();
            }
            else
            {
                rbtPCI.Show();
                rbtPCI.Checked = true;
                rbtPFUI.Hide();
            }
            lblVigencia.Show();
            lblVigencia.Text = prop.Vigencia;
            ///CABEÇALHO
            txtPropNome.Text = prop.Prop_Nome;
            txtPropCPF.Text = prop.Prop_CPF;
            txtPropDDD.Text = prop.Prop_DDD;
            txtPropTelefone.Text = prop.Prop_Telefone;
            txtRTNome.Text = prop.Rt_Nome;
            txtRTCauCrea.Text = prop.Rt_CAU_CREA;
            txtRTUF.Text = prop.Rt_UF;
            txtRTCPF.Text = prop.Rt_CPF;
            txtRTDDD.Text = prop.Rt_DDD;
            txtRTTelefone.Text = prop.Rt_Telefone;
            txtIdEndereco.Text = prop.End_Endereço;
            txtIdComplemento.Text = prop.End_Complemento;
            txtIdCEP.Text = prop.End_CEP;
            txtIdBairro.Text = prop.End_Bairro;
            txtIdMunicipio.Text = prop.End_Municipio;
            txtIdUF.Text = prop.End_UF;
            txtTerrenoValorProposto.Text = prop.Imov_Valor_Terreno;
            txtTerrenoMatricula.Text = prop.Imov_Matricula;
            txtTerrenoOficio.Text = prop.Imov_Oficio;
            txtTerrenoComarca.Text = prop.Imov_Comarca;
            txtTerrenoUF.Text = prop.Imov_UF;
            ///ORÇAMENTO (PERCENTUAIS)
            txt1701.Text = prop.Item_17_01;
            txt1702.Text = prop.Item_17_02;
            txt1703.Text = prop.Item_17_03;
            txt1704.Text = prop.Item_17_04;
            txt1705.Text = prop.Item_17_05;
            txt1706.Text = prop.Item_17_06;
            txt1707.Text = prop.Item_17_07;
            txt1708.Text = prop.Item_17_08;
            txt1709.Text = prop.Item_17_09;
            txt1710.Text = prop.Item_17_10;
            txt1711.Text = prop.Item_17_11;
            txt1712.Text = prop.Item_17_12;
            txt1713.Text = prop.Item_17_13;
            txt1714.Text = prop.Item_17_14;
            txt1715.Text = prop.Item_17_15;
            txt1716.Text = prop.Item_17_16;
            txt1717.Text = prop.Item_17_17;
            txt1718.Text = prop.Item_17_18;
            txt1719.Text = prop.Item_17_19;
            txt1720.Text = prop.Item_17_20;
            ///CRONOGRAMA
            txtExecutado.Text = prop.Cron_Executado;
            txtParcela1.Text = prop.Cron_Parc_1;
            txtParcela2.Text = prop.Cron_Parc_2;
            txtParcela3.Text = prop.Cron_Parc_3;
            txtParcela4.Text = prop.Cron_Parc_4;
            txtParcela5.Text = prop.Cron_Parc_5;
            txtParcela6.Text = prop.Cron_Parc_6;
            txtParcela7.Text = prop.Cron_Parc_7;
            txtParcela8.Text = prop.Cron_Parc_8;
            txtParcela9.Text = prop.Cron_Parc_9;
            txtParcela10.Text = prop.Cron_Parc_10;
            txtParcela11.Text = prop.Cron_Parc_11;
            txtParcela12.Text = prop.Cron_Parc_12;
            txtParcela13.Text = prop.Cron_Parc_13;
            txtParcela14.Text = prop.Cron_Parc_14;
            txtParcela15.Text = prop.Cron_Parc_15;
            txtParcela16.Text = prop.Cron_Parc_16;
            txtParcela17.Text = prop.Cron_Parc_17;
            txtParcela18.Text = prop.Cron_Parc_18;
            txtParcela19.Text = prop.Cron_Parc_19;
            txtParcela20.Text = prop.Cron_Parc_20;
            txtParcela21.Text = prop.Cron_Parc_21;
            txtParcela22.Text = prop.Cron_Parc_22;
            txtParcela23.Text = prop.Cron_Parc_23;
            txtParcela24.Text = prop.Cron_Parc_24;
            txtParcela25.Text = prop.Cron_Parc_25;
            txtParcela26.Text = prop.Cron_Parc_29;
            txtParcela27.Text = prop.Cron_Parc_27;
            txtParcela28.Text = prop.Cron_Parc_28;
            txtParcela29.Text = prop.Cron_Parc_29;
            txtParcela30.Text = prop.Cron_Parc_30;
        }







        //:SHARED EVENT
        private void NextTabControl(object sender, EventArgs e)
        {
            if (rbtPCI.Checked)
                txtEtapaPrevista.Enabled = false;
            else
                txtEtapaPrevista.Enabled = true;
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
                    string[] referencia = new ConvocacaoController().GetReferencia(openText.FileName);

                    //Populate
                    txtRef0.Text = referencia[0];
                    txtRef1.Text = referencia[1];
                    txtRef2.Text = referencia[2];
                    txtRef3.Text = referencia[3];
                    txtRef4.Text = referencia[4];
                    txtRef5.Text = referencia[5];
                    txtRef6.Text = referencia[6];

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
                    txtRef0.Focus();
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

                    PropostaEnt prop = new PropostaController().GetProposta(openExcel.FileName);
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

                        if (RaeController.SetRAE(_caminhoModelo, saveExcel.FileName, PopulateToRAE()) == 1)
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
