using System;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Raecef.Ent;
using Raecef.Model;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Reflection;

namespace Raecef
{
    public partial class FormMain : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);

        PFUI pfui;


        public FormMain()
        {
            InitializeComponent();
        }


        private void BtnCarregarPfui_Click(object sender, EventArgs e)
        {
            try
            {
                if (OpenDialog.ShowDialog() == DialogResult.OK)
                {
                    

                    pfui = new PfuiModel().GetPFUI(OpenDialog.FileName);

                    txtPropNome.Text = pfui.Proponente_Nome;
                    txtPropCPF.Text = pfui.Proponente_CPF;
                    txtPropDDD.Text = pfui.Proponente_DDD;
                    txtPropTelefone.Text = pfui.Proponente_Telefone;
                    txtRTNome.Text = pfui.RT_Nome;
                    txtRTCauCrea.Text = pfui.RT_CauCrea;
                    txtRTUF.Text = pfui.RT_UF;
                    txtRTCPF.Text = pfui.RT_CPF;
                    txtRTDDD.Text = pfui.RT_DDD;
                    txtRTTelefone.Text = pfui.RT_Telefone;
                    txtIdEndereco.Text = pfui.Obra_Endereco;
                    txtIdComplemento.Text = pfui.Obra_Complemento;
                    txtIdCEP.Text = pfui.Obra_CEP;
                    txtIdBairro.Text = pfui.Obra_Bairro;
                    txtIdMunicipio.Text = pfui.Obra_Municipio;
                    txtIdUF.Text = pfui.Obra_UF;
                    txtTerrenoValorProposto.Text = pfui.Terreno_Valor;
                    txtTerrenoMatricula.Text = pfui.Matricula_Numero;
                    txtTerrenoOficio.Text = pfui.Matricula_Oficio;
                    txtTerrenoComarca.Text = pfui.Matricula_Comarca;
                    txtTerrenoUF.Text = pfui.Matricula_UF;

                    txt1701.Text = pfui.Item_17_01;
                    txt1702.Text = pfui.Item_17_02;
                    txt1703.Text = pfui.Item_17_03;
                    txt1704.Text = pfui.Item_17_04;
                    txt1705.Text = pfui.Item_17_05;
                    txt1706.Text = pfui.Item_17_06;
                    txt1707.Text = pfui.Item_17_07;
                    txt1708.Text = pfui.Item_17_08;
                    txt1709.Text = pfui.Item_17_09;
                    txt1710.Text = pfui.Item_17_10;
                    txt1711.Text = pfui.Item_17_11;
                    txt1712.Text = pfui.Item_17_12;
                    txt1713.Text = pfui.Item_17_13;
                    txt1714.Text = pfui.Item_17_14;
                    txt1715.Text = pfui.Item_17_15;
                    txt1716.Text = pfui.Item_17_16;
                    txt1717.Text = pfui.Item_17_17;
                    txt1718.Text = pfui.Item_17_18;
                    txt1719.Text = pfui.Item_17_19;
                    txt1720.Text = pfui.Item_17_20;

                    txtExecutado.Text = pfui.Executado;
                    txtParcela1.Text = pfui.Parcela_1;
                    txtParcela2.Text = pfui.Parcela_2;
                    txtParcela3.Text = pfui.Parcela_3;
                    txtParcela4.Text = pfui.Parcela_4;
                    txtParcela5.Text = pfui.Parcela_5;
                    txtParcela6.Text = pfui.Parcela_6;
                    txtParcela7.Text = pfui.Parcela_7;
                    txtParcela8.Text = pfui.Parcela_8;
                    txtParcela9.Text = pfui.Parcela_9;
                    txtParcela10.Text = pfui.Parcela_10;
                    txtParcela11.Text = pfui.Parcela_11;
                    txtParcela12.Text = pfui.Parcela_12;
                    txtParcela13.Text = pfui.Parcela_13;
                    txtParcela14.Text = pfui.Parcela_14;
                    txtParcela15.Text = pfui.Parcela_15;
                    txtParcela16.Text = pfui.Parcela_16;

                    //Type t = pfui.GetType();
                    //PropertyInfo[] pi = t.GetProperties();
                    //foreach (PropertyInfo p in pi)
                    //{
                    //    txtBox.Text += p.Name + ": " + p.GetValue(pfui) + "\r\n";
                    //}

                    pnlMainPg3.Show();
                    btnProximoTab3.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensagem de erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void panel3_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void BtnAppClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            TabControl.ItemSize = new System.Drawing.Size(0, 1);
        }


        private void btnProximoTab2_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(3);
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (TabControl.SelectedIndex == 1)
            {
                TabControl.SelectTab(0);
                button2.Hide();
            }
            else if (TabControl.SelectedIndex == 2)
                TabControl.SelectTab(1);
            else if (TabControl.SelectedIndex == 3)
                TabControl.SelectTab(2);
            else if (TabControl.SelectedIndex == 4)
                TabControl.SelectTab(3);

        }

        private void btnProximoTab2_Click_1(object sender, EventArgs e)
        {
            TabControl.SelectTab(2);
        }

        private void btnImportarConvocacao_Click(object sender, EventArgs e)
        {
            pnlMainPg2.Show();
            btnProximoTab2.Show();
        }

        private void btnProximoTab2_Click_2(object sender, EventArgs e)
        {
            TabControl.SelectTab(2);
        }

        private void btnIniciar_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(1);
            button2.Show();
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

        }
    }
}
