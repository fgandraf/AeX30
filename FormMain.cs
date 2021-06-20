using System;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;

namespace Raecef
{
    public partial class FormMain : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);




        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            TabControl.ItemSize = new System.Drawing.Size(0, 1);
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
            if (TabControl.SelectedIndex == 1)
            {
                TabControl.SelectTab(0);
                btnBack.Hide();
            }
            else if (TabControl.SelectedIndex == 2)
                TabControl.SelectTab(1);
            else if (TabControl.SelectedIndex == 3)
                TabControl.SelectTab(2);
            else if (TabControl.SelectedIndex == 4)
                TabControl.SelectTab(3);

        }







        private void btnImportarConvocacao_Click(object sender, EventArgs e)
        {
            try
            {
                if (openText.ShowDialog() == DialogResult.OK)
                {
                    var convocacao = File.ReadAllLines(openText.FileName)
                           .Where(l => l.StartsWith("REFERENCIA ........"))
                           .Select(l => l.Substring(l.LastIndexOf(":") + 2))
                           .ToList();

                    string referencia = Util.CleanInput(convocacao[0]);
                    
                    if (referencia.Substring(0, 1) == "0")
                        referencia = referencia.Substring(1);

                    txtRef0.Text = referencia.Substring(0, 4);
                    txtRef1.Text = referencia.Substring(4, 4);
                    txtRef2.Text = Convert.ToInt32(referencia.Substring(8, 9)).ToString();
                    txtRef3.Text = referencia.Substring(17, 4);
                    txtRef4.Text = referencia.Substring(21, 2);
                    txtRef5.Text = referencia.Substring(23, 2);
                    txtRef6.Text = referencia.Substring(25, 2);

                    pnlMainConvocacao.Show();
                    btnProximoConvocacao.Show();
                    txtRef0.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensagem de erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void btnImportarPfui_Click(object sender, EventArgs e)
        {
            try
            {
                if (openExcel.ShowDialog() == DialogResult.OK)
                {
                    var xls = new XLWorkbook(openExcel.FileName);
                    var planilha = xls.Worksheets.First(w => w.Name == "Proposta");

                    //CABEÇALHO
                    {
                        txtPropNome.Text = planilha.Cell("G42").RichText.ToString().ToUpper();
                        txtPropCPF.Text = Util.FormatedCPF(planilha.Cell("AA42").RichText.ToString());
                        txtPropDDD.Text = planilha.Cell("AG42").RichText.ToString();
                        txtPropTelefone.Text = Util.FormatedFone(planilha.Cell("AI42").RichText.ToString());
                        txtRTNome.Text = planilha.Cell("AN42").RichText.ToString().ToUpper();
                        txtRTCauCrea.Text = planilha.Cell("AX42").RichText.ToString().Replace(',', '.');
                        txtRTUF.Text = planilha.Cell("BD42").RichText.ToString();
                        txtRTCPF.Text = Util.FormatedCPF(planilha.Cell("BF42").Value.ToString());
                        txtRTDDD.Text = planilha.Cell("BL42").RichText.ToString();
                        txtRTTelefone.Text = Util.FormatedFone(planilha.Cell("BN42").RichText.ToString());
                        txtIdEndereco.Text = planilha.Cell("G46").RichText.ToString().ToUpper();
                        txtIdComplemento.Text = planilha.Cell("AP46").RichText.ToString().ToUpper();
                        txtIdCEP.Text = Util.FormatedCEP(planilha.Cell("BA46").RichText.ToString());
                        txtIdBairro.Text = planilha.Cell("BF46").RichText.ToString().ToUpper();
                        txtIdMunicipio.Text = planilha.Cell("G48").RichText.ToString().ToUpper();
                        txtIdUF.Text = planilha.Cell("AB48").RichText.ToString();
                        txtTerrenoValorProposto.Text = string.Format("{0:0,0.00}", Convert.ToDouble(planilha.Cell("AG50").CachedValue));
                        txtTerrenoMatricula.Text = planilha.Cell("AP50").RichText.ToString().Replace(',', '.');
                        txtTerrenoOficio.Text = planilha.Cell("AX50").RichText.ToString().ToUpper();
                        txtTerrenoComarca.Text = planilha.Cell("BF50").RichText.ToString().ToUpper();
                        txtTerrenoUF.Text = planilha.Cell("BN50").RichText.ToString();
                    }

                    //ORÇAMENTO (PERCENTUAIS)
                    {
                        txt1701.Text = planilha.Cell("AR116").CachedValue.ToString().Replace('.', ',');
                        txt1702.Text = planilha.Cell("AR118").CachedValue.ToString().Replace('.', ',');
                        txt1703.Text = planilha.Cell("AR130").CachedValue.ToString().Replace('.', ',');
                        txt1704.Text = planilha.Cell("AR137").CachedValue.ToString().Replace('.', ',');
                        txt1705.Text = planilha.Cell("AR146").CachedValue.ToString().Replace('.', ',');
                        txt1706.Text = planilha.Cell("AR156").CachedValue.ToString().Replace('.', ',');
                        txt1707.Text = planilha.Cell("AR165").CachedValue.ToString().Replace('.', ',');
                        txt1708.Text = planilha.Cell("AR172").CachedValue.ToString().Replace('.', ',');
                        txt1709.Text = planilha.Cell("AR179").CachedValue.ToString().Replace('.', ',');
                        txt1710.Text = planilha.Cell("AR190").CachedValue.ToString().Replace('.', ',');
                        txt1711.Text = planilha.Cell("AR197").CachedValue.ToString().Replace('.', ',');
                        txt1712.Text = planilha.Cell("AR207").CachedValue.ToString().Replace('.', ',');
                        txt1713.Text = planilha.Cell("AR217").CachedValue.ToString().Replace('.', ',');
                        txt1714.Text = planilha.Cell("AR228").CachedValue.ToString().Replace('.', ',');
                        txt1715.Text = planilha.Cell("AR234").CachedValue.ToString().Replace('.', ',');
                        txt1716.Text = planilha.Cell("AR245").CachedValue.ToString().Replace('.', ',');
                        txt1717.Text = planilha.Cell("AR254").CachedValue.ToString().Replace('.', ',');
                        txt1718.Text = planilha.Cell("AR262").CachedValue.ToString().Replace('.', ',');
                        txt1719.Text = planilha.Cell("AR271").CachedValue.ToString().Replace('.', ',');
                        txt1720.Text = planilha.Cell("AR273").CachedValue.ToString().Replace('.', ',');
                    }

                    //CRONOGRAMA - PRIMEIRA PÁGINA
                    {
                        string linha18_50_01;

                        if (planilha.Cell("G310").CachedValue.ToString() == "18.50.01")
                            linha18_50_01 = "310";
                        else if (planilha.Cell("G311").CachedValue.ToString() == "18.50.01")
                            linha18_50_01 = "311";
                        else
                            linha18_50_01 = "";


                        txtExecutado.Text = planilha.Cell("AJ" + (Convert.ToInt32(linha18_50_01) + 1).ToString()).CachedValue.ToString().Replace('.', ',');
                        txtParcela1.Text = planilha.Cell("AL" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                        txtParcela2.Text = planilha.Cell("AP" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                        txtParcela3.Text = planilha.Cell("AT" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                        txtParcela4.Text = planilha.Cell("AX" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                        txtParcela5.Text = planilha.Cell("BB" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                        txtParcela6.Text = planilha.Cell("BF" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                        txtParcela7.Text = planilha.Cell("BJ" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                        txtParcela8.Text = planilha.Cell("BN" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                    }

                    //CRONOGRAMA - SEGUNDA PÁGINA
                    {
                        string linha18_50_02;

                        if (planilha.Cell("G350").CachedValue.ToString() == "18.50.02")
                            linha18_50_02 = "350";
                        else if (planilha.Cell("G351").CachedValue.ToString() == "18.50.02")
                            linha18_50_02 = "351";
                        else
                            linha18_50_02 = "";

                        txtParcela9.Text = planilha.Cell("AL" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                        txtParcela10.Text = planilha.Cell("AP" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                        txtParcela11.Text = planilha.Cell("AT" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                        txtParcela12.Text = planilha.Cell("AX" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                        txtParcela13.Text = planilha.Cell("BB" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                        txtParcela14.Text = planilha.Cell("BF" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                        txtParcela15.Text = planilha.Cell("BJ" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                        txtParcela16.Text = planilha.Cell("BN" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                    }


                    //Type t = pfui.GetType();
                    //PropertyInfo[] pi = t.GetProperties();
                    //foreach (PropertyInfo p in pi)
                    //{
                    //    txtBox.Text += p.Name + ": " + p.GetValue(pfui) + "\r\n";
                    //}

                    pnlMainPfui.Show();
                    btnProximoPfui.Show();
                    txtPropNome.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensagem de erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }









        private void btnIniciar_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(1);
            btnBack.Show();
        }

        private void btnProximoConvocacao_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(2);
        }

        private void btnProximoPfui_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(3);
            txtContratoInicio.Focus();
        }

        private void btnProximoAdicionais_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(4);
        }
    }
}
