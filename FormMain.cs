using System;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace Raecef
{
    public partial class FormMain : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);



        //Linha 1: Cabeçalho | Linha 2: Orçamento | Linha 3: Cronograma-1 | Linha 4: Cronograma-2 
        string[] ae130v018 = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42", "G46", "AP46", "BA46", "BF46", "G48", "AB48", "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR116", "AR118", "AR130", "AR137", "AR146", "AR156", "AR165", "AR172", "AR179", "AR190", "AR197", "AR207", "AR217", "AR228", "AR234", "AR245", "AR254", "AR262", "AR271", "AR273",
            "AJ311", "AL310", "AP310", "AT310", "AX310", "BB310", "BF310", "BJ310", "BN310",
            "AL350", "AP350", "AT350", "AX350", "BB350", "BF350", "BJ350", "BN350"
        };

        string[] ae130v020 = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42", "G46", "AP46", "BA46", "BF46", "G48", "AB48", "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR116", "AR118", "AR130", "AR137", "AR146", "AR156", "AR165", "AR172", "AR179", "AR190", "AR197", "AR207", "AR217", "AR228", "AR234", "AR245", "AR254", "AR262", "AR271", "AR273",
            "AJ312", "AL311", "AP311", "AT311", "AX311", "BB311", "BF311", "BJ311", "BN311",
            "AL351", "AP351", "AT351", "AX351", "BB351", "BF351", "BJ351", "BN351"
        };



        private void DateMask_Enter(object sender, EventArgs e)
        {
            if (((MaskedTextBox)sender).Text == "")
                ((MaskedTextBox)sender).Mask = "00/00/0000";
        }

        private void DateMask_Leave(object sender, EventArgs e)
        {
            if (((MaskedTextBox)sender).Text == "  /  /")
                ((MaskedTextBox)sender).Mask = "";
        }







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
                    string[] aeX;


                    //CHECK VERSION & CREATE RELATED ARRAY
                    {
                        if (planilha.Cell("G310").CachedValue.ToString() == "18.50.01")
                            aeX = ae130v018;
                        else if (planilha.Cell("G311").CachedValue.ToString() == "18.50.01")
                            aeX = ae130v020;
                        else
                            aeX = null;
                    }


                    //POPULATE TEXTBOX
                    {
                        //CABEÇALHO
                        txtPropNome.Text = planilha.Cell(aeX[0]).RichText.ToString().ToUpper();
                        txtPropCPF.Text = Util.FormatedCPF(planilha.Cell(aeX[1]).RichText.ToString());
                        txtPropDDD.Text = planilha.Cell(aeX[2]).RichText.ToString();
                        txtPropTelefone.Text = Util.FormatedFone(planilha.Cell(aeX[3]).RichText.ToString());
                        txtRTNome.Text = planilha.Cell(aeX[4]).RichText.ToString().ToUpper();
                        txtRTCauCrea.Text = planilha.Cell(aeX[5]).RichText.ToString().Replace(',', '.');
                        txtRTUF.Text = planilha.Cell(aeX[6]).RichText.ToString();
                        txtRTCPF.Text = Util.FormatedCPF(planilha.Cell(aeX[7]).Value.ToString());
                        txtRTDDD.Text = planilha.Cell(aeX[8]).RichText.ToString();
                        txtRTTelefone.Text = Util.FormatedFone(planilha.Cell(aeX[9]).RichText.ToString());
                        txtIdEndereco.Text = planilha.Cell(aeX[10]).RichText.ToString().ToUpper();
                        txtIdComplemento.Text = planilha.Cell(aeX[11]).RichText.ToString().ToUpper();
                        txtIdCEP.Text = Util.FormatedCEP(planilha.Cell(aeX[12]).RichText.ToString());
                        txtIdBairro.Text = planilha.Cell(aeX[13]).RichText.ToString().ToUpper();
                        txtIdMunicipio.Text = planilha.Cell(aeX[14]).RichText.ToString().ToUpper();
                        txtIdUF.Text = planilha.Cell(aeX[15]).RichText.ToString();
                        txtTerrenoValorProposto.Text = string.Format("{0:0,0.00}", Convert.ToDouble(planilha.Cell(aeX[16]).CachedValue));
                        txtTerrenoMatricula.Text = planilha.Cell(aeX[17]).RichText.ToString().Replace(',', '.');
                        txtTerrenoOficio.Text = planilha.Cell(aeX[18]).RichText.ToString().ToUpper();
                        txtTerrenoComarca.Text = planilha.Cell(aeX[19]).RichText.ToString().ToUpper();
                        txtTerrenoUF.Text = planilha.Cell(aeX[20]).RichText.ToString();
                        //ORÇAMENTO (PERCENTUAIS)
                        txt1701.Text = planilha.Cell(aeX[21]).CachedValue.ToString().Replace('.', ',');
                        txt1702.Text = planilha.Cell(aeX[22]).CachedValue.ToString().Replace('.', ',');
                        txt1703.Text = planilha.Cell(aeX[23]).CachedValue.ToString().Replace('.', ',');
                        txt1704.Text = planilha.Cell(aeX[24]).CachedValue.ToString().Replace('.', ',');
                        txt1705.Text = planilha.Cell(aeX[25]).CachedValue.ToString().Replace('.', ',');
                        txt1706.Text = planilha.Cell(aeX[26]).CachedValue.ToString().Replace('.', ',');
                        txt1707.Text = planilha.Cell(aeX[27]).CachedValue.ToString().Replace('.', ',');
                        txt1708.Text = planilha.Cell(aeX[28]).CachedValue.ToString().Replace('.', ',');
                        txt1709.Text = planilha.Cell(aeX[29]).CachedValue.ToString().Replace('.', ',');
                        txt1710.Text = planilha.Cell(aeX[30]).CachedValue.ToString().Replace('.', ',');
                        txt1711.Text = planilha.Cell(aeX[31]).CachedValue.ToString().Replace('.', ',');
                        txt1712.Text = planilha.Cell(aeX[32]).CachedValue.ToString().Replace('.', ',');
                        txt1713.Text = planilha.Cell(aeX[33]).CachedValue.ToString().Replace('.', ',');
                        txt1714.Text = planilha.Cell(aeX[34]).CachedValue.ToString().Replace('.', ',');
                        txt1715.Text = planilha.Cell(aeX[35]).CachedValue.ToString().Replace('.', ',');
                        txt1716.Text = planilha.Cell(aeX[36]).CachedValue.ToString().Replace('.', ',');
                        txt1717.Text = planilha.Cell(aeX[37]).CachedValue.ToString().Replace('.', ',');
                        txt1718.Text = planilha.Cell(aeX[38]).CachedValue.ToString().Replace('.', ',');
                        txt1719.Text = planilha.Cell(aeX[39]).CachedValue.ToString().Replace('.', ',');
                        txt1720.Text = planilha.Cell(aeX[40]).CachedValue.ToString().Replace('.', ',');
                        //CRONOGRAMA - PRIMEIRA PÁGINA
                        txtExecutado.Text = planilha.Cell(aeX[41]).CachedValue.ToString().Replace('.', ',');
                        txtParcela1.Text = planilha.Cell(aeX[42]).CachedValue.ToString().Replace('.', ',');
                        txtParcela2.Text = planilha.Cell(aeX[43]).CachedValue.ToString().Replace('.', ',');
                        txtParcela3.Text = planilha.Cell(aeX[44]).CachedValue.ToString().Replace('.', ',');
                        txtParcela4.Text = planilha.Cell(aeX[45]).CachedValue.ToString().Replace('.', ',');
                        txtParcela5.Text = planilha.Cell(aeX[46]).CachedValue.ToString().Replace('.', ',');
                        txtParcela6.Text = planilha.Cell(aeX[47]).CachedValue.ToString().Replace('.', ',');
                        txtParcela7.Text = planilha.Cell(aeX[48]).CachedValue.ToString().Replace('.', ',');
                        txtParcela8.Text = planilha.Cell(aeX[49]).CachedValue.ToString().Replace('.', ',');
                        //CRONOGRAMA - SEGUNDA PÁGINA
                        txtParcela9.Text = planilha.Cell(aeX[50]).CachedValue.ToString().Replace('.', ',');
                        txtParcela10.Text = planilha.Cell(aeX[51]).CachedValue.ToString().Replace('.', ',');
                        txtParcela11.Text = planilha.Cell(aeX[52]).CachedValue.ToString().Replace('.', ',');
                        txtParcela12.Text = planilha.Cell(aeX[53]).CachedValue.ToString().Replace('.', ',');
                        txtParcela13.Text = planilha.Cell(aeX[54]).CachedValue.ToString().Replace('.', ',');
                        txtParcela14.Text = planilha.Cell(aeX[55]).CachedValue.ToString().Replace('.', ',');
                        txtParcela15.Text = planilha.Cell(aeX[56]).CachedValue.ToString().Replace('.', ',');
                        txtParcela16.Text = planilha.Cell(aeX[57]).CachedValue.ToString().Replace('.', ',');
                    }


                    //IGNORED
                    {
                        //Função exibe no textBox todos os parâmetros(valores) de um objeto
                        //Type t = pfui.GetType();
                        //PropertyInfo[] pi = t.GetProperties();
                        //foreach (PropertyInfo p in pi)
                        //{
                        //    txtBox.Text += p.Name + ": " + p.GetValue(pfui) + "\r\n";
                        //}
                    }

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
        }

        private void btnProximoAdicionais_Click(object sender, EventArgs e)
        {
            TabControl.SelectTab(4);
        }
    }
}
