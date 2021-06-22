using System;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Reflection;

namespace Raecef
{
    public partial class FormMain : Form
    {
        
        
        
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);




        //------------------PFUI------------------//
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

        private string GetVersionPFUI(string header)
        {
            if (header != "")
            {
                int version = Convert.ToInt32(Util.CleanInput(header));

                if (version == 1413010018)
                    return "AE 130 018";
                else if (version == 1413010020)
                    return "AE 130 020";
                else if (version > 1413010020)
                    return "> AE 130 020";
                else
                    return null;
            }
            else
                return null;
        }
        private string[] SetCellsPFUI(string version)
        {
            lblVersion.Text = version;
            lblVersionTitle.Show();
            lblVersion.Show();

            if (version == "AE 130 018")
                return ae130v018;
            else if (version == "AE 130 020")
                return ae130v020;
            else
            {
                MessageBox.Show("A versão da planilha PFUI inserida não foi testada.\r\nRedobre a atenção quanto aos valores importados!", "Planilha PFUI não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return ae130v020;
            }
        }
        private void PopulateFromPFUI(string[] cell, ClosedXML.Excel.IXLWorksheet pfui)
        {
            //CABEÇALHO
            txtPropNome.Text = pfui.Cell(cell[0]).RichText.ToString().ToUpper();
            txtPropCPF.Text = Util.FormatedCPF(pfui.Cell(cell[1]).RichText.ToString());
            txtPropDDD.Text = pfui.Cell(cell[2]).RichText.ToString();
            txtPropTelefone.Text = Util.FormatedFone(pfui.Cell(cell[3]).RichText.ToString());
            txtRTNome.Text = pfui.Cell(cell[4]).RichText.ToString().ToUpper();
            txtRTCauCrea.Text = pfui.Cell(cell[5]).RichText.ToString().Replace(',', '.');
            txtRTUF.Text = pfui.Cell(cell[6]).RichText.ToString();
            txtRTCPF.Text = Util.FormatedCPF(pfui.Cell(cell[7]).Value.ToString());
            txtRTDDD.Text = pfui.Cell(cell[8]).RichText.ToString();
            txtRTTelefone.Text = Util.FormatedFone(pfui.Cell(cell[9]).RichText.ToString());
            txtIdEndereco.Text = pfui.Cell(cell[10]).RichText.ToString().ToUpper();
            txtIdComplemento.Text = pfui.Cell(cell[11]).RichText.ToString().ToUpper();
            txtIdCEP.Text = Util.FormatedCEP(pfui.Cell(cell[12]).RichText.ToString());
            txtIdBairro.Text = pfui.Cell(cell[13]).RichText.ToString().ToUpper();
            txtIdMunicipio.Text = pfui.Cell(cell[14]).RichText.ToString().ToUpper();
            txtIdUF.Text = pfui.Cell(cell[15]).RichText.ToString();
            txtTerrenoValorProposto.Text = string.Format("{0:0,0.00}", Convert.ToDouble(pfui.Cell(cell[16]).CachedValue));
            txtTerrenoMatricula.Text = pfui.Cell(cell[17]).RichText.ToString().Replace(',', '.');
            txtTerrenoOficio.Text = pfui.Cell(cell[18]).RichText.ToString().ToUpper();
            txtTerrenoComarca.Text = pfui.Cell(cell[19]).RichText.ToString().ToUpper();
            txtTerrenoUF.Text = pfui.Cell(cell[20]).RichText.ToString();
            //ORÇAMENTO (PERCENTUAIS)
            txt1701.Text = pfui.Cell(cell[21]).CachedValue.ToString().Replace('.', ',');
            txt1702.Text = pfui.Cell(cell[22]).CachedValue.ToString().Replace('.', ',');
            txt1703.Text = pfui.Cell(cell[23]).CachedValue.ToString().Replace('.', ',');
            txt1704.Text = pfui.Cell(cell[24]).CachedValue.ToString().Replace('.', ',');
            txt1705.Text = pfui.Cell(cell[25]).CachedValue.ToString().Replace('.', ',');
            txt1706.Text = pfui.Cell(cell[26]).CachedValue.ToString().Replace('.', ',');
            txt1707.Text = pfui.Cell(cell[27]).CachedValue.ToString().Replace('.', ',');
            txt1708.Text = pfui.Cell(cell[28]).CachedValue.ToString().Replace('.', ',');
            txt1709.Text = pfui.Cell(cell[29]).CachedValue.ToString().Replace('.', ',');
            txt1710.Text = pfui.Cell(cell[30]).CachedValue.ToString().Replace('.', ',');
            txt1711.Text = pfui.Cell(cell[31]).CachedValue.ToString().Replace('.', ',');
            txt1712.Text = pfui.Cell(cell[32]).CachedValue.ToString().Replace('.', ',');
            txt1713.Text = pfui.Cell(cell[33]).CachedValue.ToString().Replace('.', ',');
            txt1714.Text = pfui.Cell(cell[34]).CachedValue.ToString().Replace('.', ',');
            txt1715.Text = pfui.Cell(cell[35]).CachedValue.ToString().Replace('.', ',');
            txt1716.Text = pfui.Cell(cell[36]).CachedValue.ToString().Replace('.', ',');
            txt1717.Text = pfui.Cell(cell[37]).CachedValue.ToString().Replace('.', ',');
            txt1718.Text = pfui.Cell(cell[38]).CachedValue.ToString().Replace('.', ',');
            txt1719.Text = pfui.Cell(cell[39]).CachedValue.ToString().Replace('.', ',');
            txt1720.Text = pfui.Cell(cell[40]).CachedValue.ToString().Replace('.', ',');
            //CRONOGRAMA - PRIMEIRA PÁGINA
            txtExecutado.Text = pfui.Cell(cell[41]).CachedValue.ToString().Replace('.', ',');
            txtParcela1.Text = pfui.Cell(cell[42]).CachedValue.ToString().Replace('.', ',');
            txtParcela2.Text = pfui.Cell(cell[43]).CachedValue.ToString().Replace('.', ',');
            txtParcela3.Text = pfui.Cell(cell[44]).CachedValue.ToString().Replace('.', ',');
            txtParcela4.Text = pfui.Cell(cell[45]).CachedValue.ToString().Replace('.', ',');
            txtParcela5.Text = pfui.Cell(cell[46]).CachedValue.ToString().Replace('.', ',');
            txtParcela6.Text = pfui.Cell(cell[47]).CachedValue.ToString().Replace('.', ',');
            txtParcela7.Text = pfui.Cell(cell[48]).CachedValue.ToString().Replace('.', ',');
            txtParcela8.Text = pfui.Cell(cell[49]).CachedValue.ToString().Replace('.', ',');
            //CRONOGRAMA - SEGUNDA PÁGINA
            txtParcela9.Text = pfui.Cell(cell[50]).CachedValue.ToString().Replace('.', ',');
            txtParcela10.Text = pfui.Cell(cell[51]).CachedValue.ToString().Replace('.', ',');
            txtParcela11.Text = pfui.Cell(cell[52]).CachedValue.ToString().Replace('.', ',');
            txtParcela12.Text = pfui.Cell(cell[53]).CachedValue.ToString().Replace('.', ',');
            txtParcela13.Text = pfui.Cell(cell[54]).CachedValue.ToString().Replace('.', ',');
            txtParcela14.Text = pfui.Cell(cell[55]).CachedValue.ToString().Replace('.', ',');
            txtParcela15.Text = pfui.Cell(cell[56]).CachedValue.ToString().Replace('.', ',');
            txtParcela16.Text = pfui.Cell(cell[57]).CachedValue.ToString().Replace('.', ',');
        }
        //----------------------------------------//




        //--------------Shared Event--------------//
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
        //----------------------------------------//







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

                    string version = GetVersionPFUI(planilha.PageSetup.Header.Right.GetText(XLHFOccurrence.OddPages).ToString());
                    if (version == null)
                    {
                        MessageBox.Show("Arquivo incompatível ou versão não suportada!", "Planilha PFUI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        string[] aeX = SetCellsPFUI(version);
                        PopulateFromPFUI(aeX, planilha);
                        pnlMainPfui.Show();
                        btnProximoPfui.Show();
                        txtPropNome.Focus();
                    }
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
