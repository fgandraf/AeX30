using System;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using ClosedXML.Excel;
using Spire.Xls;

namespace Raecef
{
    public partial class FormMain : Form
    {



        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);


        string _caminhoModelo;

        //------------------PFUI------------------//
        string[] ae130v017_018 = new string[]
                {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR116", "AR118", "AR130", "AR137", "AR146", "AR156", "AR165", "AR172", "AR179", "AR190", "AR197", "AR207", "AR217", "AR228", "AR234", "AR245", "AR254", "AR262", "AR271", "AR273",
            "AJ311",
            "AL310", "AP310", "AT310", "AX310", "BB310", "BF310", "BJ310", "BN310",
            "AL350", "AP350", "AT350", "AX350", "BB350", "BF350", "BJ350", "BN350"
                };
        string[] ae130v019_020 = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR116", "AR118", "AR130", "AR137", "AR146", "AR156", "AR165", "AR172", "AR179", "AR190", "AR197", "AR207", "AR217", "AR228", "AR234", "AR245", "AR254", "AR262", "AR271", "AR273",
            "AJ312",
            "AL311", "AP311", "AT311", "AX311", "BB311", "BF311", "BJ311", "BN311",
            "AL351", "AP351", "AT351", "AX351", "BB351", "BF351", "BJ351", "BN351"
        };
        string[] ae130v021 = new string[]
        {
            "G42", "AA42", "AG42", "AI42", "AN42", "AX42", "BD42", "BF42", "BL42", "BN42",
            "G46", "AP46", "BA46", "BF46",
            "G48", "AB48",
            "AG50", "AP50", "AX50", "BF50", "BN50",
            "AR114", "AR116", "AR128", "AR135", "AR144", "AR154", "AR163", "AR170", "AR177", "AR188", "AR195", "AR205", "AR215", "AR226", "AR232", "AR243", "AR252", "AR260", "AR269", "AR271",
            "AJ310",
            "AL309", "AP309", "AT309", "AX309", "BB309", "BF309", "BJ309", "BN309",
            "AL349", "AP349", "AT349", "AX349", "BB349", "BF349", "BJ349", "BN349"
        };
        private string[] ChooseArray_PFUI(string version)
        {
            if (version == "AE 130 017" || version == "AE 130 018")
                return ae130v017_018;
            else if (version == "AE 130 019" || version == "AE 130 020")
                return ae130v019_020;
            else if (version == "AE 130 021")
                return ae130v021;
            else
            {
                MessageBox.Show("A versão da planilha PFUI inserida não foi testada.\r\nRedobre a atenção quanto aos valores importados!", "Planilha PFUI não testada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return ae130v021;
            }
        }
        private void GetAndPopulate_PFUI(string[] aeX, ClosedXML.Excel.IXLWorksheet ws)
        {
            //CABEÇALHO
            txtPropNome.Text = ws.Cell(aeX[0]).RichText.ToString().ToUpper();
            txtPropCPF.Text = Util.FormatedCPF(ws.Cell(aeX[1]).RichText.ToString());
            txtPropDDD.Text = ws.Cell(aeX[2]).RichText.ToString();
            txtPropTelefone.Text = Util.FormatedFone(ws.Cell(aeX[3]).RichText.ToString());
            txtRTNome.Text = ws.Cell(aeX[4]).RichText.ToString().ToUpper();
            txtRTCauCrea.Text = ws.Cell(aeX[5]).RichText.ToString().Replace(',', '.');
            txtRTUF.Text = ws.Cell(aeX[6]).RichText.ToString();
            txtRTCPF.Text = Util.FormatedCPF(ws.Cell(aeX[7]).Value.ToString());
            txtRTDDD.Text = ws.Cell(aeX[8]).RichText.ToString();
            txtRTTelefone.Text = Util.FormatedFone(ws.Cell(aeX[9]).RichText.ToString());
            txtIdEndereco.Text = ws.Cell(aeX[10]).RichText.ToString().ToUpper();
            txtIdComplemento.Text = ws.Cell(aeX[11]).RichText.ToString().ToUpper();
            txtIdCEP.Text = Util.FormatedCEP(ws.Cell(aeX[12]).RichText.ToString());
            txtIdBairro.Text = ws.Cell(aeX[13]).RichText.ToString().ToUpper();
            txtIdMunicipio.Text = ws.Cell(aeX[14]).RichText.ToString().ToUpper();
            txtIdUF.Text = ws.Cell(aeX[15]).RichText.ToString();
            txtTerrenoValorProposto.Text = string.Format("{0:0,0.00}", Convert.ToDouble(ws.Cell(aeX[16]).CachedValue));
            txtTerrenoMatricula.Text = ws.Cell(aeX[17]).RichText.ToString().Replace(',', '.');
            txtTerrenoOficio.Text = ws.Cell(aeX[18]).RichText.ToString().ToUpper();
            txtTerrenoComarca.Text = ws.Cell(aeX[19]).RichText.ToString().ToUpper();
            txtTerrenoUF.Text = ws.Cell(aeX[20]).RichText.ToString();
            //ORÇAMENTO (PERCENTUAIS)
            txt1701.Text = ws.Cell(aeX[21]).CachedValue.ToString().Replace('.', ',');
            txt1702.Text = ws.Cell(aeX[22]).CachedValue.ToString().Replace('.', ',');
            txt1703.Text = ws.Cell(aeX[23]).CachedValue.ToString().Replace('.', ',');
            txt1704.Text = ws.Cell(aeX[24]).CachedValue.ToString().Replace('.', ',');
            txt1705.Text = ws.Cell(aeX[25]).CachedValue.ToString().Replace('.', ',');
            txt1706.Text = ws.Cell(aeX[26]).CachedValue.ToString().Replace('.', ',');
            txt1707.Text = ws.Cell(aeX[27]).CachedValue.ToString().Replace('.', ',');
            txt1708.Text = ws.Cell(aeX[28]).CachedValue.ToString().Replace('.', ',');
            txt1709.Text = ws.Cell(aeX[29]).CachedValue.ToString().Replace('.', ',');
            txt1710.Text = ws.Cell(aeX[30]).CachedValue.ToString().Replace('.', ',');
            txt1711.Text = ws.Cell(aeX[31]).CachedValue.ToString().Replace('.', ',');
            txt1712.Text = ws.Cell(aeX[32]).CachedValue.ToString().Replace('.', ',');
            txt1713.Text = ws.Cell(aeX[33]).CachedValue.ToString().Replace('.', ',');
            txt1714.Text = ws.Cell(aeX[34]).CachedValue.ToString().Replace('.', ',');
            txt1715.Text = ws.Cell(aeX[35]).CachedValue.ToString().Replace('.', ',');
            txt1716.Text = ws.Cell(aeX[36]).CachedValue.ToString().Replace('.', ',');
            txt1717.Text = ws.Cell(aeX[37]).CachedValue.ToString().Replace('.', ',');
            txt1718.Text = ws.Cell(aeX[38]).CachedValue.ToString().Replace('.', ',');
            txt1719.Text = ws.Cell(aeX[39]).CachedValue.ToString().Replace('.', ',');
            txt1720.Text = ws.Cell(aeX[40]).CachedValue.ToString().Replace('.', ',');
            //CRONOGRAMA - PRIMEIRA PÁGINA
            txtExecutado.Text = ws.Cell(aeX[41]).CachedValue.ToString().Replace('.', ',');
            txtParcela1.Text = ws.Cell(aeX[42]).CachedValue.ToString().Replace('.', ',');
            txtParcela2.Text = ws.Cell(aeX[43]).CachedValue.ToString().Replace('.', ',');
            txtParcela3.Text = ws.Cell(aeX[44]).CachedValue.ToString().Replace('.', ',');
            txtParcela4.Text = ws.Cell(aeX[45]).CachedValue.ToString().Replace('.', ',');
            txtParcela5.Text = ws.Cell(aeX[46]).CachedValue.ToString().Replace('.', ',');
            txtParcela6.Text = ws.Cell(aeX[47]).CachedValue.ToString().Replace('.', ',');
            txtParcela7.Text = ws.Cell(aeX[48]).CachedValue.ToString().Replace('.', ',');
            txtParcela8.Text = ws.Cell(aeX[49]).CachedValue.ToString().Replace('.', ',');
            //CRONOGRAMA - SEGUNDA PÁGINA
            txtParcela9.Text = ws.Cell(aeX[50]).CachedValue.ToString().Replace('.', ',');
            txtParcela10.Text = ws.Cell(aeX[51]).CachedValue.ToString().Replace('.', ',');
            txtParcela11.Text = ws.Cell(aeX[52]).CachedValue.ToString().Replace('.', ',');
            txtParcela12.Text = ws.Cell(aeX[53]).CachedValue.ToString().Replace('.', ',');
            txtParcela13.Text = ws.Cell(aeX[54]).CachedValue.ToString().Replace('.', ',');
            txtParcela14.Text = ws.Cell(aeX[55]).CachedValue.ToString().Replace('.', ',');
            txtParcela15.Text = ws.Cell(aeX[56]).CachedValue.ToString().Replace('.', ',');
            txtParcela16.Text = ws.Cell(aeX[57]).CachedValue.ToString().Replace('.', ',');
        }
        //----------------------------------------//





        //------------------RAE-------------------//
        string[] rae130v020 = new string[]
                {
            "AD35", "AF35", "AH35", "AL35", "AN35", "AO35", "AP35", // 0 a 6
            "G43", "AJ43", "AP43", "AR43", // 7 a 10
            "G46", "Z46", "AH46", "AJ46", "AP46", "AR46", // 11 a 16
            "G49", "AJ49", // 17 a 18
            "G51", "V51", "AA51", "AS51", // 19 a 22
            "G53", "Q53", "AA53", "AJ53", "AS53", // 23 a 27
            "S68","S69","S70","S71","S72","S73","S74","S75","S76","S77","S78","S79","S80","S81","S82","S83","S84","S85","S86","S87",
            "Y89",
            "AH63", "AS63",
            "AS74",
            "BC61","BC62","BC63","BC64","BC65","BC66","BC67","BC68","BC69","BC70","BC71","BC72","BC73","BC74","BC75","BC76","BC77"
                };
        private void Populate_RAE(string[] rae, ClosedXML.Excel.IXLWorksheet ws)
        {
            //CABEÇALHO
            ws.Cell(rae[0]).Value = txtRef0.Text;
            ws.Cell(rae[1]).Value = txtRef1.Text;
            ws.Cell(rae[2]).Value = txtRef2.Text;
            ws.Cell(rae[3]).Value = txtRef3.Text;
            ws.Cell(rae[4]).Value = txtRef4.Text;
            ws.Cell(rae[5]).Value = txtRef5.Text;
            ws.Cell(rae[6]).Value = txtRef6.Text;
            ws.Cell(rae[7]).Value = txtPropNome.Text;
            ws.Cell(rae[8]).Value = txtPropCPF.Text;
            ws.Cell(rae[9]).Value = txtPropDDD.Text;
            ws.Cell(rae[10]).Value = txtPropTelefone.Text;
            ws.Cell(rae[11]).Value = txtRTNome.Text;
            ws.Cell(rae[12]).Value = txtRTCauCrea.Text;
            ws.Cell(rae[13]).Value = txtRTUF.Text;
            ws.Cell(rae[14]).Value = txtRTCPF.Text;
            ws.Cell(rae[15]).Value = txtRTDDD.Text;
            ws.Cell(rae[16]).Value = txtRTTelefone.Text;
            ws.Cell(rae[17]).Value = txtIdEndereco.Text;
            ws.Cell(rae[18]).Value = txtIdComplemento.Text;
            ws.Cell(rae[19]).Value = txtIdBairro.Text;
            ws.Cell(rae[20]).Value = txtIdCEP.Text;
            ws.Cell(rae[21]).Value = txtIdMunicipio.Text;
            ws.Cell(rae[22]).Value = txtIdUF.Text;
            ws.Cell(rae[23]).Value = txtTerrenoValorProposto.Text;
            ws.Cell(rae[24]).Value = txtTerrenoMatricula.Text;
            ws.Cell(rae[25]).Value = txtTerrenoOficio.Text;
            ws.Cell(rae[26]).Value = txtTerrenoComarca.Text;
            ws.Cell(rae[27]).Value = txtTerrenoUF.Text;
            //ORÇAMENTO (PERCENTUAIS)
            ws.Cell(rae[28]).Value = Convert.ToDouble(txt1701.Text);
            ws.Cell(rae[29]).Value = Convert.ToDouble(txt1702.Text);
            ws.Cell(rae[30]).Value = Convert.ToDouble(txt1703.Text);
            ws.Cell(rae[31]).Value = Convert.ToDouble(txt1704.Text);
            ws.Cell(rae[32]).Value = Convert.ToDouble(txt1705.Text);
            ws.Cell(rae[33]).Value = Convert.ToDouble(txt1706.Text);
            ws.Cell(rae[34]).Value = Convert.ToDouble(txt1707.Text);
            ws.Cell(rae[35]).Value = Convert.ToDouble(txt1708.Text);
            ws.Cell(rae[36]).Value = Convert.ToDouble(txt1709.Text);
            ws.Cell(rae[37]).Value = Convert.ToDouble(txt1710.Text);
            ws.Cell(rae[38]).Value = Convert.ToDouble(txt1711.Text);
            ws.Cell(rae[39]).Value = Convert.ToDouble(txt1712.Text);
            ws.Cell(rae[40]).Value = Convert.ToDouble(txt1713.Text);
            ws.Cell(rae[41]).Value = Convert.ToDouble(txt1714.Text);
            ws.Cell(rae[42]).Value = Convert.ToDouble(txt1715.Text);
            ws.Cell(rae[43]).Value = Convert.ToDouble(txt1716.Text);
            ws.Cell(rae[44]).Value = Convert.ToDouble(txt1717.Text);
            ws.Cell(rae[45]).Value = Convert.ToDouble(txt1718.Text);
            ws.Cell(rae[46]).Value = Convert.ToDouble(txt1719.Text);
            ws.Cell(rae[47]).Value = Convert.ToDouble(txt1720.Text);
            if (txtMensuradoAcumulado.Text != "")
                ws.Cell(rae[48]).Value = Convert.ToDouble(txtMensuradoAcumulado.Text);
            //DADOS ADICIONAIS
            if (txtContratoInicio.Text != "  /  /")
                ws.Cell(rae[49]).Value = txtContratoInicio.Text;
            if (txtContratoTermino.Text != "  /  /")
                ws.Cell(rae[50]).Value = txtContratoTermino.Text;
            ws.Cell(rae[51]).Value = txtEtapaPrevista.Text;
            //CRONOGRAMA
            ws.Cell(rae[52]).Value = Convert.ToDouble(txtExecutado.Text);
            ws.Cell(rae[53]).Value = Convert.ToDouble(txtParcela1.Text);
            ws.Cell(rae[54]).Value = Convert.ToDouble(txtParcela2.Text);
            ws.Cell(rae[55]).Value = Convert.ToDouble(txtParcela3.Text);
            ws.Cell(rae[56]).Value = Convert.ToDouble(txtParcela4.Text);
            ws.Cell(rae[57]).Value = Convert.ToDouble(txtParcela5.Text);
            ws.Cell(rae[58]).Value = Convert.ToDouble(txtParcela6.Text);
            ws.Cell(rae[59]).Value = Convert.ToDouble(txtParcela7.Text);
            ws.Cell(rae[60]).Value = Convert.ToDouble(txtParcela8.Text);
            ws.Cell(rae[61]).Value = Convert.ToDouble(txtParcela9.Text);
            ws.Cell(rae[62]).Value = Convert.ToDouble(txtParcela10.Text);
            ws.Cell(rae[63]).Value = Convert.ToDouble(txtParcela11.Text);
            ws.Cell(rae[64]).Value = Convert.ToDouble(txtParcela12.Text);
            ws.Cell(rae[65]).Value = Convert.ToDouble(txtParcela13.Text);
            ws.Cell(rae[66]).Value = Convert.ToDouble(txtParcela14.Text);
            ws.Cell(rae[67]).Value = Convert.ToDouble(txtParcela15.Text);
            ws.Cell(rae[68]).Value = Convert.ToDouble(txtParcela16.Text);


        }
        //----------------------------------------//




        private string SanitizeVersion(string header)
        {
            if (header != "")
            {
                long version = Convert.ToInt64(Util.CleanInput(header));

                if (version == 101413010017)
                    version = 1413010017;

                if (version == 1413010017)
                    return "AE 130 017";
                else if (version == 1413010018)
                    return "AE 130 018";
                else if (version == 1413010019)
                    return "AE 130 019";
                else if (version == 1413010020)
                    return "AE 130 020";
                else if (version == 1413010021)
                    return "AE 130 021";
                else if (version > 1413010021)
                    return "> AE 130 021";
                else
                    return null;
            }
            else
                return null;
        }






        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            TabControl.ItemSize = new System.Drawing.Size(0, 1);
            //TabControl.SelectTab(2);
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
                    Workbook workbook = new Workbook();
                    workbook.LoadFromFile(openExcel.FileName);
                    workbook.SaveToFile("converted.xlsx");

                    
                    var wb = new XLWorkbook("converted.xlsx");
                    var ws = wb.Worksheets.First(w => w.Name == "Proposta");

                    lblVersion.Text = SanitizeVersion(ws.PageSetup.Header.Right.GetText(XLHFOccurrence.OddPages));

                    if (lblVersion.Text == null || lblVersion.Text == "")
                    {
                        lblVersion.Text = "---";
                        MessageBox.Show("Arquivo incompatível ou versão não suportada!", "Planilha PFUI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        string[] aeX = ChooseArray_PFUI(lblVersion.Text);
                        GetAndPopulate_PFUI(aeX, ws);

                        File.Delete("converted.xlsx");
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






        private void btnModeloPadrao_Click(object sender, EventArgs e)
        {
            try
            {
                openExcel.Filter = "Planilhas do Excel habilitada p/ macro | *.xlsm";
                if (openExcel.ShowDialog() == DialogResult.OK)
                {
                    _caminhoModelo = openExcel.FileName;

                    txtLogFinalizar.Text = "Modelo padrão: " + openExcel.FileName + "\r\n";
                    txtLogFinalizar.Text += "Pronto para gravar. Aguardando confirmação do usuário...\r\n";
                    pnlMainFinalizar.Show();
                    btnConfirmar.Show();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Mensagem de erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        private void btnConfirmar_Click(object sender, EventArgs e)
        {
            try
            {
                //CREATING FILENAME FROM "txtPropNome" USING UPPERCASE FOR THE FIRST CHARACTERE
                string[] proponente = txtPropNome.Text.ToLower().Split(' ');
                proponente[0] = proponente[0].Substring(0, 1).ToUpper() + proponente[0].Substring(1);
                proponente[proponente.Length - 1] = proponente[proponente.Length - 1].Substring(0, 1).ToUpper() + proponente[proponente.Length - 1].Substring(1);
                saveExcel.FileName = "RAE_" + proponente[0] + "-" + proponente[proponente.Length - 1] + ".xlsx";


                if (saveExcel.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(_caminhoModelo))
                    {
                        var wbook = new XLWorkbook(_caminhoModelo);
                        var rae = wbook.Worksheets.First(w => w.Name == "RAE");

                        Populate_RAE(rae130v020, rae);
                        wbook.SaveAs(saveExcel.FileName);

                        txtLogFinalizar.Text += "--------------------------------\r\nConcluído!\r\n";
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
