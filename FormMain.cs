using System;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using ClosedXML.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace Raecef
{
    public partial class FormMain : Form
    {


        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);



        string _caminhoModelo;





        //------------------METHODS------------------//
        private void ImportPFUI(string[] aeX, ISheet ws)
        {
            //CABEÇALHO
            txtPropNome.Text = ws.GetRow(new CellReference(aeX[0]).Row).GetCell(new CellReference(aeX[0]).Col).ToString();
            txtPropCPF.Text = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[1]).Row).GetCell(new CellReference(aeX[1]).Col).ToString());
            txtPropDDD.Text = ws.GetRow(new CellReference(aeX[2]).Row).GetCell(new CellReference(aeX[2]).Col).ToString();
            txtPropTelefone.Text = Util.FormatedFone(ws.GetRow(new CellReference(aeX[3]).Row).GetCell(new CellReference(aeX[3]).Col).ToString());
            txtRTNome.Text = ws.GetRow(new CellReference(aeX[4]).Row).GetCell(new CellReference(aeX[4]).Col).ToString();
            txtRTCauCrea.Text = ws.GetRow(new CellReference(aeX[5]).Row).GetCell(new CellReference(aeX[5]).Col).ToString();     //.Replace(',', '.');
            txtRTUF.Text = ws.GetRow(new CellReference(aeX[6]).Row).GetCell(new CellReference(aeX[6]).Col).ToString();
            txtRTCPF.Text = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[7]).Row).GetCell(new CellReference(aeX[7]).Col).ToString());
            txtRTDDD.Text = ws.GetRow(new CellReference(aeX[8]).Row).GetCell(new CellReference(aeX[8]).Col).ToString();
            txtRTTelefone.Text = Util.FormatedFone(ws.GetRow(new CellReference(aeX[9]).Row).GetCell(new CellReference(aeX[9]).Col).ToString());
            txtIdEndereco.Text = ws.GetRow(new CellReference(aeX[10]).Row).GetCell(new CellReference(aeX[10]).Col).ToString();
            txtIdComplemento.Text = ws.GetRow(new CellReference(aeX[11]).Row).GetCell(new CellReference(aeX[11]).Col).ToString();
            txtIdCEP.Text = ws.GetRow(new CellReference(aeX[12]).Row).GetCell(new CellReference(aeX[12]).Col).ToString();
            txtIdBairro.Text = ws.GetRow(new CellReference(aeX[13]).Row).GetCell(new CellReference(aeX[13]).Col).ToString();
            txtIdMunicipio.Text = ws.GetRow(new CellReference(aeX[14]).Row).GetCell(new CellReference(aeX[14]).Col).ToString();
            txtIdUF.Text = ws.GetRow(new CellReference(aeX[15]).Row).GetCell(new CellReference(aeX[15]).Col).ToString();
            txtTerrenoValorProposto.Text = string.Format("{0:0,0.00}", Convert.ToDouble(ws.GetRow(new CellReference(aeX[16]).Row).GetCell(new CellReference(aeX[16]).Col).ToString()));
            txtTerrenoMatricula.Text = ws.GetRow(new CellReference(aeX[17]).Row).GetCell(new CellReference(aeX[17]).Col).ToString();     //.Replace(',', '.');
            txtTerrenoOficio.Text = ws.GetRow(new CellReference(aeX[18]).Row).GetCell(new CellReference(aeX[18]).Col).ToString();
            txtTerrenoComarca.Text = ws.GetRow(new CellReference(aeX[19]).Row).GetCell(new CellReference(aeX[19]).Col).ToString();
            txtTerrenoUF.Text = ws.GetRow(new CellReference(aeX[20]).Row).GetCell(new CellReference(aeX[20]).Col).ToString();
            //ORÇAMENTO (PERCENTUAIS)
            txt1701.Text = ws.GetRow(new CellReference(aeX[21]).Row).GetCell(new CellReference(aeX[21]).Col).NumericCellValue.ToString();
            txt1702.Text = ws.GetRow(new CellReference(aeX[22]).Row).GetCell(new CellReference(aeX[22]).Col).NumericCellValue.ToString();
            txt1703.Text = ws.GetRow(new CellReference(aeX[23]).Row).GetCell(new CellReference(aeX[23]).Col).NumericCellValue.ToString();
            txt1704.Text = ws.GetRow(new CellReference(aeX[24]).Row).GetCell(new CellReference(aeX[24]).Col).NumericCellValue.ToString();
            txt1705.Text = ws.GetRow(new CellReference(aeX[25]).Row).GetCell(new CellReference(aeX[25]).Col).NumericCellValue.ToString();
            txt1706.Text = ws.GetRow(new CellReference(aeX[26]).Row).GetCell(new CellReference(aeX[26]).Col).NumericCellValue.ToString();
            txt1707.Text = ws.GetRow(new CellReference(aeX[27]).Row).GetCell(new CellReference(aeX[27]).Col).NumericCellValue.ToString();
            txt1708.Text = ws.GetRow(new CellReference(aeX[28]).Row).GetCell(new CellReference(aeX[28]).Col).NumericCellValue.ToString();
            txt1709.Text = ws.GetRow(new CellReference(aeX[29]).Row).GetCell(new CellReference(aeX[29]).Col).NumericCellValue.ToString();
            txt1710.Text = ws.GetRow(new CellReference(aeX[30]).Row).GetCell(new CellReference(aeX[30]).Col).NumericCellValue.ToString();
            txt1711.Text = ws.GetRow(new CellReference(aeX[31]).Row).GetCell(new CellReference(aeX[31]).Col).NumericCellValue.ToString();
            txt1712.Text = ws.GetRow(new CellReference(aeX[32]).Row).GetCell(new CellReference(aeX[32]).Col).NumericCellValue.ToString();
            txt1713.Text = ws.GetRow(new CellReference(aeX[33]).Row).GetCell(new CellReference(aeX[33]).Col).NumericCellValue.ToString();
            txt1714.Text = ws.GetRow(new CellReference(aeX[34]).Row).GetCell(new CellReference(aeX[34]).Col).NumericCellValue.ToString();
            txt1715.Text = ws.GetRow(new CellReference(aeX[35]).Row).GetCell(new CellReference(aeX[35]).Col).NumericCellValue.ToString();
            txt1716.Text = ws.GetRow(new CellReference(aeX[36]).Row).GetCell(new CellReference(aeX[36]).Col).NumericCellValue.ToString();
            txt1717.Text = ws.GetRow(new CellReference(aeX[37]).Row).GetCell(new CellReference(aeX[37]).Col).NumericCellValue.ToString();
            txt1718.Text = ws.GetRow(new CellReference(aeX[38]).Row).GetCell(new CellReference(aeX[38]).Col).NumericCellValue.ToString();
            txt1719.Text = ws.GetRow(new CellReference(aeX[39]).Row).GetCell(new CellReference(aeX[39]).Col).NumericCellValue.ToString();
            txt1720.Text = ws.GetRow(new CellReference(aeX[40]).Row).GetCell(new CellReference(aeX[40]).Col).NumericCellValue.ToString();
            //CRONOGRAMA - PRIMEIRA PÁGINA
            txtExecutado.Text = ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col).NumericCellValue.ToString();
            txtParcela1.Text = ws.GetRow(new CellReference(aeX[42]).Row).GetCell(new CellReference(aeX[42]).Col).NumericCellValue.ToString();
            txtParcela2.Text = ws.GetRow(new CellReference(aeX[43]).Row).GetCell(new CellReference(aeX[43]).Col).NumericCellValue.ToString();
            txtParcela3.Text = ws.GetRow(new CellReference(aeX[44]).Row).GetCell(new CellReference(aeX[44]).Col).NumericCellValue.ToString();
            txtParcela4.Text = ws.GetRow(new CellReference(aeX[45]).Row).GetCell(new CellReference(aeX[45]).Col).NumericCellValue.ToString();
            txtParcela5.Text = ws.GetRow(new CellReference(aeX[46]).Row).GetCell(new CellReference(aeX[46]).Col).NumericCellValue.ToString();
            txtParcela6.Text = ws.GetRow(new CellReference(aeX[47]).Row).GetCell(new CellReference(aeX[47]).Col).NumericCellValue.ToString();
            txtParcela7.Text = ws.GetRow(new CellReference(aeX[48]).Row).GetCell(new CellReference(aeX[48]).Col).NumericCellValue.ToString();
            txtParcela8.Text = ws.GetRow(new CellReference(aeX[49]).Row).GetCell(new CellReference(aeX[49]).Col).NumericCellValue.ToString();
            ////CRONOGRAMA - SEGUNDA PÁGINA
            txtParcela9.Text = ws.GetRow(new CellReference(aeX[50]).Row).GetCell(new CellReference(aeX[50]).Col).NumericCellValue.ToString();
            txtParcela10.Text = ws.GetRow(new CellReference(aeX[51]).Row).GetCell(new CellReference(aeX[51]).Col).NumericCellValue.ToString();
            txtParcela11.Text = ws.GetRow(new CellReference(aeX[52]).Row).GetCell(new CellReference(aeX[52]).Col).NumericCellValue.ToString();
            txtParcela12.Text = ws.GetRow(new CellReference(aeX[53]).Row).GetCell(new CellReference(aeX[53]).Col).NumericCellValue.ToString();
            txtParcela13.Text = ws.GetRow(new CellReference(aeX[54]).Row).GetCell(new CellReference(aeX[54]).Col).NumericCellValue.ToString();
            txtParcela14.Text = ws.GetRow(new CellReference(aeX[55]).Row).GetCell(new CellReference(aeX[55]).Col).NumericCellValue.ToString();
            txtParcela15.Text = ws.GetRow(new CellReference(aeX[56]).Row).GetCell(new CellReference(aeX[56]).Col).NumericCellValue.ToString();
            txtParcela16.Text = ws.GetRow(new CellReference(aeX[57]).Row).GetCell(new CellReference(aeX[57]).Col).NumericCellValue.ToString();

        }
        private void ExportRAE(string[] rae, ClosedXML.Excel.IXLWorksheet ws)
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
        //-------------------------------------------//












        //------------------SHARED EVENT-------------------//
        private void NextTabControl(object sender, EventArgs e)
        {
            tabControl.SelectTab(tabControl.SelectedIndex + 1);

        }
        //------------------------------------------------//







        







        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            tabControl.ItemSize = new System.Drawing.Size(0, 1);
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
            if (tabControl.SelectedIndex == 1)
            {
                tabControl.SelectTab(0);
                btnBack.Hide();
            }
            else if (tabControl.SelectedIndex == 2)
                tabControl.SelectTab(1);
            else if (tabControl.SelectedIndex == 3)
                tabControl.SelectTab(2);
            else if (tabControl.SelectedIndex == 4)
                tabControl.SelectTab(3);

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
                    FileStream arquivoXLS = new FileStream(openExcel.FileName, FileMode.Open, FileAccess.Read);

                    HSSFWorkbook wbook = new HSSFWorkbook(arquivoXLS);
                    ISheet sheet = wbook.GetSheet("Proposta");

                    string version = SanitizeVersion(sheet.Header.Right);
                    if (version == null || version == "")
                    {
                        MessageBox.Show("Arquivo incompatível ou versão não suportada!", "Planilha PFUI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        lblVersion.Text = version;
                        string[] aeX = PFUI.SetArray(version);
                        ImportPFUI(aeX, sheet);
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

                        ExportRAE(RAE.SetArray("AE 130 020"), rae);
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
            NextTabControl(sender, e);
            btnBack.Show();
        }




    }
}
