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
        private string[] SetArrayForThisPFUIVersion(string version)
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
        private void GetAndPopulate_PFUI(string[] aeX, ISheet ws)
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
                        string[] aeX = SetArrayForThisPFUIVersion(version);
                        GetAndPopulate_PFUI(aeX, sheet);
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

        private void button2_Click(object sender, EventArgs e)
        {

        }
    }
}
