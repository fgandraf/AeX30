using System;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace aeX30
{
    public partial class FormMain : Form
    {


        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);



        private string _caminhoModelo;
        private string _rae;





        //:METHODS
        private void ImportPFUI(string[] aeX, ISheet ws)
        {
            //CABEÇALHO
            txtPropNome.Text = ws.GetRow(new CellReference(aeX[0]).Row).GetCell(new CellReference(aeX[0]).Col).ToString();
            txtPropCPF.Text = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[1]).Row).GetCell(new CellReference(aeX[1]).Col).ToString());
            txtPropDDD.Text = ws.GetRow(new CellReference(aeX[2]).Row).GetCell(new CellReference(aeX[2]).Col).ToString();
            txtPropTelefone.Text = Util.FormatedFone(ws.GetRow(new CellReference(aeX[3]).Row).GetCell(new CellReference(aeX[3]).Col).ToString());
            txtRTNome.Text = ws.GetRow(new CellReference(aeX[4]).Row).GetCell(new CellReference(aeX[4]).Col).ToString();
            txtRTCauCrea.Text = ws.GetRow(new CellReference(aeX[5]).Row).GetCell(new CellReference(aeX[5]).Col).ToString();
            txtRTUF.Text = ws.GetRow(new CellReference(aeX[6]).Row).GetCell(new CellReference(aeX[6]).Col).ToString();
            txtRTCPF.Text = Util.FormatedCPF(ws.GetRow(new CellReference(aeX[7]).Row).GetCell(new CellReference(aeX[7]).Col).ToString());   
            txtRTDDD.Text = ws.GetRow(new CellReference(aeX[8]).Row).GetCell(new CellReference(aeX[8]).Col).ToString();
            txtRTTelefone.Text = Util.FormatedFone(ws.GetRow(new CellReference(aeX[9]).Row).GetCell(new CellReference(aeX[9]).Col).ToString());
            txtIdEndereco.Text = ws.GetRow(new CellReference(aeX[10]).Row).GetCell(new CellReference(aeX[10]).Col).ToString();
            txtIdComplemento.Text = ws.GetRow(new CellReference(aeX[11]).Row).GetCell(new CellReference(aeX[11]).Col).ToString();
            txtIdCEP.Text = Util.FormatedCEP(ws.GetRow(new CellReference(aeX[12]).Row).GetCell(new CellReference(aeX[12]).Col).ToString());
            txtIdBairro.Text = ws.GetRow(new CellReference(aeX[13]).Row).GetCell(new CellReference(aeX[13]).Col).ToString();
            txtIdMunicipio.Text = ws.GetRow(new CellReference(aeX[14]).Row).GetCell(new CellReference(aeX[14]).Col).ToString();
            txtIdUF.Text = ws.GetRow(new CellReference(aeX[15]).Row).GetCell(new CellReference(aeX[15]).Col).ToString();
            txtTerrenoValorProposto.Text = string.Format("{0:0,0.00}", Convert.ToDouble(ws.GetRow(new CellReference(aeX[16]).Row).GetCell(new CellReference(aeX[16]).Col).ToString()));
            txtTerrenoMatricula.Text = ws.GetRow(new CellReference(aeX[17]).Row).GetCell(new CellReference(aeX[17]).Col).ToString();
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
            //CRONOGRAMA
            ///Executado
            if (ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col) != null)
                txtExecutado.Text = ws.GetRow(new CellReference(aeX[41]).Row).GetCell(new CellReference(aeX[41]).Col).NumericCellValue.ToString();
            ///Parcelas 1 a 30
            int txtBox = 1;
            int arr = 42;
            var parcelaX = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
            while (parcelaX != null)
            {
                foreach (Control c in this.pnlPFUIParcelas.Controls)
                {
                    if (c.Name == "txtParcela" + txtBox.ToString())
                        c.Text = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col).NumericCellValue.ToString();
                }
                txtBox++;
                arr++;
                parcelaX = ws.GetRow(new CellReference(aeX[arr]).Row).GetCell(new CellReference(aeX[arr]).Col);
            }
        }

        private string SanitizeVersion(string header)
        {
            if (header != "")
            {
                long version = Convert.ToInt64(Util.CleanInput(header));

                if (version == 101413010017)
                    version = 1413010017;



                if (version == 1413010016)
                    return "AE 130 016";
                else if (version == 1413010017)
                    return "AE 130 017";
                else if (version == 1413010018)
                    return "AE 130 018";
                else if (version == 1413010019)
                    return "AE 130 019";
                else if (version == 1413010020)
                    return "AE 130 020";
                else if (version == 1413010021)
                    return "AE 130 021";
                else if (version == 1413010022)
                    return "AE 130 022";
                else if (version == 1413010023)
                    return "AE 130 023";
                else if (version == 1413010024)
                    return "AE 130 024";
                else if (version == 1413010025)
                    return "AE 130 025";
                else if (version > 1413010025)
                    return "> AE 130 025";
                else
                    return null;
            }
            else
                return null;
        }

        private void ExportRAE(string caminho)
        {
                 HSSFWorkbook wbook;

                using (FileStream arquivoModelo = new FileStream(_caminhoModelo, FileMode.Open, FileAccess.Read))
                {
                    wbook = new HSSFWorkbook(arquivoModelo);
                }
                ISheet ws = wbook.GetSheet("RAE");



                string[] rae = RAE.SetArray(_rae);



                //CABEÇALHO
                ws.GetRow(new CellReference(rae[0]).Row).GetCell(new CellReference(rae[0]).Col).SetCellValue(txtRef0.Text);
                ws.GetRow(new CellReference(rae[1]).Row).GetCell(new CellReference(rae[1]).Col).SetCellValue(txtRef1.Text);
                ws.GetRow(new CellReference(rae[2]).Row).GetCell(new CellReference(rae[2]).Col).SetCellValue(Convert.ToInt32(txtRef2.Text));
                ws.GetRow(new CellReference(rae[3]).Row).GetCell(new CellReference(rae[3]).Col).SetCellValue(txtRef3.Text);
                ws.GetRow(new CellReference(rae[4]).Row).GetCell(new CellReference(rae[4]).Col).SetCellValue(txtRef4.Text);
                ws.GetRow(new CellReference(rae[5]).Row).GetCell(new CellReference(rae[5]).Col).SetCellValue(txtRef5.Text);
                ws.GetRow(new CellReference(rae[6]).Row).GetCell(new CellReference(rae[6]).Col).SetCellValue(txtRef6.Text);
                ws.GetRow(new CellReference(rae[7]).Row).GetCell(new CellReference(rae[7]).Col).SetCellValue(txtPropNome.Text);
                ws.GetRow(new CellReference(rae[8]).Row).GetCell(new CellReference(rae[8]).Col).SetCellValue(txtPropCPF.Text);
                ws.GetRow(new CellReference(rae[9]).Row).GetCell(new CellReference(rae[9]).Col).SetCellValue(txtPropDDD.Text);
                ws.GetRow(new CellReference(rae[10]).Row).GetCell(new CellReference(rae[10]).Col).SetCellValue(txtPropTelefone.Text);
                ws.GetRow(new CellReference(rae[11]).Row).GetCell(new CellReference(rae[11]).Col).SetCellValue(txtRTNome.Text);
                ws.GetRow(new CellReference(rae[12]).Row).GetCell(new CellReference(rae[12]).Col).SetCellValue(txtRTCauCrea.Text);
                ws.GetRow(new CellReference(rae[13]).Row).GetCell(new CellReference(rae[13]).Col).SetCellValue(txtRTUF.Text);
                ws.GetRow(new CellReference(rae[14]).Row).GetCell(new CellReference(rae[14]).Col).SetCellValue(txtRTCPF.Text);
                ws.GetRow(new CellReference(rae[15]).Row).GetCell(new CellReference(rae[15]).Col).SetCellValue(txtRTDDD.Text);
                ws.GetRow(new CellReference(rae[16]).Row).GetCell(new CellReference(rae[16]).Col).SetCellValue(txtRTTelefone.Text);
                ws.GetRow(new CellReference(rae[17]).Row).GetCell(new CellReference(rae[17]).Col).SetCellValue(txtIdEndereco.Text);
                ws.GetRow(new CellReference(rae[18]).Row).GetCell(new CellReference(rae[18]).Col).SetCellValue(txtIdComplemento.Text);
                ws.GetRow(new CellReference(rae[19]).Row).GetCell(new CellReference(rae[19]).Col).SetCellValue(txtIdBairro.Text);
                ws.GetRow(new CellReference(rae[20]).Row).GetCell(new CellReference(rae[20]).Col).SetCellValue(txtIdCEP.Text);
                ws.GetRow(new CellReference(rae[21]).Row).GetCell(new CellReference(rae[21]).Col).SetCellValue(txtIdMunicipio.Text);
                ws.GetRow(new CellReference(rae[22]).Row).GetCell(new CellReference(rae[22]).Col).SetCellValue(txtIdUF.Text);
                ws.GetRow(new CellReference(rae[23]).Row).GetCell(new CellReference(rae[23]).Col).SetCellValue(txtTerrenoValorProposto.Text);
                ws.GetRow(new CellReference(rae[24]).Row).GetCell(new CellReference(rae[24]).Col).SetCellValue(txtTerrenoMatricula.Text);
                ws.GetRow(new CellReference(rae[25]).Row).GetCell(new CellReference(rae[25]).Col).SetCellValue(txtTerrenoOficio.Text);
                ws.GetRow(new CellReference(rae[26]).Row).GetCell(new CellReference(rae[26]).Col).SetCellValue(txtTerrenoComarca.Text);
                ws.GetRow(new CellReference(rae[27]).Row).GetCell(new CellReference(rae[27]).Col).SetCellValue(txtTerrenoUF.Text);
                //ORÇAMENTO (PERCENTUAIS)
                ws.GetRow(new CellReference(rae[28]).Row).GetCell(new CellReference(rae[28]).Col).SetCellValue(Convert.ToDouble(txt1701.Text));
                ws.GetRow(new CellReference(rae[29]).Row).GetCell(new CellReference(rae[29]).Col).SetCellValue(Convert.ToDouble(txt1702.Text));
                ws.GetRow(new CellReference(rae[30]).Row).GetCell(new CellReference(rae[30]).Col).SetCellValue(Convert.ToDouble(txt1703.Text));
                ws.GetRow(new CellReference(rae[31]).Row).GetCell(new CellReference(rae[31]).Col).SetCellValue(Convert.ToDouble(txt1704.Text));
                ws.GetRow(new CellReference(rae[32]).Row).GetCell(new CellReference(rae[32]).Col).SetCellValue(Convert.ToDouble(txt1705.Text));
                ws.GetRow(new CellReference(rae[33]).Row).GetCell(new CellReference(rae[33]).Col).SetCellValue(Convert.ToDouble(txt1706.Text));
                ws.GetRow(new CellReference(rae[34]).Row).GetCell(new CellReference(rae[34]).Col).SetCellValue(Convert.ToDouble(txt1707.Text));
                ws.GetRow(new CellReference(rae[35]).Row).GetCell(new CellReference(rae[35]).Col).SetCellValue(Convert.ToDouble(txt1708.Text));
                ws.GetRow(new CellReference(rae[36]).Row).GetCell(new CellReference(rae[36]).Col).SetCellValue(Convert.ToDouble(txt1709.Text));
                ws.GetRow(new CellReference(rae[37]).Row).GetCell(new CellReference(rae[37]).Col).SetCellValue(Convert.ToDouble(txt1710.Text));
                ws.GetRow(new CellReference(rae[38]).Row).GetCell(new CellReference(rae[38]).Col).SetCellValue(Convert.ToDouble(txt1711.Text));
                ws.GetRow(new CellReference(rae[39]).Row).GetCell(new CellReference(rae[39]).Col).SetCellValue(Convert.ToDouble(txt1712.Text));
                ws.GetRow(new CellReference(rae[40]).Row).GetCell(new CellReference(rae[40]).Col).SetCellValue(Convert.ToDouble(txt1713.Text));
                ws.GetRow(new CellReference(rae[41]).Row).GetCell(new CellReference(rae[41]).Col).SetCellValue(Convert.ToDouble(txt1714.Text));
                ws.GetRow(new CellReference(rae[42]).Row).GetCell(new CellReference(rae[42]).Col).SetCellValue(Convert.ToDouble(txt1715.Text));
                ws.GetRow(new CellReference(rae[43]).Row).GetCell(new CellReference(rae[43]).Col).SetCellValue(Convert.ToDouble(txt1716.Text));
                ws.GetRow(new CellReference(rae[44]).Row).GetCell(new CellReference(rae[44]).Col).SetCellValue(Convert.ToDouble(txt1717.Text));
                ws.GetRow(new CellReference(rae[45]).Row).GetCell(new CellReference(rae[45]).Col).SetCellValue(Convert.ToDouble(txt1718.Text));
                ws.GetRow(new CellReference(rae[46]).Row).GetCell(new CellReference(rae[46]).Col).SetCellValue(Convert.ToDouble(txt1719.Text));
                ws.GetRow(new CellReference(rae[47]).Row).GetCell(new CellReference(rae[47]).Col).SetCellValue(Convert.ToDouble(txt1720.Text));
                if (txtMensuradoAcumulado.Text != "")
                    ws.GetRow(new CellReference(rae[48]).Row).GetCell(new CellReference(rae[48]).Col).SetCellValue(Convert.ToDouble(txtMensuradoAcumulado.Text));
                //DADOS ADICIONAIS
                if (txtContratoInicio.Text != "  /  /")
                    ws.GetRow(new CellReference(rae[49]).Row).GetCell(new CellReference(rae[49]).Col).SetCellValue(txtContratoInicio.Text);
                if (txtContratoTermino.Text != "  /  /")
                    ws.GetRow(new CellReference(rae[50]).Row).GetCell(new CellReference(rae[50]).Col).SetCellValue(txtContratoTermino.Text);
                if (_rae != "22/10/2021")
                    ws.GetRow(new CellReference(rae[51]).Row).GetCell(new CellReference(rae[51]).Col).SetCellValue(txtEtapaPrevista.Text);
                
                ////CRONOGRAMA
                ws.GetRow(new CellReference(rae[52]).Row).GetCell(new CellReference(rae[52]).Col).SetCellValue(Convert.ToDouble(txtExecutado.Text));
                ws.GetRow(new CellReference(rae[53]).Row).GetCell(new CellReference(rae[53]).Col).SetCellValue(Convert.ToDouble(txtParcela1.Text));
                ws.GetRow(new CellReference(rae[54]).Row).GetCell(new CellReference(rae[54]).Col).SetCellValue(Convert.ToDouble(txtParcela2.Text));
                ws.GetRow(new CellReference(rae[55]).Row).GetCell(new CellReference(rae[55]).Col).SetCellValue(Convert.ToDouble(txtParcela3.Text));
                ws.GetRow(new CellReference(rae[56]).Row).GetCell(new CellReference(rae[56]).Col).SetCellValue(Convert.ToDouble(txtParcela4.Text));
                ws.GetRow(new CellReference(rae[57]).Row).GetCell(new CellReference(rae[57]).Col).SetCellValue(Convert.ToDouble(txtParcela5.Text));
                ws.GetRow(new CellReference(rae[58]).Row).GetCell(new CellReference(rae[58]).Col).SetCellValue(Convert.ToDouble(txtParcela6.Text));
                ws.GetRow(new CellReference(rae[59]).Row).GetCell(new CellReference(rae[59]).Col).SetCellValue(Convert.ToDouble(txtParcela7.Text));
                ws.GetRow(new CellReference(rae[60]).Row).GetCell(new CellReference(rae[60]).Col).SetCellValue(Convert.ToDouble(txtParcela8.Text));
                ws.GetRow(new CellReference(rae[61]).Row).GetCell(new CellReference(rae[61]).Col).SetCellValue(Convert.ToDouble(txtParcela9.Text));
                ws.GetRow(new CellReference(rae[62]).Row).GetCell(new CellReference(rae[62]).Col).SetCellValue(Convert.ToDouble(txtParcela10.Text));
                ws.GetRow(new CellReference(rae[63]).Row).GetCell(new CellReference(rae[63]).Col).SetCellValue(Convert.ToDouble(txtParcela11.Text));
                ws.GetRow(new CellReference(rae[64]).Row).GetCell(new CellReference(rae[64]).Col).SetCellValue(Convert.ToDouble(txtParcela12.Text));
                ws.GetRow(new CellReference(rae[65]).Row).GetCell(new CellReference(rae[65]).Col).SetCellValue(Convert.ToDouble(txtParcela13.Text));
                ws.GetRow(new CellReference(rae[66]).Row).GetCell(new CellReference(rae[66]).Col).SetCellValue(Convert.ToDouble(txtParcela14.Text));
                ws.GetRow(new CellReference(rae[67]).Row).GetCell(new CellReference(rae[67]).Col).SetCellValue(Convert.ToDouble(txtParcela15.Text));
                ws.GetRow(new CellReference(rae[68]).Row).GetCell(new CellReference(rae[68]).Col).SetCellValue(Convert.ToDouble(txtParcela16.Text));
                ws.GetRow(new CellReference(rae[69]).Row).GetCell(new CellReference(rae[69]).Col).SetCellValue(Convert.ToDouble(txtParcela17.Text));
                ws.GetRow(new CellReference(rae[70]).Row).GetCell(new CellReference(rae[70]).Col).SetCellValue(Convert.ToDouble(txtParcela18.Text));
                ws.GetRow(new CellReference(rae[71]).Row).GetCell(new CellReference(rae[71]).Col).SetCellValue(Convert.ToDouble(txtParcela19.Text));
                ws.GetRow(new CellReference(rae[72]).Row).GetCell(new CellReference(rae[72]).Col).SetCellValue(Convert.ToDouble(txtParcela20.Text));
                ws.GetRow(new CellReference(rae[73]).Row).GetCell(new CellReference(rae[73]).Col).SetCellValue(Convert.ToDouble(txtParcela21.Text));
                ws.GetRow(new CellReference(rae[74]).Row).GetCell(new CellReference(rae[74]).Col).SetCellValue(Convert.ToDouble(txtParcela22.Text));
                ws.GetRow(new CellReference(rae[75]).Row).GetCell(new CellReference(rae[75]).Col).SetCellValue(Convert.ToDouble(txtParcela23.Text));
                ws.GetRow(new CellReference(rae[76]).Row).GetCell(new CellReference(rae[76]).Col).SetCellValue(Convert.ToDouble(txtParcela24.Text));
                ws.GetRow(new CellReference(rae[77]).Row).GetCell(new CellReference(rae[77]).Col).SetCellValue(Convert.ToDouble(txtParcela25.Text));
                ws.GetRow(new CellReference(rae[78]).Row).GetCell(new CellReference(rae[78]).Col).SetCellValue(Convert.ToDouble(txtParcela26.Text));
                ws.GetRow(new CellReference(rae[79]).Row).GetCell(new CellReference(rae[79]).Col).SetCellValue(Convert.ToDouble(txtParcela27.Text));
                ws.GetRow(new CellReference(rae[80]).Row).GetCell(new CellReference(rae[80]).Col).SetCellValue(Convert.ToDouble(txtParcela28.Text));
                ws.GetRow(new CellReference(rae[81]).Row).GetCell(new CellReference(rae[81]).Col).SetCellValue(Convert.ToDouble(txtParcela29.Text));
                ws.GetRow(new CellReference(rae[82]).Row).GetCell(new CellReference(rae[82]).Col).SetCellValue(Convert.ToDouble(txtParcela30.Text));



                using (FileStream arquivoRAE = new FileStream(caminho, FileMode.Create, FileAccess.Write))
                {
                    wbook.ForceFormulaRecalculation = true;
                    wbook.Write(arquivoRAE);
                }

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
                    //string line = File.ReadLines(openText.FileName).Skip(9).Take(1).First();
                    //string referencia = Util.CleanInput(line.Substring(31));

                    var convocacao = File.ReadAllLines(openText.FileName)
                           .Where(l => l.StartsWith("Refer"))
                           .Select(l => l.Substring(l.LastIndexOf("-") + 2))
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
                    txtRef6.Text = "01";

                    txtRef0.Focus();

                    pnlMainConvocacao.Show();
                    btnProximoConvocacao.Show();

                }
            }
            catch (Exception ex)
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
                        MessageBox.Show("Arquivo incompatível ou versão não suportada!\r\n\r\n" + version, "Planilha PFUI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    else
                    {
                        string[] aeX = PFUI.SetArray(version);
                        ImportPFUI(aeX, sheet);

                        lblVersion.Text = version;
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
                    
                    
                    FileStream arquivoXLS = new FileStream(openExcel.FileName, FileMode.Open, FileAccess.Read);
                    HSSFWorkbook wbook = new HSSFWorkbook(arquivoXLS);
                    ISheet sheet = wbook.GetSheet("RAE");
                    _rae = (sheet.Footer.Left).Substring(10);

                    txtLogFinalizar.Text += "Vigência da planilha: " + _rae + "\r\n";

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
                        ExportRAE(saveExcel.FileName);
                        txtLogFinalizar.Text += "\r\n--------------------------------\r\n\r\nConcluído!";
                        btnNew.Show();
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

        private void label101_Click(object sender, EventArgs e)
        {

        }
    }
}
