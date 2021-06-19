using System;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;
using Raecef.Ent;
using System.Globalization;
using Raecef;

namespace Raecef.Model
{
    class PfuiModel
    {
        public PFUI GetPFUI(string filePatch)
        {
            var xls = new XLWorkbook(filePatch);
            var planilha = xls.Worksheets.First(w => w.Name == "Proposta");

            PFUI pfui = new PFUI();

            //Cabeçalho da PFUI
            {
                pfui.Proponente_Nome = planilha.Cell("G42").RichText.ToString().ToUpper();
                pfui.Proponente_CPF = Util.FormatedCPF(planilha.Cell("AA42").RichText.ToString());
                pfui.Proponente_DDD = planilha.Cell("AG42").RichText.ToString();
                pfui.Proponente_Telefone = Util.FormatedFone(planilha.Cell("AI42").RichText.ToString());
                pfui.RT_Nome = planilha.Cell("AN42").RichText.ToString().ToUpper();
                pfui.RT_CauCrea = planilha.Cell("AX42").RichText.ToString().Replace(',', '.');
                pfui.RT_UF = planilha.Cell("BD42").RichText.ToString();
                pfui.RT_CPF = Util.FormatedCPF(planilha.Cell("BF42").Value.ToString());
                pfui.RT_DDD = planilha.Cell("BL42").RichText.ToString();
                pfui.RT_Telefone = Util.FormatedFone(planilha.Cell("BN42").RichText.ToString());
                pfui.Obra_Endereco = planilha.Cell("G46").RichText.ToString().ToUpper();
                pfui.Obra_Complemento = planilha.Cell("AP46").RichText.ToString().ToUpper();
                pfui.Obra_CEP = Util.FormatedCEP(planilha.Cell("BA46").RichText.ToString());
                pfui.Obra_Bairro = planilha.Cell("BF46").RichText.ToString().ToUpper();
                pfui.Obra_Municipio = planilha.Cell("G48").RichText.ToString().ToUpper();
                pfui.Obra_UF = planilha.Cell("AB48").RichText.ToString();
                pfui.Terreno_Valor = string.Format("{0:0,0.00}", Convert.ToDouble(planilha.Cell("AG50").CachedValue));
                pfui.Matricula_Numero = planilha.Cell("AP50").RichText.ToString().Replace(',', '.');
                pfui.Matricula_Oficio = planilha.Cell("AX50").RichText.ToString().ToUpper();
                pfui.Matricula_Comarca = planilha.Cell("BF50").RichText.ToString().ToUpper();
                pfui.Matricula_UF = planilha.Cell("BN50").RichText.ToString();
            }

            //Itens do orçamento (percentuais)
            {
                pfui.Item_17_01 = planilha.Cell("AR116").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_02 = planilha.Cell("AR118").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_03 = planilha.Cell("AR130").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_04 = planilha.Cell("AR137").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_05 = planilha.Cell("AR146").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_06 = planilha.Cell("AR156").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_07 = planilha.Cell("AR165").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_08 = planilha.Cell("AR172").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_09 = planilha.Cell("AR179").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_10 = planilha.Cell("AR190").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_11 = planilha.Cell("AR197").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_12 = planilha.Cell("AR207").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_13 = planilha.Cell("AR217").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_14 = planilha.Cell("AR228").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_15 = planilha.Cell("AR234").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_16 = planilha.Cell("AR245").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_17 = planilha.Cell("AR254").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_18 = planilha.Cell("AR262").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_19 = planilha.Cell("AR271").CachedValue.ToString().Replace('.', ',');
                pfui.Item_17_20 = planilha.Cell("AR273").CachedValue.ToString().Replace('.', ',');
            }

            //Cronograma
            //---Executada, Parcela 1 a Parcela 8
            {
                string linha18_50_01;

                if (planilha.Cell("G310").CachedValue.ToString() == "18.50.01")
                    linha18_50_01 = "310";
                else if (planilha.Cell("G311").CachedValue.ToString() == "18.50.01")
                    linha18_50_01 = "311";
                else
                    linha18_50_01 = "";


                pfui.Executado = planilha.Cell("AJ" + (Convert.ToInt32(linha18_50_01) + 1).ToString()).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_1 = planilha.Cell("AL" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_2 = planilha.Cell("AP" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_3 = planilha.Cell("AT" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_4 = planilha.Cell("AX" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_5 = planilha.Cell("BB" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_6 = planilha.Cell("BF" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_7 = planilha.Cell("BJ" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_8 = planilha.Cell("BN" + linha18_50_01).CachedValue.ToString().Replace('.', ',');
            }

            //---Parcela 9 a Parcela 16
            {
                string linha18_50_02;

                if (planilha.Cell("G350").CachedValue.ToString() == "18.50.02")
                    linha18_50_02 = "350";
                else if (planilha.Cell("G351").CachedValue.ToString() == "18.50.02")
                    linha18_50_02 = "351";
                else
                    linha18_50_02 = "";

                pfui.Parcela_9 = planilha.Cell("AL" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_10 = planilha.Cell("AP" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_11 = planilha.Cell("AT" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_12 = planilha.Cell("AX" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_13 = planilha.Cell("BB" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_14 = planilha.Cell("BF" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_15 = planilha.Cell("BJ" + linha18_50_02).CachedValue.ToString().Replace('.', ',');
                pfui.Parcela_16 = planilha.Cell("BN" + linha18_50_02).CachedValue.ToString().Replace('.', ',');



            }




            return pfui;
        }


    }
}
