using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace SQT
{
    public partial class Form1 : Form
    {
        // VARS
        public string quoteNumber = "";
        public float applicableExchangeRate = 1;
        public int exCurrency = 0; // 0 AUD, 1 USD, 2 EUR
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();
        public string exchangeRateDate;
        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public int num20Ft;
        public int num40Ft;
        public float freightTotal = 0;
        public float liftPrice;
        Word.Application fileOpen;
        Word.Document document;
        public Dictionary<string, string> wordExportData = new Dictionary<string, string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Enabled = true;
            FetchBasePrices();
            FetchCurrencyRates();
            FetchLabourPrices();
            GeneratePriceList();
            lblCostOfParts.Text = "$0";
            lblCostIncludingMargin.Text = "$0";
            lblGST.Text = "$0";
            lblPriceIncludingGST.Text = "$0";
            quoteNumber = ("Qu" + DateTime.Now.ToString("yy") + "-000");
            tBQuoteNumber.Text = quoteNumber;
            lbWait.Visible = false;
        }

        private void tBQuoteNumber_TextChanged(object sender, EventArgs e)
        {
            //when the quote number text box is changed, if it is different to the existing quote number update the quote number
            if (quoteNumber != tBQuoteNumber.Text)
            {
                quoteNumber = tBQuoteNumber.Text;
            }
        }

        //find the live currency rates from floatrates.com and write them into the Exchange rate dictionary 
        private void FetchCurrencyRates()
        {
            string dKey = "";
            float dName = 0;
            bool b = true;

            XmlTextReader XMLR = new XmlTextReader("http://www.floatrates.com/daily/aud.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "targetCurrency")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "inverseRate")
                {
                    dName = float.Parse(XMLR.ReadElementContentAsString());
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "pubDate" && b)
                {
                    exchangeRateDate = XMLR.ReadElementContentAsString();
                    b = false;
                }

                if (dKey != "" && dName != 0)
                {
                    exchangeRates.Add(dKey, dName * basePrices["16CurrencyMargin"]);
                    dKey = "";
                    dName = 0;
                }
            }

            XMLR.Close();
        }

        //find the quote price list list from the file in the server and write them into the base prices dictionary 
        private void FetchBasePrices()
        {
            string dKey = "";
            float dName = -1;

            XmlTextReader XMLR = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\QuotePriceList.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "costItem")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "price")
                {
                    dName = float.Parse(XMLR.ReadElementContentAsString());
                }

                if (dKey != "" && dName != -1)
                {
                    basePrices.Add(dKey, dName);
                    dKey = "";
                    dName = -1;
                }
            }

            XMLR.Close();
        }

        //find the labour costs from the file int he server and write them into the labour prices dictionary 
        private void FetchLabourPrices()
        {
            int dKey = 0;
            float dName = -1;

            XmlTextReader XMLR = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\LabourCosts.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Floors")
                {
                    dKey = int.Parse(XMLR.ReadElementContentAsString());
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Price")
                {
                    dName = float.Parse(XMLR.ReadElementContentAsString());
                }

                if (dKey != 0 && dName != -1)
                {
                    labourPrice.Add(dKey, dName);
                    dKey = 0;
                    dName = -1;
                }
            }

            XMLR.Close();
        }

        //set the exchange rate from the dictionary and update texts based on the clicked button.

        private void buttonAUD_Click(object sender, EventArgs e)
        {
            SelectCurrency("A");
        }
        private void buttonUSD_Click(object sender, EventArgs e)
        {
            SelectCurrency("U");
        }
        private void buttonEUR_Click(object sender, EventArgs e)
        {
            SelectCurrency("E");
        }

        public void SelectCurrency(string selector)
        {
            currencySelectionGroup.Enabled = false;
            currencySelectionGroup.Visible = false;

            if (selector == "A")
            {
                applicableExchangeRate = 1;
                exCurrency = 0;
            }
            else if (selector == "U")
            {
                applicableExchangeRate = exchangeRates["USD"];
                exchangeRateLbl.Visible = true;
                exchangeRateLbl.Enabled = true;
                lblExchangeDate.Enabled = true;
                lblExchangeDate.Visible = true;
                exchangeRateLbl.Text = "The current Exchange rate is $1 USD to " + PriceRounding(exchangeRates["USD"]) + " AUD";
                lblExchangeDate.Text = "Correct as of " + exchangeRateDate;
                exCurrency = 1;
            }
            else if (selector == "E")
            {
                applicableExchangeRate = exchangeRates["EUR"];
                exchangeRateLbl.Visible = true;
                exchangeRateLbl.Enabled = true;
                lblExchangeDate.Enabled = true;
                lblExchangeDate.Visible = true;
                exchangeRateLbl.Text = "The current Exchange rate is €1 EUR to " + PriceRounding(exchangeRates["EUR"]) + " AUD";
                lblExchangeDate.Text = "Correct as of " + exchangeRateDate;
                exCurrency = 2;
            }
        }

        private void btnShippingReset_Click(object sender, EventArgs e)
        {
            ShippingCalculation(1);
        }
        private void btn20Ft_Click(object sender, EventArgs e)
        {
            ShippingCalculation(2);
        }
        private void btn40Ft_Click(object sender, EventArgs e)
        {
            ShippingCalculation(3);
        }

        public void ShippingCalculation(int selector)
        {
            if (selector == 1)
            {
                num20Ft = 0;
                shippingLbl20.Text = num20Ft + "x 20ft Container(s) - $0";
                num40Ft = 0;
                shippingLbl40.Text = num40Ft + "x 40ft Container(s) - $0";
                shippingLblTotal.Text = "Total of $0 for shipping";
            }
            else if (selector == 2)
            {
                num20Ft++;
                shippingLbl20.Text = num20Ft + "x 20ft Container(s) - " + PriceRounding(basePrices["20ftFreight"] * num20Ft);
            }
            else if (selector == 3)
            {
                num40Ft++;
                shippingLbl40.Text = num40Ft + "x 40ft Container(s) - " + PriceRounding(basePrices["40ftFreight"] * num40Ft);
            }
            freightTotal = (num20Ft * basePrices["20ftFreight"]) + (num40Ft * basePrices["40ftFreight"]);
            shippingLblTotal.Text = "Total of " + PriceRounding(freightTotal) + " for shipping";
        }

        public void GeneratePriceList()
        {
            liftPrice = 0;

            //prices pulled from text boxes of program
            PriceListFormatting(lblCost, float.Parse(tbCost.Text) * applicableExchangeRate);
            PriceListFormatting(lblSundries, float.Parse(tbSundries.Text));
            PriceListFormatting(lblShaft, float.Parse(tbShaftLight.Text));
            PriceListFormatting(lblDuct, float.Parse(tbDuct.Text));
            PriceListFormatting(lblAccommodation, float.Parse(tbAccomodation.Text));
            PriceListFormatting(lblStorage, float.Parse(tbStorage.Text));
            PriceListFormatting(lblTravel, float.Parse(tbTravel.Text));
            //prices pulled from base prices dictionary
            PriceListFormatting(lblFinishes, basePrices["1CarFinishes"]);
            PriceListFormatting(lblFire, basePrices["2FireExtinguisher"]);
            PriceListFormatting(lblGSM, basePrices["3GSMUnitPhone"]);
            PriceListFormatting(lblBlanket, basePrices["4ProtectiveBlanket"]);
            PriceListFormatting(lblSump, basePrices["5SumpCover"]);
            PriceListFormatting(lblWiring, basePrices["6Wiring"]);
            PriceListFormatting(lblSign, basePrices["7Sinage"]);
            PriceListFormatting(lblElectrical, basePrices["8ElectricalBox"]);
            PriceListFormatting(lblCartage, float.Parse(tbCartage.Text));
            PriceListFormatting(lblDrawing, basePrices["9Drawings"]);
            PriceListFormatting(lblFork, basePrices["10ForkLift"]);
            PriceListFormatting(lblMaintenance, basePrices["11Maintenance"]);
            PriceListFormatting(lblManuals, basePrices["12Manuals"]);
            PriceListFormatting(lblWorkcover, basePrices["13WorkcoverFees"]);
            PriceListFormatting(lblScaffold, float.Parse(tbScaffold.Text) * basePrices["14Scaffolds"]);
            PriceListFormatting(lblEntrance, float.Parse(tbEntranceGuards.Text) * float.Parse(tbWeeksRequired.Text) * basePrices["15EntranceGuards"]);
            //add security from base prices dictionary if box is checked
            if (cbSecurity.Checked)
            {
                PriceListFormatting(lblSecurity, basePrices["Security"] + (basePrices["SecurityPerFloor"] * int.Parse(tBFloors.Text)));
            }
            else
            {
                PriceListFormatting(lblSecurity, 0);
            }
            // if needed show the unconverted cost of the lift
            if (exCurrency == 0)
            {
                lblLiftNoConvert.Visible = false;
                lblLiftNoConvertPrice.Visible = false;  
            }
            else if (exCurrency == 1)
            {
                lblLiftNoConvert.Visible = true;
                lblLiftNoConvertPrice.Visible = true;
                lblLiftNoConvert.Text = "Cost of Lift (USD)";
                PriceListFormatting(lblLiftNoConvertPrice, float.Parse(tbCost.Text));
            }
            else if (exCurrency == 2)
            {
                lblLiftNoConvert.Visible = true;
                lblLiftNoConvertPrice.Visible = true;
                lblLiftNoConvert.Text = "Cost of Lift (EUR)";
                PriceListFormatting(lblLiftNoConvertPrice, float.Parse(tbCost.Text));
            }

            //add freight based on number of required containers 
            PriceListFormatting(lblFreight, freightTotal);
            //add labour from the labour costs dictionary based on number of floors in the building 
            PriceListFormatting(lblLabour, labourPrice[int.Parse(tBFloors.Text)]);

            //lblCostOfParts.Text = "$" + Math.Round(liftPrice, 2).ToString("0.00");
            //lblCostIncludingMargin.Text = "$" + Math.Round(liftPrice * marginPercent, 2).ToString("0.00");
            // lblGST.Text = "$" + Math.Round(liftPrice * marginPercent * 0.1, 2).ToString("0.00");
            //  lblPriceIncludingGST.Text = "$" + Math.Round(liftPrice * marginPercent * 1.1, 2).ToString("0.00");

            float marginPercent = 1 + (float.Parse(tbMargin.Text) / 100);
            liftPrice *= int.Parse(tbNumberLifts.Text);

            lblCostOfParts.Text = PriceRounding(liftPrice);
            lblCostIncludingMargin.Text = PriceRounding(liftPrice * marginPercent);
            lblGST.Text = PriceRounding(liftPrice * marginPercent * 0.1f);
            lblPriceIncludingGST.Text = PriceRounding(liftPrice * marginPercent * 1.1f);
        }

        private string PriceRounding(float s)
        {
            return "$" + Math.Round(s, 2).ToString("0.00");
        }

        public void PriceListFormatting(Label label, float cost)
        {
            if (cost > 0)
            {
                label.Text = PriceRounding(cost);
                liftPrice += cost;
            }
            else
            {
                label.Text = "N/A";
                liftPrice += 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneratePriceList();
        }

        private void label13_Click(object sender, EventArgs e)
        {
            //
        }
        private void costOfLiftTB_TextChanged(object sender, EventArgs e)
        {
            //
        }
        private void labelLiftCurrency_Click(object sender, EventArgs e)
        {
            //
        }
        private void label35_Click(object sender, EventArgs e)
        {
            //
        }

        private void button2_Click(object sender, EventArgs e)
        {
            QuoteInfo qI = new QuoteInfo();
            qI.Show();
        }

        //public string PullTextData(Label l, TextBox t)
        //{
        //    if (l != null)
        //    {
        //        return l.Text;
        //    }
        //    else if (t != null)
        //    {
        //        return t.Text;
        //    }
        //    else
        //    {
        //        return null;
        //    }
        //}

        private void button3_Click(object sender, EventArgs e)
        {
            WordSetup();//find and set vars to the quote template document 
            WordSave(false); // save the doc

            QuoteInfo qI = new QuoteInfo();
            qI.Show();//open questionaire 
            this.Enabled = false;
            //questions complete method called from final form of querstions to continue the export to word function. 
        }
        public void QuestionsComplete()
        {
            //called from the final question to continue the export to word function 
            WordReplaceLooper();// loop the find and replace method to populate the info 
            WordSave(true);// save the doc again 
            WordFinish();//finish the methods 
        }

        public void WordSetup()
        {
            lbWait.Visible = true;
                        fileOpen = new Word.Application();
            document = fileOpen.Documents.Open("X:\\Program Dependancies\\Quote tool\\SQT.docm", ReadOnly: false);
            fileOpen.Visible = false;
            document.Activate();
        }

        private void WordFinish()
        {
            fileOpen.Quit();
            lbWait.Visible = false;
            this.Enabled = true;
        }

        private void WordSave(bool b)
        {
            if (!b)
            {
                saveFileDialog1.Title = ("Where to save the quote");
                saveFileDialog1.InitialDirectory = "X:\\Sales\\Qu-" + DateTime.Now.ToString("yyyy");
                saveFileDialog1.FileName = tBAddress.Text + "Quote";
                saveFileDialog1.DefaultExt = "docm";
                saveFileDialog1.ShowDialog();
                if (saveFileDialog1.FileName != null || saveFileDialog1.FileName != "")
                {
                    document.SaveAs2(saveFileDialog1.FileName);
                }
                else
                {
                    MessageBox.Show("Saving Error, Document not saved");
                }
            }
            else if (b)
            {
                document.SaveAs2(saveFileDialog1.FileName);
            }
        }

        public void WordData(string k, string v)
        {
            wordExportData.Add(k, v);
        }

        private void WordReplaceLooper()
        {
            foreach (KeyValuePair<string, string> i in wordExportData)
            {
                FindAndReplace(fileOpen, i.Key, i.Value);
            }
        }

        static void FindAndReplace(Word.Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            //execute find and replace
            fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        public void QuestionCloseCall(Form f)
        {
            Form[] questionForms = new Form[11];

            wordExportData.Clear();
            WordFinish();
            MessageBox.Show("Word Export Canceled");
            f.Close();
        }
    }
}