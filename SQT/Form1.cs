using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace SQT
{
    public partial class Form1 : Form
    {
        #region VARS
        public bool sucessfulSave = false;
        public string quoteNumber = "";
        public string exchangeRateDate;
        public string exchangeRateText;
        public float applicableExchangeRate = 1;
        public float freightTotal = 0;
        public float liftPrice;
        public float lowestMargin;
        public int exCurrency = 0; // 0 AUD, 1 USD, 2 EUR
        public int num20Ft;
        public int num40Ft;
        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public Dictionary<string, string> wordExportData = new Dictionary<string, string>();
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();
        public Dictionary<string, string> priceExports = new Dictionary<string, string>();
        Word.Application fileOpen;
        Word.Document document;
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //this.Enabled = true;
            FetchBasePrices();
            FetchCurrencyRates();
            FetchLabourPrices();
            GeneratePriceList();

            lowestMargin = basePrices["17LowestMargin"];
            tbMargin.Text = basePrices["18DefaultMargin"].ToString();
            lblCostOfParts.Text = "$0";
            lblCostIncludingMargin.Text = "$0";
            lblGST.Text = "$0";
            lblPriceIncludingGST.Text = "$0";
            quoteNumber = ("Qu" + DateTime.Now.ToString("yy") + "-000");
            tBQuoteNumber.Text = quoteNumber;
            lbWait.Visible = false;
            button3.Visible = false;
            button3.Enabled = false;
            printButton.Visible = false;
            printButton.Enabled = false;
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
                exchangeRateText = "AUD";
            }
            else if (selector == "U")
            {
                applicableExchangeRate = exchangeRates["USD"];
                exchangeRateLbl.Visible = true;
                exchangeRateLbl.Enabled = true;
                lblExchangeDate.Enabled = true;
                lblExchangeDate.Visible = true;
                exchangeRateLbl.Text = "The current exchange rate is $1 USD to " + PriceRounding(exchangeRates["USD"]) + " AUD";
                lblExchangeDate.Text = "Correct as of " + exchangeRateDate;
                exCurrency = 1;
                exchangeRateText = "USD";
            }
            else if (selector == "E")
            {
                applicableExchangeRate = exchangeRates["EUR"];
                exchangeRateLbl.Visible = true;
                exchangeRateLbl.Enabled = true;
                lblExchangeDate.Enabled = true;
                lblExchangeDate.Visible = true;
                exchangeRateLbl.Text = "The current exchange rate is €1 EUR to " + PriceRounding(exchangeRates["EUR"]) + " AUD";
                lblExchangeDate.Text = "Correct as of " + exchangeRateDate;
                exCurrency = 2;
                exchangeRateText = "EUR";
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
                PriceListFormatting(lblLiftNoConvertPrice, float.Parse(tbCost.Text), true);
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

        private string PriceRounding(float s, bool b = false)
        {
            if (b)
            {
                return "€" + Math.Round(s, 2).ToString("N", new System.Globalization.CultureInfo("en-US"));
            }
            else
            {
                return "$" + Math.Round(s, 2).ToString("N", new System.Globalization.CultureInfo("en-US"));
            }
        }

        public void PriceListFormatting(Label label, float cost, bool b = false)
        {           
            if (cost > 0)
            {
                label.Text = PriceRounding(cost, b);
                liftPrice += cost;
            }
            else
            {
                label.Text = "N/A";
                liftPrice += 0;
            }
        }

        private void button1_Click(object sender, EventArgs e) // generate price list button
        {
            lblWaitControl(true);
            if (floorsTbChecker() && marginTbChecker())
            {
                GeneratePriceList();
                button3.Visible = true;
                button3.Enabled = true;
                printButton.Visible = true;
                printButton.Enabled = true;
            }
            lblWaitControl(false);
        }

        //private void button2_Click(object sender, EventArgs e) // export to quote button
        //{
        //    QuoteInfo qI = new QuoteInfo();
        //    qI.Show();
        //}

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
            if (sucessfulSave)
            {
                WordData("AE101", tBAddress.Text); //address
                WordData("AE102", tBQuoteNumber.Text);//quote number
                WordData("AE103", tbNumberLifts.Text);//number of lifts
                WordData("AE104", tBFloors.Text);//number of floors

                QuoteInfo2 qI = new QuoteInfo2();
                qI.Show();//open questionaire 
                //questions complete method called from final form of querstions to continue the export to word function. 
            }
            else
            {
                MessageBox.Show("Saving Error, Document not saved");
                return;
            }
        }

        public void QuestionsComplete() //called from the final question to continue the export to word function 
        {
            WordData("AE211", FormalDate());
            WordData("AE212", lblCostIncludingMargin.Text);
            WordData("AE213", lblGST.Text);
            WordData("AE214", lblPriceIncludingGST.Text);

            WordReplaceLooper(wordExportData);// loop the find and replace method to populate the info 
            WordSave(true);// save the doc again 
            WordFinish();//finish the methods 

        }

        public void WordSetup() // sets up the word document ready to be written
        {
            lblWaitControl(true);
            fileOpen = new Word.Application();
            document = fileOpen.Documents.Open("X:\\Program Dependancies\\Quote tool\\SQT.docm", ReadOnly: false);
            fileOpen.Visible = false;
            document.Activate();
        }

        private void WordFinish() // closes the word document 
        {
            fileOpen.Quit();
            lblWaitControl(false);
        }

        private void WordSave(bool b) // if false,asks where to save the word doc and saves it. if true saves in the previously set location
        {
            if (!b)
            {
                saveFileDialog1.Title = ("Where to save the quote");
                saveFileDialog1.InitialDirectory = "X:\\Sales\\Qu-" + DateTime.Now.ToString("yyyy");
                saveFileDialog1.FileName = tBAddress.Text + " Quote";
                saveFileDialog1.DefaultExt = "docm";
                saveFileDialog1.Filter = "Word Docs (*.docm; *.docx) |*.docm;*.docx|All files (*.*) |*.*";
                //saveFileDialog1.ShowDialog();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    document.SaveAs2(saveFileDialog1.FileName);
                    sucessfulSave = true;
                }
            }
            else if (b)
            {
                document.SaveAs2(saveFileDialog1.FileName);
            }
        }

        public void WordData(string k, string v) // called from question forms to take data and write it to the dictionary
        {
            wordExportData.Add(k, v);
            //MessageBox.Show("Word Data Method called with: " + k + " " + v);
        }

        private void WordReplaceLooper(Dictionary<string, string> d) // loops through the word document performing a find and replace operation
        {
            foreach (KeyValuePair<string, string> i in d)
            {
                FindAndReplace(fileOpen, i.Key, i.Value);
            }
        }

        static void FindAndReplace(Word.Application fileOpen, object findText, object replaceWithText)
        {
            object matchCase = true;
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

        public void QuestionCloseCall(Form f) // method called by a closing Question form to signal the program to close the word document being used and to wipe the data from the dictionary
        {
            //Form[] questionForms = new Form[11];
            DialogResult d = MessageBox.Show("Are you sure you wish to cancel quote exporting?", "Close?", MessageBoxButtons.YesNo);

            if (d == DialogResult.Yes)
            {
                wordExportData.Clear();
                WordFinish();
                MessageBox.Show("Word Export Canceled");
                f.Close();
            }
        }

        public string RadioButtonHandeler(TextBox tb = null, params RadioButton[] rb) // called with all the radio buttons in each group to find which is checked and return the text 
        {
            foreach (RadioButton i in rb)
            {
                if (i.Checked) //find the checked box
                {
                    if (i.Text == "")// if radio button has no text return the text of the textbox
                    {
                        return tb.Text;
                    }
                    else// otherwise retun the text of the radio button
                    {
                        return i.Text;
                    }
                }
            }
            return "";
        }

        public string CheckBoxHandler(params CheckBox[] cB)
        {
            string str = "";
            int count = 0;

            foreach (CheckBox i in cB)
            {
                if (i.Checked)
                {
                    count++;
                }
            }

            foreach (CheckBox i in cB)
            {
                if (i.Checked)
                {
                    str += i.Text;

                    if (count > 0)
                    {
                        if (count >= 3)
                        {
                            str += ", ";
                        }
                        else if (count == 2)
                        {
                            str += " & ";
                        }
                        count--;
                    }
                }
            }

            return str;
        }

        public string MeasureStringChecker(string text, string measurementSuffix) //called to check if the text has the correct measurment suffix before exporting to quote document
        {
            //check if text ends in the correct suffix
            if (text.EndsWith(measurementSuffix))
            {
                //if so return the text as presented
                return text;
            }
            else
            {
                //if not, append a space and add the measurement suffix
                string s = text + " " + measurementSuffix;

                //then return the new string
                return s;
            }
        }

        public string FormalDate()
        {
            string day = DateTime.Now.ToString("%d");
            string monthYear = DateTime.Now.ToString("Y");

            bool singleDigit = day.Length == 1;
            bool endIn1 = day.EndsWith("1");
            bool endIn2 = day.EndsWith("2");
            bool endIn3 = day.EndsWith("3");
            bool startWith1 = day.StartsWith("1");

            if (endIn1)
            {
                if (singleDigit)
                {
                    day += "st";
                }
                else if (startWith1 && !singleDigit)
                {
                    day += "th";
                }
                else
                {
                    day += "st";
                }
            }
            else if (endIn2)
            {
                if (singleDigit)
                {
                    day += "nd";
                }
                else if (startWith1 && !singleDigit)
                {
                    day += "th";
                }
                else
                {
                    day += "nd";
                }
            }
            else if (endIn3)
            {
                if (singleDigit)
                {
                    day += "rd";
                }
                else if (startWith1 && !singleDigit)
                {
                    day += "th";
                }
                else
                {
                    day += "rd";
                }
            }
            else
            {
                day += "th";
            }

            string date = day + " " + monthYear;
            return date;
        }

        private void button2_Click_1(object sender, EventArgs e) // close button
        {
            // lines below this till "return" are used for the close button to function as a generic debug button for testing. 
            //close method works and requires no further testing at this time
            //MessageBox.Show(FormalDate());

            // return; // remove this line and above to have the close button function normally

            if (document != null)
            {
                document.Close();
            }
            this.Close();
        }

        #region unused methods
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
        #endregion

        private void printButton_Click(object sender, EventArgs e)
        {
            lblWaitControl(true);
            printButton.BackColor = Color.Blue;
            if (SavePricesDocument())
            {
                MessageBox.Show("Prices exported as " + saveFileDialog1.FileName);
                printButton.BackColor = Color.Green;
            }
            else
            {
                MessageBox.Show("Price Saving Failed");
                printButton.BackColor = Color.Red;
            }
            lblWaitControl(false);
        }

        private bool SavePricesDocument()
        {
            saveFileDialog1.Title = ("Where to save the prices");
            saveFileDialog1.InitialDirectory = "X:\\Sales\\Qu-" + DateTime.Now.ToString("yyyy");
            saveFileDialog1.FileName = tBAddress.Text + " Price Breakdown";
            saveFileDialog1.DefaultExt = "docx";
            saveFileDialog1.Filter = "Word Doc (*.docx) |*.docx| All files (*.*) |*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileOpen = new Word.Application();
                document = fileOpen.Documents.Open("X:\\Program Dependancies\\Quote tool\\PriceExport.docx", ReadOnly: false);
                SavePricesToDict();
                fileOpen.Visible = false;
                document.Activate();
                WordReplaceLooper(priceExports);
                document.SaveAs2(saveFileDialog1.FileName);
                document.Close();
                return true;
            }
            else
            {
                return false;
            }
        }

        public void SavePricesToDict()
        {
            priceExports.Clear();
            priceExports.Add("AEP1", tBAddress.Text);
            priceExports.Add("AEP2", tBQuoteNumber.Text);
            priceExports.Add("AEP3", FormalDate());
            priceExports.Add("AEP4", exchangeRateText);
            priceExports.Add("AEP5", lblLiftNoConvertPrice.Text);
            priceExports.Add("AEP6", lblCost.Text);
            priceExports.Add("AEP7", lblFinishes.Text);
            priceExports.Add("AEP8", lblFire.Text);
            priceExports.Add("AEP9", lblGSM.Text);
            priceExports.Add("AEP10", lblBlanket.Text);
            priceExports.Add("AEP11", lblSump.Text);
            priceExports.Add("AEP12", lblSundries.Text);
            priceExports.Add("AEP13", lblWiring.Text);
            priceExports.Add("AEP14", lblSign.Text);
            priceExports.Add("AEP15", lblShaft.Text);
            priceExports.Add("AEP16", lblDuct.Text);
            priceExports.Add("AEP17", lblElectrical.Text);
            priceExports.Add("AEP18", lblAccommodation.Text);
            priceExports.Add("AEP19", lblCartage.Text);
            priceExports.Add("AEP20", lblDrawing.Text);
            priceExports.Add("AEP21", lblFork.Text);
            priceExports.Add("AEP22", lblMaintenance.Text);
            priceExports.Add("AEP23", lblManuals.Text);
            priceExports.Add("AEP24", lblStorage.Text);
            priceExports.Add("AEP25", lblTravel.Text);
            priceExports.Add("AEP26", lblWorkcover.Text);
            priceExports.Add("AEP27", lblScaffold.Text);
            priceExports.Add("AEP28", lblEntrance.Text);
            priceExports.Add("AEP29", lblSecurity.Text);
            priceExports.Add("AEP30", lblFreight.Text);
            priceExports.Add("AEP31", lblLabour.Text);
            priceExports.Add("AEP32", lblCostOfParts.Text);
            priceExports.Add("AEP33", tbMargin.Text + "%");
            priceExports.Add("AEP34", lblCostIncludingMargin.Text);
            priceExports.Add("AEP35", lblGST.Text);
            priceExports.Add("AEP36", lblPriceIncludingGST.Text);
        }

        //private void tBFloors_TextChanged(object sender, EventArgs e)
        //{
        //    floorsTbChecker();
        //}

        public bool floorsTbChecker()
        {
            int i = 0;
            try
            {
                i = int.Parse(tBFloors.Text);
            }
            catch
            {
                //MessageBox.Show("Invalid floor number entered ");
                return false;  
            }

            if (i > 16 || i < 0)
            {
                MessageBox.Show("Invalid floor number entered ");
                return false;
            }
            else
            {
                return true;
            }
        }

        //private void tbMargin_Leave(object sender, EventArgs e)
        //{
        //    marginTbChecker();
        //}

        public bool marginTbChecker()
        {
            try
            {
                if (lowestMargin > float.Parse(tbMargin.Text))
                {
                    MessageBox.Show("Margin % is below the allowed minimum of " + lowestMargin + "%");
                    return false;
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        public void lblWaitControl(bool b)
        {
            lbWait.Enabled = b;
            lbWait.Visible = b;
            this.Enabled = !b;
        }
    }
}