using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace SQT
{
    public partial class Pin_Dif_Calc : Form
    {
        #region VARS
        public bool debugMode = false; // bool switch to show all debug message boxes in the program, turn off before release

        public bool sucessfulSave = false;
        public bool loadingPreviousData = false;
        private bool costInEuro;

        public string salesRep = "Pintaric";
        public string quoteNumber = "";
        public string exchangeRateDate;
        public string exchangeRateText;

        public float applicableExchangeRate = 1;
        public float freightTotal = 0;
        public float liftPrice;
        public float lowestMargin;
        private float marginPercent;

        public int exCurrency = 0; // 0 AUD, 1 USD, 2 EUR
        public int num20Ft;
        public int num40Ft;
        public int maxFloorNumber;
        public int numberOfPagesNeeded;

        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public Dictionary<string, string> wordExportData = new Dictionary<string, string>();
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();
        public Dictionary<string, string> priceExports = new Dictionary<string, string>();
        public Dictionary<string, string> saveData = new Dictionary<string, string>();

        Word.Application fileOpen;
        Word.Document document;
        #endregion

        #region Form Loading Methods
        public Pin_Dif_Calc()
        {
            InitializeComponent();
        }

        private void Pin_Dif_Calc_Load(object sender, EventArgs e)
        {
            //this.Enabled = true;
            FetchBasePrices();
            FetchCurrencyRates();
            FetchLabourPrices();
            GeneratePriceList();
            SetLablesToDefault();
            lbWait.Visible = false;
            button3.Visible = false;
            button3.Enabled = false;
            printButton.Visible = false;
            printButton.Enabled = false;
        }

        private void SetLablesToDefault()
        {
            string s = "0";
            lowestMargin = basePrices["17LowestMargin"];
            tbMainMargin.Text = basePrices["18DefaultMargin"].ToString();
            lblCostOfParts.Text = PriceRounding(float.Parse(s));
            lblCostIncludingMargin.Text = PriceRounding(float.Parse(s));
            lblGST.Text = PriceRounding(float.Parse(s));
            lblPriceIncludingGST.Text = PriceRounding(float.Parse(s));
            quoteNumber = ("Qu" + DateTime.Now.ToString("yy") + "MA-000");
            lblLift10Total.Text = s;
            lblLift11Total.Text = s;
            lblLift12Total.Text = s;
            lblLift1Total.Text = s;
            lblLift2Total.Text = s;
            lblLift3Total.Text = s;
            lblLift4Total.Text = s;
            lblLift5Total.Text = s;
            lblLift6Total.Text = s;
            lblLift7Total.Text = s;
            lblLift8Total.Text = s;
            lblLift9Total.Text = s;
            lblTotalLiftPrice.Text = "0";
            tBMainQuoteNumber.Text = quoteNumber;
        }

        #endregion

        #region Debug Box 
        public void DebugBoxCall(string textSent)
        {
            if (!debugMode)
            {
                return;
            }

            DialogResult dialogResult = MessageBox.Show(textSent + "\nDo you wish to continue run?", "Debug Log", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                return;
            }
            else if (dialogResult == DialogResult.No)
            {
                this.Close();
            }
        }
        #endregion

        #region Importing Data from XML Files

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
            int dKey = -1;
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

                if (dKey != -1 && dName != -1)
                {
                    labourPrice.Add(dKey, dName);
                    if (dKey > maxFloorNumber)
                    {
                        maxFloorNumber = dKey;
                    }
                    dKey = -1;
                    dName = -1;
                }
            }

            XMLR.Close();
            DebugBoxCall("Max Floor Number: " + maxFloorNumber);
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

        //read the XML file of a previous job and reload in its data
        private void FetchsaveData(string loadPath)
        {
            string dKey = "";
            string dName = "";

            XmlTextReader XMLR = new XmlTextReader(loadPath);
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Name")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Info")
                {
                    dName = XMLR.ReadElementContentAsString();
                }

                if (dKey != "" && dName != "")
                {
                    saveData.Add(dKey, dName);

                    dKey = "";
                    dName = "";
                }
            }

            XMLR.Close();
        }

        #endregion

        #region Setting Currency Exchange Rate

        //clicked AUD button
        private void buttonAUD_Click(object sender, EventArgs e)
        {
            SelectCurrency("A");
        }

        //clicked USD button
        private void buttonUSD_Click(object sender, EventArgs e)
        {
            SelectCurrency("U");
        }

        //clicked EUR button
        private void buttonEUR_Click(object sender, EventArgs e)
        {
            SelectCurrency("E");
        }

        //Called with what currency is selected and sets the exchange rate accordingly
        public void SelectCurrency(string selector)
        {
            currencySelectionGroup.Enabled = false;
            currencySelectionGroup.Visible = false;
            liftPricesPanel.Visible = true;
            liftPricesPanel.Enabled = true;

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

        #endregion

        #region Setting Shipping Costs

        //Clicked button to reset the prices for shipping containers
        private void btnShippingReset_Click(object sender, EventArgs e)
        {
            ShippingCalculation(1);
        }

        //Clicked button to add a 20ft shipping container
        private void btn20Ft_Click(object sender, EventArgs e)
        {
            ShippingCalculation(2);
        }

        //Clicked button to add a 40ft shipping container
        private void btn40Ft_Click(object sender, EventArgs e)
        {
            ShippingCalculation(3);
        }

        //Called from the shipping buttons to add the price of the selected container or to reset the prices
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

        #endregion

        #region Generating Price List

        // generate price list button
        private void button1_Click(object sender, EventArgs e)
        {
            GenerateListOfPrices();
        }

        //fix the text boxes to prevent any errors and checks that values are within thresholds
        public void GenerateListOfPrices()
        {
            // Textboxes = [tbSundries, tbShaftLight, tbDuct, tbAccomodation, tbCartage, tbStorage, tbTravel];

            TextBoxFixer(tbMainSundries);
            TextBoxFixer(tbMainShaftLight);
            TextBoxFixer(tbMainDuct);
            TextBoxFixer(tbMainAccomodation);
            TextBoxFixer(tbMainCartage);
            TextBoxFixer(tbMainStorage);
            TextBoxFixer(tbMainTravel);
            TextBoxFixer(tbMainBlankets);
            TextBoxFixer(tbMainScaffold);
            TextBoxFixer(tbMainEntranceGuards);
            TextBoxFixer(tbMainWeeksRequired);

            if (marginTbChecker()) // may need to add floors checker here??
            {
                GeneratePriceList();
                Form1SaveToXML();
                button3.Visible = true;
                button3.Enabled = true;
                printButton.Visible = true;
                printButton.Enabled = true;
            }
        }

        public bool CheckAddressForSlash()
        {
            bool b = tBMainAddress.Text.Contains(@"/");
            if (b)
            {
                MessageBox.Show("Addresses can't have slashes, please remove the slash");
                return !b;
            }
            return b;
        }

        //if a textbox is blank fill it with a 0 instead to prevent errors
        private void TextBoxFixer(TextBox tb)
        {
            if (tb.Text == "")
            {
                tb.Text = 0.ToString();
            }
        }

        //populate the price list with the $ values and add the totals to the end
        public void GeneratePriceList()
        {
            liftPrice = 0;

            //prices pulled from text boxes of program
            PriceListFormatting(lblSundries, float.Parse(tbMainSundries.Text));
            PriceListFormatting(lblShaft, float.Parse(tbMainShaftLight.Text));
            PriceListFormatting(lblDuct, float.Parse(tbMainDuct.Text));
            PriceListFormatting(lblAccommodation, float.Parse(tbMainAccomodation.Text));
            PriceListFormatting(lblStorage, float.Parse(tbMainStorage.Text));
            PriceListFormatting(lblTravel, float.Parse(tbMainTravel.Text));
            //prices pulled from base prices dictionary
            PriceListFormatting(lblFinishes, basePrices["1CarFinishes"]);
            PriceListFormatting(lblFire, basePrices["2FireExtinguisher"]);
            PriceListFormatting(lblGSM, basePrices["3GSMUnitPhone"]);
            //PriceListFormatting(lblBlanket, basePrices["4ProtectiveBlanket"]);
            PriceListFormatting(lblBlanket, float.Parse(tbMainBlankets.Text));
            PriceListFormatting(lblSump, basePrices["5SumpCover"]);
            PriceListFormatting(lblWiring, basePrices["6Wiring"]);
            PriceListFormatting(lblSign, basePrices["7Sinage"]);
            PriceListFormatting(lblElectrical, basePrices["8ElectricalBox"]);
            PriceListFormatting(lblCartage, float.Parse(tbMainCartage.Text));
            PriceListFormatting(lblDrawing, basePrices["9Drawings"]);
            PriceListFormatting(lblFork, basePrices["10ForkLift"]);
            PriceListFormatting(lblMaintenance, basePrices["11Maintenance"]);
            PriceListFormatting(lblManuals, basePrices["12Manuals"]);
            PriceListFormatting(lblWorkcover, basePrices["13WorkcoverFees"]);
            PriceListFormatting(lblScaffold, float.Parse(tbMainScaffold.Text) * basePrices["14Scaffolds"]);
            PriceListFormatting(lblEntrance, float.Parse(tbMainEntranceGuards.Text) * float.Parse(tbMainWeeksRequired.Text) * basePrices["15EntranceGuards"]);
            //add security from base prices dictionary if box is checked
            if (cbMainSecurity.Checked)
            {
                PriceListFormatting(lblSecurity, basePrices["Security"] + (basePrices["SecurityPerFloor"] * FloorsAdder()));
            }
            else
            {
                PriceListFormatting(lblSecurity, 0);
            }

            TotalCostAdder();
            PriceListFormatting(lblCost, float.Parse(lblTotalLiftPrice.Text) * applicableExchangeRate);

            // if needed show the unconverted cost of the lift
            if (exCurrency == 0)
            {
                lblLiftNoConvert.Visible = false;
                lblLiftNoConvertPrice.Visible = false;
                costInEuro = false;
            }
            else if (exCurrency == 1)
            {
                lblLiftNoConvert.Visible = true;
                lblLiftNoConvertPrice.Visible = true;
                lblLiftNoConvert.Text = "Cost of Lift (USD)";
                costInEuro = false;
                lblLiftNoConvertPrice.Text = PriceRounding(float.Parse(lblTotalLiftPrice.Text), costInEuro);
            }
            else if (exCurrency == 2)
            {
                lblLiftNoConvert.Visible = true;
                lblLiftNoConvertPrice.Visible = true;
                lblLiftNoConvert.Text = "Cost of Lift (EUR)";
                costInEuro = true;
                lblLiftNoConvertPrice.Text = PriceRounding(float.Parse(lblTotalLiftPrice.Text), costInEuro);
            }

            //add freight based on number of required containers 
            PriceListFormatting(lblFreight, freightTotal);
            //add labour from the labour costs dictionary based on number of floors in the building 
            PriceListFormatting(lblLabour, LabourAdder());

            marginPercent = 1 + (float.Parse(tbMainMargin.Text) / 100);
            float marginValue = (float.Parse(tbMainMargin.Text) / 100) * liftPrice;
            //liftPrice *= int.Parse(tbNumberLifts.Text);

            lblCostOfParts.Text = PriceRounding(liftPrice);
            lblCostIncludingMargin.Text = PriceRounding(liftPrice * marginPercent); //+ " (" + PriceRounding(marginValue) + ")";
            lblGST.Text = PriceRounding(liftPrice * marginPercent * 0.1f);
            lblPriceIncludingGST.Text = PriceRounding(liftPrice * marginPercent * 1.1f);
        }

        private float FloorsAdder()
        {
            float floors = 0;
            floors += int.Parse(tbLift1Floors.Text);
            floors += int.Parse(tbLift2Floors.Text);
            floors += int.Parse(tbLift3Floors.Text);
            floors += int.Parse(tbLift4Floors.Text);
            floors += int.Parse(tbLift5Floors.Text);
            floors += int.Parse(tbLift6Floors.Text);
            floors += int.Parse(tbLift7Floors.Text);
            floors += int.Parse(tbLift8Floors.Text);
            floors += int.Parse(tbLift9Floors.Text);
            floors += int.Parse(tbLift10Floors.Text);
            floors += int.Parse(tbLift11Floors.Text);
            floors += int.Parse(tbLift12Floors.Text);
            return floors;
        }

        private float LabourAdder()
        {
            TextBox[] tbs = { tbLift1Floors, tbLift2Floors, tbLift3Floors, tbLift4Floors, tbLift5Floors, tbLift6Floors, tbLift7Floors, tbLift7Floors, tbLift8Floors, tbLift9Floors, tbLift10Floors, tbLift11Floors, tbLift12Floors };
            float labour = 0;

            foreach (TextBox i in tbs)
            {
                floorsTbChecker(i);
            }

            labour += labourPrice[int.Parse(tbLift1Floors.Text)];
            labour += labourPrice[int.Parse(tbLift2Floors.Text)];
            labour += labourPrice[int.Parse(tbLift3Floors.Text)];
            labour += labourPrice[int.Parse(tbLift4Floors.Text)];
            labour += labourPrice[int.Parse(tbLift5Floors.Text)];
            labour += labourPrice[int.Parse(tbLift6Floors.Text)];
            labour += labourPrice[int.Parse(tbLift7Floors.Text)];
            labour += labourPrice[int.Parse(tbLift8Floors.Text)];
            labour += labourPrice[int.Parse(tbLift9Floors.Text)];
            labour += labourPrice[int.Parse(tbLift10Floors.Text)];
            labour += labourPrice[int.Parse(tbLift11Floors.Text)];
            labour += labourPrice[int.Parse(tbLift12Floors.Text)];
            return labour;
        }

        //Sends prices through the rounder method as well as adding them to the total cost of the lift for the total
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

        //rounds all prices to 2 decimal places and added the applicable currency symbols as well as commas and decimal points
        private string PriceRounding(float s, bool isPriceInEuro = false)
        {
            if (isPriceInEuro)
            {
                return "€" + Math.Round(s, 2).ToString("N", new System.Globalization.CultureInfo("en-US"));
            }
            else
            {
                return "$" + Math.Round(s, 2).ToString("N", new System.Globalization.CultureInfo("en-US"));
            }
        }

        //Checks that the floors entered is within the acceptable threshold
        public bool floorsTbChecker(TextBox tb)
        {
            int i = 0;
            try
            {
                i = int.Parse(tb.Text);
            }
            catch
            {
                //MessageBox.Show("Invalid floor number entered ");
                tb.Text = "0";
            }

            if (i > maxFloorNumber || i < 0 || i == 1)
            {
                MessageBox.Show("Invalid floor number entered ");
                return false;
            }
            else
            {
                return true;
            }
        }

        //checks that the margin is within the acceptable threshold
        public bool marginTbChecker()
        {
            try
            {
                if (tbMainMargin.Text == "")
                {
                    MessageBox.Show("Margin % is below the allowed minimum of " + lowestMargin + "%");
                    return false;
                }
                else if (lowestMargin > float.Parse(tbMainMargin.Text))
                {
                    MessageBox.Show("Margin % is below the allowed minimum of " + lowestMargin + "%");
                    return false;
                }
            }
            catch
            {
                MessageBox.Show("Margin % is below the allowed minimum of " + lowestMargin + "%");
                return false;
            }

            return true;
        }

        #endregion

        #region Generate Quote Word Document

        // export to quote button
        private void button3_Click(object sender, EventArgs e)
        {
            GenerateListOfPrices();
            WordSetup();//find and set vars to the quote template document 
            WordSave(false); // save the doc
            if (sucessfulSave)
            {
                WordData("AE101", tBMainAddress.Text); //address
                WordData("AE102", tBMainQuoteNumber.Text);//quote number
                WordData("AE103", TotalLifts().ToString());//number of lifts

                Pin_Dif_Exp qI = new Pin_Dif_Exp();
                qI.Show();//open questionaire 
                //questions complete method called from final form of querstions to continue the export to word function. 
            }
            else
            {
                MessageBox.Show("Saving Error, Document not saved");
                lblWaitControl(false);
            }
        }

        // sets up the word document ready to be written
        public void WordSetup()
        {
            lblWaitControl(true);
            string filePath = @"X:\Program Dependancies\Quote tool\Template Word Docs\Template-" + salesRep + "-Diff-" + numberOfPagesNeeded + ".docx";
            fileOpen = new Word.Application();
            fileOpen.Visible = true;
            document = fileOpen.Documents.Open(filePath, ReadOnly: false);
            document.Activate();
        }

        // if false,asks where to save the word doc and saves it. if true saves in the previously set location
        private void WordSave(bool b)
        {
            if (!b)
            {
                saveFileDialog1.Title = ("Where to save the quote");
                saveFileDialog1.InitialDirectory = "X:\\Sales\\Qu-" + DateTime.Now.ToString("yyyy");
                saveFileDialog1.FileName = tBMainQuoteNumber.Text + " - " + tBMainAddress.Text + " Quote";
                saveFileDialog1.DefaultExt = "docx";
                saveFileDialog1.Filter = "Word Doc (*.docx) |*.docx|All files (*.*) |*.*";
                //saveFileDialog1.ShowDialog();
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    document.SaveAs2(saveFileDialog1.FileName);
                    sucessfulSave = true;
                }
                else
                {
                    sucessfulSave = false;
                    document.Close();
                }
            }
            else if (b)
            {
                document.SaveAs2(saveFileDialog1.FileName);
            }
        }

        // called from question forms to take data and write it to the dictionary
        public void WordData(string k, string v)
        {
            //wordExportData.Add(k, v);
            wordExportData[k] = v;
            //MessageBox.Show("Word Data Method called with: " + k + " " + v);
        }

        //called from the final question to continue the export to word function 
        public void QuestionsComplete()
        {
            WordData("AE211", FormalDate());
            WordData("AE212", lblCostIncludingMargin.Text);
            WordData("AE213", lblGST.Text);
            WordData("AE214", lblPriceIncludingGST.Text);

            float totalPrice = float.Parse(lblTotalLiftPrice.Text) * applicableExchangeRate;
            WordData("AE220", PriceRounding(totalPrice));

            WordReplaceLooper(wordExportData);// loop the find and replace method to populate the info 
            WordSave(true);// save the doc again 
            SaveReloadXMLFile(saveData);
            WordFinish();//finish the methods 

        }

        // loops through the word document performing a find and replace operation
        private void WordReplaceLooper(Dictionary<string, string> d)
        {
            foreach (KeyValuePair<string, string> i in d)
            {
                FindAndReplace(fileOpen, i.Key, i.Value);
            }
        }

        //Finds text A and replaces with text B while maintaining all the relevent settings to maintain the word document
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

        // closes the word document 
        private void WordFinish()
        {
            //fileOpen.ShowMe();
            fileOpen.Quit();

            MessageBox.Show("Quote sucessfully exported");
            lblWaitControl(false);
        }

        // method called by a closing Question form to signal the program to close the word document being used and to wipe the data from the dictionary
        public void QuestionCloseCall(Form f)
        {
            //Form[] questionForms = new Form[11];
            DialogResult d = MessageBox.Show("Are you sure you wish to cancel quote exporting?", "Close?", MessageBoxButtons.YesNo);

            if (d == DialogResult.Yes)
            {
                wordExportData.Clear();
                fileOpen.Quit();

                MessageBox.Show("Quote sucessfully exported");
                lblWaitControl(false);
                MessageBox.Show("Word Export Canceled");
                f.Close();
            }
        }

        public void lblWaitControl(bool b)
        {
            // if called true will enable the wait message and disable the form
            // if called false will disable the wait message and enable the form 

            lbWait.Enabled = b;
            lbWait.Visible = b;
            this.Enabled = !b;
        }

        #endregion

        #region Export Prices to Word Document
        //uses some additional methods from the "Generate Quote Word Document" region

        private void printButton_Click(object sender, EventArgs e)
        {
            GenerateListOfPrices();
            lblWaitControl(true);
            printButton.BackColor = Color.Blue;
            if (SavePricesDocument())
            {
                MessageBox.Show("Prices exported as " + saveFileDialog1.FileName);
                Form1SaveToXML();
                SaveReloadXMLFile(saveData);
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
            saveFileDialog1.FileName = tBMainQuoteNumber.Text + " - " + tBMainAddress.Text + " Price Breakdown";
            saveFileDialog1.DefaultExt = "docx";
            saveFileDialog1.Filter = "Word Doc (*.docx) |*.docx| All files (*.*) |*.*";

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileOpen = new Word.Application();
                document = fileOpen.Documents.Open("X:\\Program Dependancies\\Quote tool\\Template Word Docs\\Template-" + salesRep + "-Price-" + numberOfPagesNeeded + ".docx", ReadOnly: false);
                SavePricesToDict();
                fileOpen.Visible = true;
                document.Activate();
                WordReplaceLooper(priceExports);
                document.SaveAs2(saveFileDialog1.FileName);
                document.Close();
                fileOpen.Quit();
                return true;
            }
            else
            {
                try
                {
                    //document.Close();
                }
                catch
                {
                    return false;
                }
                return false;
            }
        }

        public void SavePricesToDict()

        {
            float f = liftPrice * (marginPercent - 1);
            float f2;
            priceExports.Clear();

            priceExports.Add("AEP1", tBMainAddress.Text);
            priceExports.Add("AEP2", tBMainQuoteNumber.Text);
            priceExports.Add("AEP3", FormalDate());
            priceExports.Add("AEP4", exchangeRateText);

            f2 = float.Parse(tbLift1Price.Text) * applicableExchangeRate;
            priceExports.Add("P1AEP5", PriceRounding(float.Parse(tbLift1Price.Text), costInEuro));
            priceExports.Add("P1AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P1AEP38", tbLift1Floors.Text);

            f2 = float.Parse(tbLift2Price.Text) * applicableExchangeRate;
            priceExports.Add("P2AEP5", PriceRounding(float.Parse(tbLift2Price.Text), costInEuro));
            priceExports.Add("P2AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P2AEP38", tbLift2Floors.Text);

            f2 = float.Parse(tb3Lift3Price.Text) * applicableExchangeRate;
            priceExports.Add("P3AEP5", PriceRounding(float.Parse(tb3Lift3Price.Text), costInEuro));
            priceExports.Add("P3AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P3AEP38", tbLift3Floors.Text);

            f2 = float.Parse(tbLift4Price.Text) * applicableExchangeRate;
            priceExports.Add("P4AEP5", PriceRounding(float.Parse(tbLift4Price.Text), costInEuro));
            priceExports.Add("P4AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P4AEP38", tbLift4Floors.Text);

            f2 = float.Parse(tbLift5Price.Text) * applicableExchangeRate;
            priceExports.Add("P5AEP5", PriceRounding(float.Parse(tbLift5Price.Text), costInEuro));
            priceExports.Add("P5AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P5AEP38", tbLift5Floors.Text);

            f2 = float.Parse(tbLift6Price.Text) * applicableExchangeRate;
            priceExports.Add("P6AEP5", PriceRounding(float.Parse(tbLift6Price.Text), costInEuro));
            priceExports.Add("P6AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P6AEP38", tbLift6Floors.Text);

            f2 = float.Parse(tbLift7Price.Text) * applicableExchangeRate;
            priceExports.Add("P7AEP5", PriceRounding(float.Parse(tbLift7Price.Text), costInEuro));
            priceExports.Add("P7AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P7AEP38", tbLift7Floors.Text);

            f2 = float.Parse(tbLift8Price.Text) * applicableExchangeRate;
            priceExports.Add("P8AEP5", PriceRounding(float.Parse(tbLift8Price.Text), costInEuro));
            priceExports.Add("P8AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P8AEP38", tbLift8Floors.Text);

            f2 = float.Parse(tbLift9Price.Text) * applicableExchangeRate;
            priceExports.Add("P9AEP5", PriceRounding(float.Parse(tbLift9Price.Text), costInEuro));
            priceExports.Add("P9AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P9AEP38", tbLift9Floors.Text);

            f2 = float.Parse(tbLift10Price.Text) * applicableExchangeRate;
            priceExports.Add("P10AEP5", PriceRounding(float.Parse(tbLift10Price.Text), costInEuro));
            priceExports.Add("P10AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P10AEP38", tbLift10Floors.Text);

            f2 = float.Parse(tbLift11Price.Text) * applicableExchangeRate;
            priceExports.Add("P11AEP5", PriceRounding(float.Parse(tbLift11Price.Text), costInEuro));
            priceExports.Add("P11AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P11AEP38", tbLift11Floors.Text);

            f2 = float.Parse(tbLift12Price.Text) * applicableExchangeRate;
            priceExports.Add("P12AEP5", PriceRounding(float.Parse(tbLift12Price.Text), costInEuro));
            priceExports.Add("P12AEP6", PriceRounding(float.Parse(f2.ToString()), false));
            priceExports.Add("P12AEP38", tbLift12Floors.Text);

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
            priceExports.Add("AEP33", tbMainMargin.Text + "%");
            priceExports.Add("AEP34", lblCostIncludingMargin.Text);
            priceExports.Add("AEP35", lblGST.Text);
            priceExports.Add("AEP36", lblPriceIncludingGST.Text);
            priceExports.Add("AEP37", PriceRounding(f).ToString());
        }

        #endregion

        #region Load Data from old quote via XML file

        private void btLoad_Click(object sender, EventArgs e)
        {
            LoadPreviousQuote();
            GenerateListOfPrices();
        }

        private void LoadPreviousQuote()
        {
            // prompt user to select word doc
            // remove "quote" or "price breakdown" and file extenstion
            // find XML doc with the same name in the dependancies folder
            // call seprate method to load data from that XML into a dictionary 
            // populate form with data from dictionary
            // when opening each question form load all relevent data from the dictionary from Form1

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string xmlPath = FindXmlFile(openFileDialog1.FileName);
                if (xmlPath != null)
                {
                    saveData.Clear();
                    xmlPath = @"X:\Program Dependancies\Quote tool\Previous Prices\" + xmlPath + ".xml";
                    //MessageBox.Show(xmlPath);
                    loadingPreviousData = true;
                    FetchsaveData(xmlPath);
                    Form1LoadFromXML();
                    GenerateListOfPrices();
                    liftPricesPanel.Visible = false;
                }
                else
                {
                    MessageBox.Show("Invalid file selected");
                }
            }
            else
            {
                MessageBox.Show("Invalid file selected");
            }
        }

        private string FindXmlFile(string fileName)
        {
            string rtnString = null;
            try
            {
                if (fileName.Contains("Quote"))
                {
                    string subString = " Quote.docx";
                    int index = fileName.IndexOf(subString);
                    rtnString = fileName.Remove(index, subString.Length);
                    int lastIndex = rtnString.LastIndexOf(@"\", rtnString.Length);
                    rtnString = rtnString.Remove(0, lastIndex + 1);
                    return rtnString;
                }
                else if (fileName.Contains("Price Breakdown"))
                {
                    string subString = " Price Breakdown.docx";
                    int index = fileName.IndexOf(subString);
                    rtnString = fileName.Remove(index, subString.Length);
                    int lastIndex = rtnString.LastIndexOf(@"\", rtnString.Length);
                    rtnString = rtnString.Remove(0, lastIndex + 1);
                    return rtnString;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("Unable to find Load Data for this Quote");
                return null;
            }
        }

        private void Form1LoadFromXML()
        {
            if (saveData.ContainsKey("NumberOfPagesOpen"))
            {
                numberOfPagesNeeded = int.Parse(saveData["NumberOfPagesOpen"]);
            }
            LoadPreviousXmlTb(tbMainAccomodation, tBMainAddress, tbMainBlankets, tbMainCartage, tbMainDuct, tbMainEntranceGuards, tbMainMargin,
                 tBMainQuoteNumber, tbMainScaffold, tbMainShaftLight, tbMainStorage, tbMainSundries, tbMainTravel, tbMainWeeksRequired,
                 tbLift1Price, tbLift1Floors, tbLift2Price, tbLift2Floors, tb3Lift3Price, tbLift3Floors,
                 tbLift4Price, tbLift4Floors, tbLift5Price, tbLift5Floors, tbLift6Price, tbLift6Floors,
                 tbLift7Price, tbLift7Floors, tbLift8Price, tbLift8Floors, tbLift9Price, tbLift9Floors,
                 tbLift10Price, tbLift10Floors, tbLift11Price, tbLift11Floors, tbLift12Price, tbLift12Floors);
            LoadPreviousXmlCb(cbMainSecurity);
            //num20Ft = int.Parse(saveData["num20Ft"]);
            //num40Ft = int.Parse(saveData["num40Ft"]);
            ShippingCalculation(1);
            for (int i = 0; i < int.Parse(saveData["num20Ft"]); i++)
            {
                ShippingCalculation(2);
            }
            for (int i = 0; i < int.Parse(saveData["num40Ft"]); i++)
            {
                ShippingCalculation(3);
            }

            if (int.Parse(saveData["exCurrency"]) == 0)
            {
                //AUD
                SelectCurrency("A");
            }
            else if (int.Parse(saveData["exCurrency"]) == 1)
            {
                //USD
                SelectCurrency("U");
            }
            else if (int.Parse(saveData["exCurrency"]) == 2)
            {
                //EUR
                SelectCurrency("E");
            }
        }

        public void LoadPreviousXmlTb(params TextBox[] tb)
        {
            foreach (TextBox Box in tb)
            {
                if (saveData.ContainsKey(Box.Name.ToString()))
                {
                    Box.Text = saveData[Box.Name.ToString()];
                }
            }
        }

        public void LoadPreviousXmlCb(params CheckBox[] cb)
        {
            foreach (CheckBox Box in cb)
            {
                if (saveData.ContainsKey(Box.Name.ToString()))
                {
                    Box.Checked = bool.Parse(saveData[Box.Name.ToString()]);
                }
            }
        }

        public void LoadPreviousXmlRb(params RadioButton[] rb)
        {
            foreach (RadioButton radio in rb)
            {
                if (saveData.ContainsKey(radio.Name.ToString()))
                {
                    radio.Checked = bool.Parse(saveData[radio.Name.ToString()]);
                }
                /*
                if (radio.Checked == true && radio.Text == "")
                {
                    tb.Text = saveData[tb.Name.ToString()];
                    return;
                }
                */
            }
        }

        #endregion

        #region Save data to XML file for future loading

        public void SaveTbToXML(params TextBox[] tb)
        {
            foreach (TextBox item in tb)
            {
                //saveData.Add(item.Name.ToString(), item.Text.ToString());
                saveData[item.Name.ToString()] = item.Text.ToString();
            }
        }

        public void SaveRbToXML(params RadioButton[] rb)
        {
            foreach (RadioButton item in rb)
            {
                //saveData.Add(item.Name.ToString(), item.Checked.ToString());
                saveData[item.Name.ToString()] = item.Checked.ToString();
            }
        }

        public void SaveCbToXML(params CheckBox[] cb)
        {
            foreach (CheckBox item in cb)
            {
                //saveData.Add(item.Name.ToString(), item.Checked.ToString());
                saveData[item.Name.ToString()] = item.Checked.ToString();
            }
        }

        private void SaveReloadXMLFile(Dictionary<string, string> kvp)
        {
            //string path = "X:\\Program Dependancies\\Quote tool\\Previous Prices\\" + saveFileDialog1.FileName.ToString() + ".xml";
            string path = "X:\\Program Dependancies\\Quote tool\\Previous Prices\\" + tBMainQuoteNumber.Text.ToString() + " - " + tBMainAddress.Text.ToString() + ".xml";

            XmlTextWriter xmlWriter = new XmlTextWriter(path, Encoding.UTF8);
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("Data");

            foreach (KeyValuePair<string, string> i in kvp)
            {
                xmlWriter.WriteStartElement("Object");

                xmlWriter.WriteElementString("Name", i.Key);
                xmlWriter.WriteElementString("Info", i.Value);

                xmlWriter.WriteEndElement(); //Object end
            }
            xmlWriter.WriteEndElement();//Data end
            xmlWriter.Close();
        }

        private void Form1SaveToXML()
        {
            SaveTbToXML(tbMainAccomodation, tBMainAddress, tbMainBlankets, tbMainCartage, tbMainDuct, tbMainEntranceGuards, tbMainMargin,
                 tBMainQuoteNumber, tbMainScaffold, tbMainShaftLight, tbMainStorage, tbMainSundries, tbMainTravel, tbMainWeeksRequired,
                 tbLift1Price, tbLift1Floors, tbLift2Price, tbLift2Floors, tb3Lift3Price, tbLift3Floors,
                 tbLift4Price, tbLift4Floors, tbLift5Price, tbLift5Floors, tbLift6Price, tbLift6Floors,
                 tbLift7Price, tbLift7Floors, tbLift8Price, tbLift8Floors, tbLift9Price, tbLift9Floors,
                 tbLift10Price, tbLift10Floors, tbLift11Price, tbLift11Floors, tbLift12Price, tbLift12Floors);
            SaveCbToXML(cbMainSecurity);
            saveData["num20Ft"] = num20Ft.ToString();
            saveData["num40Ft"] = num40Ft.ToString();
            saveData["exCurrency"] = exCurrency.ToString();
        }
        #endregion

        #region Data Formatting methods for external calls

        public string CheckboxTrueToYes(CheckBox cb)
        {
            if (cb.Checked)
            {
                return "Yes";
            }
            else
            {
                return "No";
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

            //if the days number ends in a 1 it requires either "th" or "st" suffix 
            if (endIn1)
            {
                // if it is single digit it means it is the 1st and requires the "st" suffix
                if (singleDigit)
                {
                    day += "st";
                }
                // if it starts with a 1 and is 2 digits it is the 11th and requires the "th" suffix
                else if (startWith1 && !singleDigit)
                {
                    day += "th";
                }
                //if it starts with any other number and si not single digit it means it is not the 1st or the 11th and requires the "st" suffix
                else
                {
                    day += "st";
                }
            }
            //if the days number ends in a 2 it requires either "th" or "nd" suffix
            else if (endIn2)
            {
                // if it is single digit it means it is the 2nd and requires the "nd" suffix
                if (singleDigit)
                {
                    day += "nd";
                }
                // if it starts with a 1 and is 2 digits it is the 12th and requires the "th" suffix
                else if (startWith1 && !singleDigit)
                {
                    day += "th";
                }
                //if it starts with any other number and si not single digit it means it is not the 2nd or the 12th and requires the "nd" suffix
                else
                {
                    day += "nd";
                }
            }
            //if the days number ends in 3 it requires either "th" or "rd" suffix
            else if (endIn3)
            {
                // if it is single digit it means it is the 3rd and requires the "rd" suffix
                if (singleDigit)
                {
                    day += "rd";
                }
                // if it starts with a 1 and is 2 digits it is the 13th and requires the "th" suffix
                else if (startWith1 && !singleDigit)
                {
                    day += "th";
                }
                //if it starts with any other number and si not single digit it means it is not the 3rd or the 13th and requires the "rd" suffix
                else
                {
                    day += "rd";
                }
            }
            // if the days number ends in any other number it requires a "th" suffix
            else
            {
                day += "th";
            }

            //take the correctly formatted date number and add on a full length month and year 
            string date = day + " " + monthYear;
            //return the date as a string 
            return date;
        }

        //called to check if the text has the correct measurment suffix before exporting to quote document
        public string MeasureStringChecker(string text, string measurementSuffix)
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

        // called with all the radio buttons in each group to find which is checked and return the text
        public string RadioButtonHandeler(TextBox tb = null, params RadioButton[] rb)
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

        public string RadioButtonToAsteriskHandeler(RadioButton yes, RadioButton no)
        {
            if (yes.Checked == true)
            {
                return "*";
            }
            else
            {
                return "";
            }
        }

        //called with all the checkboxes in each group, it will then read them and determine how to return a string
        public string CheckBoxHandler(params CheckBox[] cB)
        {
            string str = "";
            int count = 0;

            // count the number of checked boxes
            foreach (CheckBox i in cB)
            {
                if (i.Checked)
                {
                    count++;
                }
            }

            //loop through all the check boxes in the array
            foreach (CheckBox i in cB)
            {
                if (i.Checked)
                {
                    //when a checked box is found add it's text to the end of the string
                    str += i.Text;

                    //if the counter is 1 or lower it means that the last box processed was the final checked box an thus requires no connector for addition boxes texts. 
                    if (count > 0)
                    {
                        // if there is 3 checked boxes or more remaining add a comma
                        if (count >= 3)
                        {
                            str += ", ";
                        }
                        // if there is 2 checked boxes remaining add an &
                        else if (count == 2)
                        {
                            str += " & ";
                        }
                        //then reduce the counter of remaining checked boxes
                        count--;
                    }
                }
            }

            //return the final string of all checkboxes texts joined. 
            return str;
        }

        #endregion

        #region Close Form 

        private void button2_Click_1(object sender, EventArgs e) // close button
        {
            if (document != null)
            {
                try
                {
                    document.Close();
                }
                catch (Exception)
                {
                    // return;
                }
            }
            this.Close();
        }

        #endregion

        #region unused methods
        private void liftPricesPanel_Paint(object sender, PaintEventArgs e)
        {
            //
        }

        private void label34_Click(object sender, EventArgs e)
        {
            //
        }


        private void label36_Click(object sender, EventArgs e)
        {
            //
        }


        private void label38_Click(object sender, EventArgs e)
        {
            //
        }


        private void label40_Click(object sender, EventArgs e)
        {
            //
        }




        private void label14_Click(object sender, EventArgs e)
        {
            //
        }

        private void label13_Click(object sender, EventArgs e)
        {
            //
        }

        private void lbWait_Click(object sender, EventArgs e)
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

        #region Update Quote Number

        private void tBQuoteNumber_TextChanged(object sender, EventArgs e)
        {
            //when the quote number text box is changed, if it is different to the existing quote number update the quote number
            if (quoteNumber != tBMainQuoteNumber.Text)
            {
                quoteNumber = tBMainQuoteNumber.Text;
            }
        }

        #endregion

        #region Configure Lift Prices Menu
        private void BtGenerateLiftPrices_Click(object sender, EventArgs e)
        {
            TotalCostAdder();
            PagesRequired();
            GenerateListOfPrices();
            DebugBoxCall("number of pages needed " + numberOfPagesNeeded.ToString());
        }

        private void PagesRequired()
        {
            int i = 0;
            if (tbLift1Floors.Text == "0")
            {
                i++;
            }
            if (tbLift2Floors.Text == "0")
            {
                i++;
            }
            if (tbLift3Floors.Text == "0")
            {
                i++;
            }
            if (tbLift4Floors.Text == "0")
            {
                i++;
            }
            if (tbLift5Floors.Text == "0")
            {
                i++;
            }
            if (tbLift6Floors.Text == "0")
            {
                i++;
            }
            if (tbLift7Floors.Text == "0")
            {
                i++;
            }
            if (tbLift8Floors.Text == "0")
            {
                i++;
            }
            if (tbLift9Floors.Text == "0")
            {
                i++;
            }
            if (tbLift10Floors.Text == "0")
            {
                i++;
            }
            if (tbLift11Floors.Text == "0")
            {
                i++;
            }
            if (tbLift12Floors.Text == "0")
            {
                i++;
            }
            numberOfPagesNeeded = 12 - i;
        }
        #region Text Changed Methods
        private void tbLift1Floors_TextChanged(object sender, EventArgs e)
        {
            floorsTbChecker(tbLift1Floors);
        }

        private void tbLift2Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift2Floors);
        }

        private void tbLift3Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift3Floors);
        }

        private void tbLift4Floors_TextChanged(object sender, EventArgs e)
        {
            floorsTbChecker(tbLift4Floors);

        }

        private void tbLift5Floors_TextChanged(object sender, EventArgs e)
        {
            floorsTbChecker(tbLift5Floors);

        }

        private void tbLift6Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift6Floors);
        }

        private void tbLift7Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift7Floors);
        }

        private void tbLift8Floors_TextChanged(object sender, EventArgs e)
        {
            floorsTbChecker(tbLift8Floors);

        }

        private void tbLift9Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift9Floors);
        }

        private void tbLift10Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift10Floors);
        }

        private void tbLift11Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift11Floors);
        }

        private void tbLift12Floors_TextChanged(object sender, EventArgs e)
        {

            floorsTbChecker(tbLift12Floors);
        }
        private void textBox5_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift8Price, lblLift8Total);
        }
        private void textBox12_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift9Price, lblLift9Total);
        }
        private void textBox11_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift10Price, lblLift10Total);
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift11Price, lblLift11Total);
        }
        private void textBox9_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift12Price, lblLift12Total);
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            RefreshLiftPrices(tbLift1Price, lblLift1Total);
        }

        private void tbLift2Price_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift2Price, lblLift2Total);
        }

        private void tb3Lift3Price_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tb3Lift3Price, lblLift3Total);
        }

        private void tbLift4Price_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift4Price, lblLift4Total);
        }

        private void tbLift5Price_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift5Price, lblLift5Total);
        }

        private void tbLift6Price_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift6Price, lblLift6Total);
        }

        private void tbLift7Price_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift7Price, lblLift7Total);
        }

        private void tbLift1Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift1Price, lblLift1Total);
        }

        private void tbLift2Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift2Price, lblLift2Total);
        }

        private void tbLift3Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tb3Lift3Price, lblLift3Total);
        }

        private void tbLift4Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift4Price, lblLift4Total);
        }

        private void tbLift5Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift5Price, lblLift5Total);
        }

        private void tbLift6Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift6Price, lblLift6Total);
        }

        private void tbLift7Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift7Price, lblLift7Total);
        }

        private void tbLift8Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift8Price, lblLift8Total);
        }

        private void tbLift10Number_TextChanged(object sender, EventArgs e)
        {
            RefreshLiftPrices(tbLift10Price, lblLift10Total);

        }

        private void tbLift11Numebr_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift11Price, lblLift11Total);
        }
        private void tbLift9Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift9Price, lblLift9Total);
        }

        private void tbLift12Number_TextChanged(object sender, EventArgs e)
        {

            RefreshLiftPrices(tbLift12Price, lblLift12Total);
        }
        #endregion
        private void btnConfigurePrices_Click(object sender, EventArgs e)
        {
            liftPricesPanel.Enabled = true;
            liftPricesPanel.Visible = true;
        }

        private void HideButton_Click(object sender, EventArgs e)
        {
            SendPricesToMainForm();
            liftPricesPanel.Visible = false;
        }

        private void SendPricesToMainForm()
        {/*
            bool euro = false;
            if (exCurrency == 2)
            {
                euro = true;
            }
            lblLiftNoConvertPrice.Text = PriceRounding(float.Parse(lblTotalLiftPrice.Text));
            PriceListFormatting(lblCost, float.Parse(lblTotalLiftPrice.Text), euro);
            */
        }

        private void RefreshLiftPrices(TextBox price, Label total)
        {
            try
            {
                total.Text = (float.Parse(price.Text)).ToString();
            }
            catch (Exception)
            {

            }
        }

        public int TotalLifts()
        {
            int total = 0;
            TextBox[] lifts = { tbLift10Floors, tbLift11Floors, tbLift12Floors, tbLift1Floors, tbLift2Floors, tbLift3Floors, tbLift4Floors, tbLift5Floors, tbLift6Floors, tbLift7Floors, tbLift8Floors, tbLift9Floors };

            foreach (TextBox tb in lifts)
            {
                try
                {
                    if (float.Parse(tb.Text) != 0)
                    {
                        total++;
                    }
                }
                catch (Exception)
                {
                }
            }
            return total;
        }

        private void TotalCostAdder()
        {
            lblWaitControl(true);
            float f = 0;
            f += LabelToFloat(lblLift1Total);
            f += LabelToFloat(lblLift2Total);
            f += LabelToFloat(lblLift3Total);
            f += LabelToFloat(lblLift4Total);
            f += LabelToFloat(lblLift5Total);
            f += LabelToFloat(lblLift6Total);
            f += LabelToFloat(lblLift7Total);
            f += LabelToFloat(lblLift8Total);
            f += LabelToFloat(lblLift9Total);
            f += LabelToFloat(lblLift10Total);
            f += LabelToFloat(lblLift11Total);
            f += LabelToFloat(lblLift12Total);
            lblTotalLiftPrice.Text = f.ToString();
            lblWaitControl(false);
        }

        private float LabelToFloat(Label tb)
        {
            float f = -1;
            try
            {
                f = float.Parse(tb.Text);
            }
            catch (Exception)
            {
                return f;
            }
            return f;
        }

        #endregion

    }
}