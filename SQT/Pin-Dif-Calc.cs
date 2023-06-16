using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace SQT
{
    public partial class Pin_Dif_Calc : Form
    {
        #region VARS
        MainMenu mm = Application.OpenForms.OfType<MainMenu>().Single();

        public bool sucessfulSave = false;
        public bool loadingPreviousData = false;
        bool costInEuro;
        bool rearDoorChecker = false;
        bool[] exporterPageOpened = { false, false, false, false, false, false, false, false, false, false, false, false };

        public string quoteNumber = "";
        public string exchangeRateText;

        public float applicableExchangeRate = 1;
        public float freightTotal = 0;
        public float liftPrice;
        public float lowestMargin;
        float marginPercent;

        public int exCurrency = 0; // 0 AUD, 1 USD, 2 EUR
        public int num20Ft;
        public int num40Ft;
        public int numberOfPagesNeeded = 1;
        public int numberOfPagesUsed = 0;
        int pageTracker;
        int activePage = 0;

        public Dictionary<string, string> wordExportData = new Dictionary<string, string>();
        public Dictionary<string, string> priceExports = new Dictionary<string, string>();
        public Dictionary<string, string> saveData = new Dictionary<string, string>();

        Word.Application fileOpen;
        Word.Document document;

        Button[] pageButtons; // = new Button[12];
        Panel[] infoPages;
        Object[,] costsGroups;
        #endregion

        #region Form Loading Methods
        public Pin_Dif_Calc()
        {
            InitializeComponent();

        }

        private void Pin_Dif_Calc_Load(object sender, EventArgs e)
        {
            RuntimeVarSetter();
            SetupPricesPanel();
            GeneratePriceList();
            SetLablesToDefault();
            SetPanelVisabilityDefaults();
            PageButtonSetup();
            this.Text = tBMainQuoteNumber.Text + " Quote Calculation";
        }

        private void RuntimeVarSetter()
        {
            costsGroups = new Object[12, 4] {
                { label5, tbLift1Price, tbLift1Floors, lblLift1Total },
                { label7, tbLift2Price, tbLift2Floors, lblLift2Total },
                { label9, tb3Lift3Price, tbLift3Floors, lblLift3Total },
                { label12, tbLift4Price, tbLift4Floors, lblLift4Total },
                { label32, tbLift5Price, tbLift5Floors, lblLift5Total },
                { label30, tbLift6Price, tbLift6Floors, lblLift6Total },
                { label28, tbLift7Price, tbLift7Floors, lblLift7Total },
                { label14, tbLift8Price, tbLift8Floors, lblLift8Total },
                { label40, tbLift9Price, tbLift9Floors, lblLift9Total },
                { label38, tbLift10Price, tbLift10Floors, lblLift10Total },
                { label36, tbLift11Price, tbLift11Floors, lblLift11Total },
                { label34, tbLift12Price, tbLift12Floors, lblLift12Total }
            };
        }

        private void SetPanelVisabilityDefaults()
        {
            Point rightPanelLocation = new Point(662, 10);
            this.Size = new Size(1127, 959);

            lbWait.Visible = false;

            btnExportQuote.Visible = false;
            btnExportQuote.Enabled = false;

            btnEditQuotePrices.Enabled = false;
            btnEditQuotePrices.Visible = false;
            btnEditQuotePrices.Location = btnGenerateCustomerQuote.Location;

            printButton.Visible = false;
            printButton.Enabled = false;

            PanelDefaultSettings(panelAdditionalCosts, rightPanelLocation);
            PanelDefaultSettings(panelShipping, rightPanelLocation);
            PanelDefaultSettings(panelLiftPrices, rightPanelLocation);
            PanelDefaultSettings(panelPageNumberButtons, new Point(997, 12));
            PanelDefaultSettings(panelExportQuote, new Point(713, 755));
            PanelDefaultSettings(panelContactDetails, new Point(705, 11));

            infoPages = new Panel[] { panelLift1, panelLift2, panelLift3, panelLift4, panelLift5, panelLift6, panelLift7, panelLift8, panelLift9, panelLift10, panelLift11, panelLift12 };
            foreach (Panel p in infoPages)
            {
                PanelDefaultSettings(p, new Point(12, 12));
            }
        }

        private void PanelDefaultSettings(Panel panel, Point location, bool visible = false, bool enabled = false)
        {
            panel.Location = location;
            panel.BringToFront();
            panel.Visible = visible;
            panel.Enabled = enabled;

        }

        private void SetupPricesPanel()
        {
            for (int i = 1; i < 12; i++)
            {
                NewCostsRow(i, false);

            }
            NewCostsRow(0);
            btnAddLiftCostField.Enabled = true;
            btnAddLiftCostField.Visible = true;
            numberOfPagesNeeded = 1;
            //btnAddLiftCostField.Location = new Point(111, 186);

        }

        private void SetLablesToDefault()
        {
            string s = "0";
            lowestMargin = mm.basePrices["17LowestMargin"];
            tbMainMargin.Text = mm.basePrices["18DefaultMargin"].ToString();
            lblCostOfParts.Text = PriceRounding(float.Parse(s));
            lblCostIncludingMargin.Text = PriceRounding(float.Parse(s));
            lblGST.Text = PriceRounding(float.Parse(s));
            lblPriceIncludingGST.Text = PriceRounding(float.Parse(s));
            quoteNumber = mm.quNumber;
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

        private void PageButtonSetup()
        {
            pageButtons = new Button[] { btPanel1, btPanel2, btPanel3, btPanel4, btPanel5, btPanel6, btPanel7, btPanel8, btPanel9, btPanel10, btPanel11, btPanel12 };

            btNewPanel.Location = btPanel2.Location;
            foreach (Button bt in pageButtons)
            {
                bt.Enabled = false;
                bt.Visible = false;
            }
            btPanel1.Enabled = true;
            btPanel1.Visible = true;
            btPanel1.ForeColor = Color.Blue;
        }

        #endregion

        #region Importing Data from XML Files
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

        private void button4_Click(object sender, EventArgs e)
        {
            currencySelectionGroup.Visible = !currencySelectionGroup.Visible;
        }

        //Called with what currency is selected and sets the exchange rate accordingly
        public void SelectCurrency(string selector)
        {
            currencySelectionGroup.Visible = false;
            exchangeRateLbl.Visible = true;
            exchangeRateLbl.Enabled = true;
            lblExchangeDate.Enabled = true;
            lblExchangeDate.Visible = true;

            if (selector == "A")
            {
                applicableExchangeRate = 1;
                exCurrency = 0;
                exchangeRateText = "AUD";
                btnCurrency.Text = "AUD";
                exchangeRateLbl.Text = "No exchange rate is being applied to this form";
                lblExchangeDate.Text = "";
            }
            else if (selector == "U")
            {
                applicableExchangeRate = mm.exchangeRates["USD"];
                exchangeRateLbl.Text = "The current exchange rate is $1 USD to " + PriceRounding(mm.exchangeRates["USD"]) + " AUD";
                lblExchangeDate.Text = "Correct as of " + mm.exchangeRateDate;
                exCurrency = 1;
                exchangeRateText = "USD";
                btnCurrency.Text = "USD";
            }
            else if (selector == "E")
            {
                applicableExchangeRate = mm.exchangeRates["EUR"];
                exchangeRateLbl.Text = "The current exchange rate is €1 EUR to " + PriceRounding(mm.exchangeRates["EUR"]) + " AUD";
                lblExchangeDate.Text = "Correct as of " + mm.exchangeRateDate;
                exCurrency = 2;
                exchangeRateText = "EUR";
                btnCurrency.Text = "EUR";
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
                shippingLbl20.Text = num20Ft + "x 20ft Container(s) - " + PriceRounding(mm.basePrices["20ftFreight"] * num20Ft);
            }
            else if (selector == 3)
            {
                num40Ft++;
                shippingLbl40.Text = num40Ft + "x 40ft Container(s) - " + PriceRounding(mm.basePrices["40ftFreight"] * num40Ft);
            }
            freightTotal = (num20Ft * mm.basePrices["20ftFreight"]) + (num40Ft * mm.basePrices["40ftFreight"]);
            shippingLblTotal.Text = "Total of " + PriceRounding(freightTotal) + " for shipping";
        }

        #endregion

        #region Generating Price List
        //update values in main menu when changed
        private void SetValuesForMainMenuLabels()
        {
            TotalCostAdder();

            lblNumOfLifts.Text = ("Price for " + numberOfPagesNeeded + " lift(s)");
            float f2 = float.Parse(lblTotalLiftPrice.Text) * applicableExchangeRate;
            lblTotaliftCosts.Text = ("Totaling $" + f2 + " AUD");
            lblNumOfShipContainers.Text = ("Price for " + num20Ft + " 20ft and " + num40Ft + " 40ft Containers");
            lblShippingTotal.Text = ("Totaling $" + freightTotal + " AUD");

            float f = TBToFloat(tbMainSundries) + TBToFloat(tbMainBlankets) + TBToFloat(tbMainShaftLight) + TBToFloat(tbMainDuct) +
                TBToFloat(tbMainAccomodation) + TBToFloat(tbMainCartage) + TBToFloat(tbMainStorage) + TBToFloat(tbTravel);
            if (cbMainSecurity.Checked)
            {
                f += mm.basePrices["Security"] + (mm.basePrices["SecurityPerFloor"] * FloorsAdder());
            }
            f += float.Parse(tbMainEntranceGuards.Text) * float.Parse(tbMainWeeksRequired.Text) * mm.basePrices["15EntranceGuards"];
            f += float.Parse(tbMainScaffold.Text) * mm.basePrices["14Scaffolds"];

            lblExtraCostsTotal.Text = ("Totaling $" + f + " AUD");
        }

        private float TBToFloat(TextBox textBoxToBeConverted)
        {
            float f = 0;
            try
            {
                f = float.Parse(textBoxToBeConverted.Text);
            }
            catch (Exception)
            {
                f = 0;
            }
            return f;
        }

        // generate price list button
        private void button1_Click(object sender, EventArgs e)
        {
            PanelMenuChange(null);
            SetValuesForMainMenuLabels();
            TotalCostAdder();
            PagesRequired();
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
                btnExportQuote.Visible = true;
                btnExportQuote.Enabled = true;
                btnGenerateCustomerQuote.Visible = true;
                btnGenerateCustomerQuote.Enabled = true;
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
            PriceListFormatting(lblFinishes, mm.basePrices["1CarFinishes"]);
            PriceListFormatting(lblFire, mm.basePrices["2FireExtinguisher"]);
            PriceListFormatting(lblGSM, mm.basePrices["3GSMUnitPhone"]);
            //PriceListFormatting(lblBlanket, mm.basePrices["4ProtectiveBlanket"]);
            PriceListFormatting(lblBlanket, float.Parse(tbMainBlankets.Text));
            PriceListFormatting(lblSump, mm.basePrices["5SumpCover"]);
            PriceListFormatting(lblWiring, mm.basePrices["6Wiring"]);
            PriceListFormatting(lblSign, mm.basePrices["7Sinage"]);
            PriceListFormatting(lblElectrical, mm.basePrices["8ElectricalBox"]);
            PriceListFormatting(lblCartage, float.Parse(tbMainCartage.Text));
            PriceListFormatting(lblDrawing, mm.basePrices["9Drawings"]);
            PriceListFormatting(lblFork, mm.basePrices["10ForkLift"]);
            PriceListFormatting(lblMaintenance, mm.basePrices["11Maintenance"]);
            PriceListFormatting(lblManuals, mm.basePrices["12Manuals"]);
            PriceListFormatting(lblWorkcover, mm.basePrices["13WorkcoverFees"]);
            PriceListFormatting(lblScaffold, float.Parse(tbMainScaffold.Text) * mm.basePrices["14Scaffolds"]);
            PriceListFormatting(lblEntrance, float.Parse(tbMainEntranceGuards.Text) * float.Parse(tbMainWeeksRequired.Text) * mm.basePrices["15EntranceGuards"]);
            //add security from base prices dictionary if box is checked
            if (cbMainSecurity.Checked)
            {
                PriceListFormatting(lblSecurity, mm.basePrices["Security"] + (mm.basePrices["SecurityPerFloor"] * FloorsAdder()));
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
                //lblLiftNoConvertPrice.Text = PriceRounding(float.Parse(lblTotalLiftPrice.Text), costInEuro);
                lblLiftNoConvertPrice.Text = PriceRounding(float.Parse(lblTotalLiftPrice.Text), costInEuro);
            }
            else if (exCurrency == 2)
            {
                lblLiftNoConvert.Visible = true;
                lblLiftNoConvertPrice.Visible = true;
                lblLiftNoConvert.Text = "Cost of Lift (EUR)";
                costInEuro = true;
                //lblLiftNoConvertPrice.Text = PriceRounding(float.Parse(lblTotalLiftPrice.Text), costInEuro);
                lblLiftNoConvertPrice.Text = PriceRounding(float.Parse(lblTotalLiftPrice.Text), costInEuro);
            }

            //add freight based on number of required containers 
            PriceListFormatting(lblFreight, freightTotal);
            //add labour from the labour costs dictionary based on number of floors in the building 
            PriceListFormatting(lblLabour, LabourAdder());

            marginPercent = 1 + (float.Parse(tbMainMargin.Text) / 100);
            float marginValue = (float.Parse(tbMainMargin.Text) / 100) * liftPrice;
            float roundingAdjustment = Roundingbuffer((liftPrice * marginPercent) * 1.1f);
            //liftPrice *= int.Parse(tbNumberLifts.Text);

            lblCostOfParts.Text = PriceRounding(liftPrice);
            lblCostIncludingMargin.Text = PriceRounding(liftPrice * marginPercent); //+ " (" + PriceRounding(marginValue) + ")";

            //auto rounding methods 
            // lblPriceIncludingGST.Text = PriceRounding(((liftPrice * marginPercent) * 1.1f) + roundingAdjustment);
            //lblGST.Text = PriceRounding((((liftPrice * marginPercent) * 1.1f) + roundingAdjustment) / 11);

            lblPriceIncludingGST.Text = PriceRounding(((liftPrice * marginPercent) * 1.1f) + PlusMinusAdjust());
            lblGST.Text = PriceRounding((((liftPrice * marginPercent) * 1.1f) + PlusMinusAdjust()) / 11);

        }

        private float PlusMinusAdjust()
        {
            float f = 0;
            try
            {
                f = float.Parse(tbMinorPriceAdjustment.Text);
                return f;
            }
            catch (Exception)
            {
                return f;
            }
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

            labour += mm.labourPrice[int.Parse(tbLift1Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift2Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift3Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift4Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift5Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift6Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift7Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift8Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift9Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift10Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift11Floors.Text)];
            labour += mm.labourPrice[int.Parse(tbLift12Floors.Text)];
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

            if (i > mm.maxFloorNumber || i < 0)
            {
                //MessageBox.Show("Invalid floor number entered ");
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

                SaveData();
                QuestionsComplete();
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
            string filePath = @"X:\Program Dependancies\Quote tool\Template Word Docs\Template-" + Environment.UserName + "-Diff-" + numberOfPagesNeeded + ".docx";
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

            try
            {
                //execute find and replace
                fileOpen.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                    ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                    ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            }
            catch (Exception)
            {

            }
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

        public void lblWaitControl(bool enablewait)
        {
            // if called true will enable the wait message and disable the form
            // if called false will disable the wait message and enable the form 

            lbWait.Location = new Point(42, 358);
            lbWait.BringToFront();
            lbWait.Enabled = enablewait;
            lbWait.Visible = enablewait;
            this.Enabled = !enablewait;
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
                document = fileOpen.Documents.Open("X:\\Program Dependancies\\Quote tool\\Template Word Docs\\Template-" + Environment.UserName + "-Price-" + numberOfPagesNeeded + ".docx", ReadOnly: false);
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

        public bool LoadPreviousQuote(string fileToBeLoaded)
        {
            // prompt user to select word doc
            // remove "quote" or "price breakdown" and file extenstion
            // find XML doc with the same name in the dependancies folder
            // call seprate method to load data from that XML into a dictionary 
            // populate form with data from dictionary
            // when opening each question form load all relevent data from the dictionary from Form1

            if (fileToBeLoaded != null)
            {
                saveData.Clear();
                string xmlPath = @"X:\Program Dependancies\Quote tool\Previous Prices\" + fileToBeLoaded;
                //MessageBox.Show(xmlPath);
                loadingPreviousData = true;
                FetchsaveData(xmlPath);
                Form1LoadFromXML();
                GenerateListOfPrices();
                return true;
            }
            else
            {
                MessageBox.Show("Invalid file selected");
            }

            /* Old load method using afile picker, being replaced with the above "list of files" method
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
                    return true;
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
            */
            return false;
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
                 tbLift10Price, tbLift10Floors, tbLift11Price, tbLift11Floors, tbLift12Price, tbLift12Floors, tBMainQuoteNumber);
            V1Fixer(); // corrects the bug with differently named textboxes between multi and single in the v1.0 program that no longer exists in the v2.0 program. 
            this.Text = tBMainQuoteNumber.Text + " Calculation Window";
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

        public void V1Fixer()
        {
            // data meant to be saved in the textboxes from the single lift program in 1.0 is redirected into the correct boxes in the 2.0 version
            // this issue only exists when loading an old quote made under specific circumstances
            // once loaded for the first time in 2.0 this method will repair the error and upon saving anew will correect for all future loads. 
            if (saveData.ContainsKey("tbCost"))
            {
                tbLift1Price.Text = saveData["tbCost"];
            }
            if (saveData.ContainsKey("tBMainFloors"))
            {
                tbLift1Floors.Text = saveData["tBMainFloors"];
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

        private void tbNumofCarEntrances_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void RearDoorChecker(TextBox carEntrance, RadioButton rbYes, RadioButton rbNo)
        {
            try
            {
                if (int.Parse(carEntrance.Text) >= 2)
                {
                    rbYes.Checked = true;
                }
                else
                {
                    rbNo.Checked = true;
                }
                //if try fails the bool will remain false and thus be able to try again
                // if try is sucessful it will change the bool to true and prevent additional edits
                rearDoorChecker = true;
            }
            catch (Exception)
            {
                return;
            }
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

        private void btnAddLiftCostField_Click(object sender, EventArgs e)
        {
            NewCostsRow(numberOfPagesNeeded);
        }

        private void NewCostsRow(int costRow, bool Visibility = true)
        {
            Label l;
            TextBox tb;

            for (int i = 0; i < 4; i++)
            {
                if (costsGroups[costRow, i] is Label)
                {
                    l = (Label)costsGroups[costRow, i];

                    l.Visible = Visibility;
                    l.Enabled = Visibility;
                }
                else if (costsGroups[costRow, i] is TextBox)
                {
                    tb = (TextBox)costsGroups[costRow, i];

                    tb.Visible = Visibility;
                    tb.Enabled = Visibility;
                }
            }

            if (costRow < 11)
            {
                tb = (TextBox)costsGroups[costRow + 1, 1];
                btnAddLiftCostField.Location = tb.Location;
            }
            else
            {
                btnAddLiftCostField.Enabled = false;
                btnAddLiftCostField.Visible = false;
            }

            numberOfPagesNeeded++;
        }

        private void HideButton_Click(object sender, EventArgs e)
        {
            //
        }

        private void RefreshLiftPrices(TextBox price, Label total)
        {
            try
            {
                total.Text = (float.Parse(price.Text)).ToString();
            }
            catch (Exception)
            {
                //
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
            float f = 0f;
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

        #region Multi Lift Exporter Pages
        private void btNewPanel_Click(object sender, EventArgs e)
        {
            NewPage();
        }

        private void btPanel1_Click(object sender, EventArgs e)
        {
            OpenInfoPage(0);
        }

        private void btPanel2_Click(object sender, EventArgs e)
        {
            OpenInfoPage(1);
        }

        private void btPanel3_Click(object sender, EventArgs e)
        {
            OpenInfoPage(2);
        }

        private void btPanel4_Click(object sender, EventArgs e)
        {
            OpenInfoPage(3);
        }

        private void btPanel5_Click(object sender, EventArgs e)
        {
            OpenInfoPage(4);
        }

        private void btPanel6_Click(object sender, EventArgs e)
        {
            OpenInfoPage(5);
        }

        private void btPanel7_Click(object sender, EventArgs e)
        {
            OpenInfoPage(6);
        }

        private void btPanel8_Click(object sender, EventArgs e)
        {
            OpenInfoPage(7);
        }

        private void btPanel9_Click(object sender, EventArgs e)
        {
            OpenInfoPage(8);
        }

        private void btPanel10_Click(object sender, EventArgs e)
        {
            OpenInfoPage(9);
        }

        private void btPanel11_Click(object sender, EventArgs e)
        {
            OpenInfoPage(10);
        }

        private void btPanel12_Click(object sender, EventArgs e)
        {
            OpenInfoPage(11);
        }

        private void NewPage()
        {
            if (numberOfPagesUsed >= numberOfPagesNeeded)
            {
                return;
            }

            pageButtons[pageTracker].Visible = true;
            pageButtons[pageTracker].Enabled = true;
            pageTracker++;
            exporterPageOpened[pageTracker] = true;

            if (pageTracker <= 11)
            {
                btNewPanel.Location = pageButtons[pageTracker].Location;
            }
            else
            {
                btNewPanel.Visible = false;
                btNewPanel.Enabled = false;
            }
        }

        private void OpenInfoPage(int pageToOpen)
        {
            if (pageToOpen == activePage)
            {
                return;
            }
            else
            {
                infoPages[pageToOpen].Visible = true;
                infoPages[pageToOpen].Enabled = true;
                pageButtons[pageToOpen].ForeColor = System.Drawing.Color.Blue;
                infoPages[activePage].Visible = false;
                infoPages[activePage].Enabled = false;
                pageButtons[activePage].ForeColor = System.Drawing.Color.Black;
                activePage = pageToOpen;
            }
        }

        #endregion

        #region Fill data from 1st page into additional pages
        private void FillPageWithMainData(int pageToBeFilled, ref bool pageOpenedPreviously, bool previousQuoteWasLoaded)
        {
            if (pageOpenedPreviously || previousQuoteWasLoaded)
            {
                return;
            }

            pageToBeFilled--;
            pageOpenedPreviously = true;

            #region Page Data Vars
            Object[][] pageObjects = new object[][]
            {

           new Object[] { tbAuxCOPLocation, tbDepth, tbCarDoorFinish, tbHeight, tbLoad, tbwidth, tbCeilingFinish,
                tbControlerLocation, tbCOPFinish, tbDesignations, tbDoorHeight, tbDoorTracks, tbDoorWidth, tbFacePlateMaterial,
                tbFloorFinish, tbFrontWall, tbHandrail, tbHeadroom, tbKeyswitchLocation, tbLandingDoorFinish, tbLiftNumbers,
                tbLiftRating, tbMainCOPLocation, tbMirror, tbLiftCarNotes, tbNumofCarEntrances, tbNumberOfCOPS, tbNumofLandingDoors,
                tbNumofLandings, tbNumOfLEDLights, tbPitDepth, tbRearWall, tbShaftDepth, tbShaftWidth, tbSideWall, tbSpeed,
                tbStructureShaft, tbTravel, tbTypeofLift, rbAdvancedOpeningNo, rbAdvancedOpeningYes, rbBumpRailNo, rbBumpRailYes,
                rbCarDoorFinishOther, rbCarDoorFInishBrushedStainlessSteel, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther,
                rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishWhite, rbControlerLocationBottomLanding, rbControlerLocationOther,
                rbControlerlocationShaft, rbControlerLoactionTopLanding, rbCOPFinishOther, rbCOPFinishSatinStainlessSteel, rbDoorNudgingNo,
                rbDoorNudgingYes, rbDoorTracksAnodisedAluminium, rbDoorTracksOther, rbDoorTypeCentreOpening, rbDoorTypeSideOpening,
                rbEmergencyLoweringSystemNo, rbEmergencyLoweringSystemYes, rbExclusiveServiceNo, rbExclusiveServiceYes,
                rbFacePlateMaterialOther, rbFacePlateMaterialSatinStainlessSteel, rbFalseCeilingNo, rbFalseCeilingYes, rbFalseFloorNo, rbFalseFloorYes,
                rbFireServiceNo, rbFireServiceYes, tbFrontWallOther, rbFrontWallBrushedStainlessSteel, rbGPOInCarNo, rbGPOInCarYes, rbHandrailOther,
                rbHandrailBrushedStainlessSTeel, rbIndependentServiceNo, rbIndependentServiceYes, rbLandingDoorFinishOther,
                rbLandingDoorFinishStainlessSteel, rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite, rbLoadWeighingNo, rbLoadWeighingYes,
                rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther, rbOutofServiceNo, rbOutofServiceYes, rbPositionIndicatorTypeFlushMount,
                rbPositionIndicatorTypeSurfaceMount, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes, rbRearDoorKeySwitchNo,
                rbRearDoorKeySwitchYes, rbRearWallOther, rbRearWallBrushedStainlessSteel, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes,
                rbSideWallOther, rbSideWallBrushedStainlessSteel, rbStructureShaftConcrete, rbStructureShaftOther, rbTrimmerBeamsNo,
                rbTrimmerBeamsYes, rbVoiceAnnunciationNo, rbVoiceAnnunciationYes },

                 new Object[]  { tb2AuxCOPLocation, tb2CarDepth, tb2CarDoorFinishText, tb2CarHeight, tb2CarLoad, tb2CarWidth, tb2CeilingFinishText,
                    tb2ControlerLocationText, tb2COPFinishText, tb2Designations, tb2DoorHeight, tb2DoorTracksText, tb2DoorWidth, tb2FacePlateMaterialText,
                    tb2FloorFinish, tb2FrontWallText, tb2HandrailText, tb2Headroom, tb2KeyswitchLocation, tb2LandingDoorFinishText, tb2LiftNumbers,
                    tb2LiftRating, tb2MainCOPLocation, tb2MirrorText, tb2Note, tb2NumberOfCarEntrances, tb2NumberOfCOPs, tb2NumberOfLAndingDoors,
                    tb2NumberOfLandings, tb2NumberofLEDLights, tb2PitDepth, tb2RearWallText, tb2ShaftDepth, tb2ShaftWidth, tb2SideWallText, tb2Speed,
                    tb2StructureShaftText, tb2Travel, tb2TypeOfLift, rb2AdvancedOpeningNo, rb2AdvancedOpeningYes, rb2BumpRailNo, rb2BumpRailYes,
                    rb2CarDoorFinishOther, rb2CarDoorFinishStainlessSteel, rb2CeilingFinishMirrorStainlessSteel, rb2CeilingFinishOther,
                    rb2CeilingFinishStainlessSteel, rb2CeilingFinishWhite, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther,
                    rb2ControlerLocationShaft, rb2ControlerLocationTopLanding, rb2COPFinishOther, rb2COPFinishStainlessSTeell, rb2DoorNudgingNo,
                    rb2DoorNudgingYes, rb2DoorTracksAluminium, rb2DoorTracksOther, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening,
                    rb2EmergemncyLoweringSystemNo, rb2EmergencyLoweringSystemYes, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes,
                    rb2FacePlateMaterialOther, rb2FacePlateMaterialStainlessSteel, rb2FalseCeilingNo, rb2FalseCeilingYes, rb2FalseFloorNo, rb2FalseFloorYes,
                    rb2FireSErviceNo, rb2FireSErviceYes, rb2FrontWallOther, rb2FrontWallStainlessSteel, rb2GPOInCarNo, rb2GPOInCarYes, rb2HandrailOther,
                    rb2HandRailStainlessSteel, rb2IndependentServiceNo, rb2IndependentServiceYes, rb2LandingDoorFinishOther,
                    rb2LandingDoorFinishStainlessSteel, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite, rb2LoadWeighingNo, rb2LoadWeighingYes,
                    rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther, rb2OutOfServiceNo, rb2OutOfServiceYes, rb2PositionIndicatorTypeFlushMount,
                    rb2PositionIndicatorTypeSurfaceMount, rb2ProtectiveBlanketsNo, rb2ProtectiveBlanketsYes, rb2RearDoorKeySwitchNo,
                    rb2RearDoorKeySwitchYes, rb2RearWallOther, rb2RearWallStainlessSteel, rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes,
                    rb2SideWallOther, rb2SideWallStainlessSteel, rb2StructureShaftConcrete, rb2StructureShaftOther, rb2TrimmerBeamsNo,
                    rb2TrimmerBeamsYes, rb2VoiceAnnunciationNo, rb2VoiceAnnunciationYes },

           new Object[] { tb3AuxCOPLocation, tb3CarDepth, tb3CarDoorFinishText, tb3CarHeight, tb3Load, tb3CarWidth, tb3CEilingFinishText,
                tb3ControlerLocationText, tb3COPFinishText, tb3Designations, tb3DoorHeight, tb3DoorTracksText, tb3DoorWidth, tb3FacePlaterMaterialText,
                tb3FloorFinish, tb3FrontWallText, tb3HandrailText, tb3HeadRoom, tb3KeyswitchLocation, tb3LandingDoorFinishText, tb3LiftNumbers,
                tb3LiftRating, tb3MainCOPLocation, tb3MirrorText, tb3CarNote, tb3NumberOfCarEntrances, tb3NumberOfCOPs, tb3NumberOfLandingDoors,
                tb3NumberOfLandings, tb3NumberOfLEDLights, tb3PitDepth, tb3RearWallText, tb3ShaftDepth, tb3ShaftWidth, tb3SideWallText, tb3Speed,
                tb3StructureShaftText, tb3Travel, tb3TypeOfLift, rb3AdvancedOpeningNo, rb3AdvancedOpeningYes, rb3BumpRailNo, rb3BumpRailYes,
                rb3CarDoorFinishOther, rb3CarDoorFinishStainlessSteel, rb3MirrorStainlessSteel, rb3CeilingFinishOther,
                rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite, rb3ControleRLocationBottomLanding, rb3ControlerLocationOther,
                rb3ControlerLocationShaft, rb3ControlerLocationTopLanding, rb3COPFinishOther, rb3COPFinishStainlessSteel, rb3DoorNudgingNo,
                rb3DoorNudgingYes, rb3DoorTracksAluminium, rb3DoorTracksOther, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening,
                rb3EmergencyLoweringSystemNo, rb3EmergencyLoweringSystemYes, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes,
                rb3FacePlateMaterialOther, rb3FacePlateMaterialStainlessSteel, rb3FalseCeilingNo, rb3FalseCeilingYes, rb3FalseFloorNo, rb3FalseFloorYes,
                rb3FireServieNo, rb3FireServiceYes, rb3FrontWallOther, rb3FrontWallStainlessSteel, rb3GPOInCarNo, rb3GPOInCarYes, rb3HandrailOther,
                rb3HandrailStainlessSteel, rb3IndependentServiceNo, rb3IndependentServiceYes, rb3landingDoorFinishOther,
                rb3LandingDoorFinishStainlessSteel, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite, rb3LoadWeighingNo, rb3LoadWeighingYes,
                rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther, rb3OutOfServiceNo, rb3OutOfSErviceYes, rb3PositionIndicatorTypeFlushMount,
                rb3PositionIndicatorTypeSurfaceMount, rb3ProtectiveBlanketsNo, rb3ProtectiveBlanketsYes, rb3RearDoorKeySwitchNo,
                rb3RearDoorKeySwitchYes, rb3RearWallOther, rb3RearWallStainlessSteel, rb3SecurityKeySwitchNo, rb3SecurityKeySwitchYes,
                rb3SideWallOther, rb3SideWallStainlessSteel, rb3StructureShaftConcrete, rb3StructureShaftOther, rb3TrimmerBeamsNo,
                rb3TrimmerBeamsYes, rb3VoiceAnnunciationNo, rb3VoiceAnnunciationYes },

           new Object[] { tb4AuxCOPLocation, tb4CarDepth, tb4CarDoorFinish, tb4CarHeight, tb4Load, tb4CarWidth, tb4CeilingFinishText,
                tb4ControlerLocationText, tb4COPFinishText, tb4Designations, tb4DoorHeight, tb4DoorTracksText, tb4DoorWidth, tb4FacePlateMaterialText,
                tb4FloorFinish, tb4FrontWallText, tb4HandrailText, tb4Headroom, tb4KeyswitchLocations, tb4LandingDoorFinishText, tb4LiftNumbers,
                tb4LiftRating, tb4MainCOPLocation, tb4MirrorText, tb4CarNote, tb4NumberOfCarEntrances, tb4NumberOfCOPs, tb4NumberOfLandingDoors,
                tb4NumberOfLandings, tb4NumbeROfLEDLights, tb4PitDepth, tb4RearWallText, tb4ShaftDepth, tb4ShaftWidth, tb4SideWallText, tb4Speed,
                tb4StructureShaftText, tb4Travel, tb4TypeOfLift, rb4AdvancedOpeningNo, rb4AdvancedOpeningYes, rb4BumpRailNo, rb4BumpRailYes,
                rb4CarDoorFinishOther, rb4CarDoorFinishStainlessSteel, rb4CeilingFinishMirrorStainlessSteel, rb4CeilingFinishOther,
                rb4CeilingFinishStainlessSteel, rb4CeilingFinishWhite, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther,
                rb4ControlerLocationShaft, rb4ControelrLocationTopLanding, rb4COPFinishOther, rb4COPFinishStainlessSteel, rb4DoorNudgingNo,
                rb4DoorNudgingYes, rb4DoorTracksAluminium, rb4DoorTracksOther, rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening,
                rb4EmergencyLoweringSystemNo, rb4EmergencyLoweringSystemYes, rb4ExclusiveServiceNo, rb4ExclusiveServiceYes,
                rb4FacePlateMaterialOther, rb4FacePlateMaterialStainlessSteel, rb4FalseCeilingNO, rb4FalseCeilingYes, rb4FalseFloorNo, rb4FalseFloorYes,
                rb4FireSErviceNo, rb4FireServiceYes, rb4FrotnWallOther, rb4FrontWallStainlessSteel, rb4GPOInCarNo, rb4GPOInCarYes, rb4HandrailOther,
                rb4HandRailStainlesSteel, IndependentServiceNo, rb4IndependentServiceYes, rb4LandingDoorFinishOther,
                rb4LandingDoorFinishStainlessSteel, rb4LCDColourBlue, rb4LCDColourRed, rb4LCDColourWhite, rb4LoadWeighingNo, rb4LoadWeighingYes,
                rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther, rb4OutOfServiceNo, rb4OutOfServiceYes, rb4PositionIndicatorTypeFlushMount,
                rb4PositionIndicatorTypeSurfaceMount, rb4ProtetiveBlanketsNo, rb4ProtectiveBlanketsYes, rb4RearDoorKeySwitchNo,
                rb4RearDoorKeySwitchYes, rb4RearWallOther, rb4RearWallStainlessSteel, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes,
                rb4SideWallOther, rb4SideWallStainlessSteel, rb4StructureShaftConcrete, rb4StructureShaftOther, rb4TrimmerBeamsNo,
                rb4TrimmerBeamsYes, rb4VoiceAnnunciationNo, rb4VoiceAnnunciationYes },

           new Object[] { tb5AuxCOPLocation, tb5CarDepth, tb5CarDoorFinishText, tb5CaRHeight, tb5Load, tb5CarWidth, tb5CeilingFinishText,
                tb5ControlerLocationText, tb5COPFinishText, tb5Designations, tb5DoorHeight, tb5DoorTRacksText, tb5DoorWidth, tb5FacePlateMaterialText,
                tb5FloorFinish, tb5FrontWallText, tb5HandrailText, tb5Headroom, tb5KetyswitchLocation, tb5LandingDoorFinishText, tb5LiftNumbers,
                tb5LiftRating, tb5MainCOPLocation, tb5MirrorText, tb5CarNote, tb5NumberOfCarEntrances, tb5NumberOfCOPs, tb5NumberOfLandingDoors,
                tb5NumberOfLandings, tb5NumberOIfLEDLights, tb5PitDepth, tb5RearWallText, tb5ShaftDEpth, tb5ShaftWidth, tb5SideWallText, tb5Speed,
                tb5StructureShaftText, tb5Travel, tb5TypeOfLift, rb5AdvancedOpeningNo, rb5AdvancedOpeningYes, rb5BumpRailNo, rb5BumpRailYes,
                rb5CarDoorFinishOther, rb5CarDoorFinishStainlessSteel, rb5CeilingFinishMirrorStainlessSTeel, rb5CeilingFinishOther,
                rb5CeilingFinishStainlessSteel, rb5CeilingFinishWhite, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther,
                rb5ControlerLocationShaft, rb5ControlerLocationTopLanding, rb5COPFinishOther, rb5COPFinishStainlessSteel, rb5DoorNudgingNo,
                rb5DoorNudgingYes, rb5DoorTracksAluminium, rb5DoorTracksOther, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening,
                rb5EmergencyLoweringSystemNo, rb5EmergencyLoweringSystemYes, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes,
                rb5FacePlateMAterialOtjer, rb5FacePlateMaterialStainlessSteel, rb5FalseCeilingNo, rb5FalseCeilingYes, rb5FalseFloorNo, rb5FalseFloorYes,
                rb5FireSErviceNo, rb5FireServiceYes, rb5FrontWallOther, rb5FrontWallStainlessSteel, rb5GPOInCarNo, rb5GPOInCarYes, rb5HandRailOther,
                rb5HandRailStainlessSteel, rb5IndependentServiceNo, rb5IndependentServiceYes, rb5LAndingDoorFinishOther,
                rb5LandingDoorFinishStainlessSteel, rb5LCDColourBlue, rb5LCDColourRed, rb5LCDColoiurWHite, rb5LoadWeighingNo, rb5LoadWeighingYes,
                rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther, rb5OutOfServiceNo, rb5OutOfServiceYes, rb5PostionIndicatorTypeFlushMount,
                rb5PositionIndicatorTypeSurfaceMount, rb5ProtectiveBlanketsNo, rb5ProtectiveBlanketsYes, rb5RearDoorKeySwitchNo,
                rb5RearDoorKeySwitchYes, rb5RearWallOther, rb5RearWallStainlesSteel, rb5SecurityKeySwitchNo, rb5SecurityServiceYes,
                rb5SideWallOther, rb5SideWallStainlessSTeel, rb5StructureShaftConcrete, rb5StructureShaftOther, rb5TrimmerBeamsNo,
                rb5TrimmerBeamsYes, rb5VoiceAnnunciationNo, rb5VoiceAnnunciationYes },

           new Object[] { tb6AuxCOPLocation, tb6CarDepth, tb6CarDoorFinishText, tb6CarHeight, tb6CarLoad, tb6CarWidth, tb6CeilingFinishText,
                tb6ControlerLocationText, tb6COPFinishText, tb6Designations, tb6DoorHeight, tb6DoorTracksOther, tb6DoorWidth, tb6FacePlateMaterialText,
                tb6FloorFinish, tb6FrontWallText, tb6HAndrailText, tb6Headroom, tb6KeySwitchLocation, tb6LandingDoorFinishText, tb6LiftNumbers,
                tb6LiftRating, tb6MainCOPLocation, tb6MirrorText, tb6CarNote, tb6NumberOfCarEntrances, tb6NumberOFCOPs, tb6NumberOfLandingDoors,
                tb6NumberOfLandings, tb6NumberOfLEDLights, tb6PitDepth, tb6RearWallText, tb6ShaftDepth, tb6ShaftWidth, tb6SideWallText, tb6CarSpeed,
                tb6StructureShaftText, tb6Travel, tb6TypeOfLift, rb6AdvancedOpeningNo, rb6AdvancedOpeningYes, rb6BumpRailNo, rb6BumpRailYes,
                rb6CarDoorFinishOther, rb6CarDoorFinishStainlessSteel, rb6CeilingFinishMirrorStainlessSteel, rb6CeilingFinishOther,
                rb6CeilingFinishStainlessSteel, rb6CeilingFinishWhite, rb6ControlerLocationBottomLanding, rb6ControlerLocationOther,
                rb6ControlerLocationShaft, rb6ControlerLocationTopLanding, rb6COPFinishOther, rb6COPFinishStainlessSteel, rb6DoorNudgingNo,
                rb6DoorNudgingYes, rb6DoorTracksAluminium, rb6DoorTracksOther, rb6DoorTypeCentreOpening, rb6DoorTypeSideOpening,
                rb6EmergencyLoweringSystemNo, rb6EmergencyLoweringSystemYes, rb6ExclusiveServiceNo, rb6ExclusiveServiceYes,
                rb6FacePlateMaterialOther, rb6FacvePlateMaterialStainlessSteel, rb6FalseCeilingNo, rb6FalseCeilingYes, rb6FalseFloorNo, rb6FalseFloorYes,
                rb6FireServiceNo, rb6FireServiceYes, rb6FrontWallOther, rb6FrontWallStainlessSteel, rb6GPOInCarNo, rb6GPOInCarYes, rb6HandrailOther,
                rb6HandrailStainlessSteel, rb6IndependentNo, rb6IndependentServiceYes, rb6LandingDoorFinishOther,
                rb6LandingDoorFinishStainlessSteel, rb6LCDColourBlue, rb6LCDColourRed, rb6LCDColourWhite, rb6LoadWeighingNo, rb6LoadWeighingYes,
                rb6MirrorFullSize, rb6MirrorHalfSize, rb6MirrorOther, rb6OutOfServiceNo, rb6OutOfServiceYes, rb6PositionIndicatorTypeFlushMount,
                rb6PositionIndicatorTypeSurfaceMount, rb6ProtectiveBlanketsNo, rb6ProtectiveBlanketsYes, rb6RearDoorKeySwitchNo,
                rb6RearDoorKeySwitchYes, rb6RearWallOther, rb6RearWallStainlessSteel, rb6SecurityKeySwitchNo, rb6SecurityKeySwitchYes,
                rb6SideWallOther, rb6SideWallStainlessSteel, rb6StructureShaftCOncrete, rb6StructureShaftOther, rb6TrimmerBeamsNo,
                rb6TrimmerBeamsYes, rb6VoiceAnnunciationNo, rb6VoiceAnnunciationYes },

           new Object[] { tb7AuzCOPLocation, tb7CarDepth, tb7CarDoorFinishText, tb7CarHeight, tb7CarLoad, tb7CarWidth, tb7CEilingFinishText,
                tb7ControlerLocationText, tb7COPFinishText, tb7Designations, tb7DoorHeight, tb7DoorTracksText, tb7DoorWidth, tb7FacePlateMaterialText,
                tb7FloorFinish, tb7FrontWallText, tb7HandrailText, tb7HeadRoom, tb7KeyswitchLocation, tb7LandingDoorFinishText, tb7LiftNumbers,
                tb7LiftRating, tb7MainCOPLocation, tb7MirrorText, tb7CarNotes, tb7NumberOfCarEntrances, tb7NumberOfCOPs, tb7NumberOfLandingDoors,
                tb7NumberOfLandings, tb7NumberOfLEDLights, tb7PitDepth, tb7RearWallText, tb7ShaftDepth, tb7ShaftWidth, tb7SideWallText, tb7CarSpeed,
                tb7StructureShaftText, tb7Travel, tb7TypeOfLift, rb7AdvancedOpeningNo, rb7AdvancedOpeningYes, rb7BumpRailNo, rb7BumpRailYes,
                rb7CarDoorFinishOther, rb7CarDoorFinishStainlessSteel, rb7CeilingFinishMirrorStainlessSteel, rb7CeilingFinishOther,
                rb7CeilingFinishStainlessSteel, rb7CEilingFinishWhite, rb7ControlerLocationBottomLAnding, rb7ControlerLocationOther,
                rb7ControlerLocationShaft, rb7ControlerLocationTopLanding, rb7COPFinishOther, rb7COPFinishStainlessSteel, rb7DoorNudgingNo,
                rb7DoorNudgingYes, rb7DoorTracksAluminium, rb7DoorTracksOther, rb7DoorTypeCentreOpening, rb7DoorTypeSideOpening,
                rb7EmergencyLoweringSystemNo, rb7EmergencyLoweringSystemYes, rb7ExclusiveServiceNo, rb7ExclusiveServiceYes,
                rb7FacePlateMaterialOther, rb7FacePlateMaterialStainlessSteel, rb7FalseCeilingNo, rb7FalseCeilingYes, rb7FalseFloorNo, rb7FalseFloorYes,
                rb7FireServiceNo, rb7FireSErviceYes, rb7FrontWallOther, rb7FrontWallStainlessSteel, rb7GPOInCarNo, rb7GPOInCarYes, rb7HandrailOther,
                rb7HandrailStainlessSteel, rb7IndependentServiceNo, rb7IndpendentServiceYes, rb7LandingDoorFinishOther,
                rb7LandingDoorFinishStainlessSteel, rb7LCDColourBlue, rb7LCDColourRed, rb7LCDColourWhite, rb7LoadWeighingNo, rb7LoadWeighingYes,
                rb7MirrorFullSize, rb7MirrorHalfSize, rb7MirrorOther, rb7OutOfSErviceNo, rb7OutOfServiceYes, rb7PositionIndicatorTypeFlushMount,
                rb7PositionIndicatorTypeSurfaceMount, rb7ProtctiveBlanketsNo, rb7ProtectiveBlanketsYes, rb7RearDoorKeySwitchNo,
                rb7RearDoorKeySwitchYes, rb7RearWallOther, rb7RearWallStainlessSteel, rb7SecurityKeySwitchNo, rb7SecurityKeySwitchYes,
                rb7SideWallOther, rb7SideWallStainlessSteel, rb7StructureShaftConcrete, rb7StructureShaftOther, rb7TrimmerBeamsNo,
                rb7TrimmerBeamsYes, rb7VoiceAnunciationNo, rb7VoiceAnnunciationYes },

           new Object[] { tb8AuxCOPLocation, tb8CarDEpth, tb8CarDoorFinishText, tb8CarHeight, tb8Load, tb8CarWidth, tb8CeilingFinishText,
                tb8ControlerLocationText, tb8COPFinishText, tb8Desiginations, tb8DoorHeight, tb8DoorTracksText, tb8DoorWidth, tb8FacePlateMaterialText,
                tb8FloorFinish, tb8FrontWallText, tb8HandrailText, tb8Headroom, tb8KeyswitchLocations, tb8LandingDoorFinishText, tb8LiftNumbers,
                tb8LiftRating, tb8MainCOPLocation, tb8MirrorText, tb8LiftCarNotes, tb8NumberOfCarEntrances, tb8NumberOfCOPs, tb8NumberOfLandingDoors,
                tb8NumberOfLandings, tb8NumberofLEDLights, tb8PitDepth, tb8RearWallText, tb8ShaftDepth, tb8ShaftWidth, tb8SideWallText, tb8Speed,
                tb8StructureShaftText, tb8Travel, tb8TypeOfLift, rb8AdvncedOpeningNo, rb8AdvancedOpeningYes, rb8BumpRailNo, rb8BumpRailYes,
                rb8CarDoorFinishOther, rb8CarDoorFinishStainlessSteel, rb8CeilingFinishMirrorStainlessSTeel, rb8CeilingFinishOther,
                rb8CeilingFinishStainlessSteel, rb8CeilingFinishWhite, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther,
                rb8ControlerLocationShaft, rb8ControlerLocationTopLanding, rb8COPFinishOther, rb8COPFinishStainlessSteel, tb8DoorNudgingNo,
                rb8DoorNudgingYes, rb8DoorTracksAluminium, rb8DoorTracksOther, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening,
                rb8EmergencyLoweringSystemNo, rb8EmergencyLoweringSystemYes, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes,
                rb8FacePlateMaterialOther, rb8FacePlateMaterialStainlessSteel, rb8FalseCeilingNo, rb8FalseCeilingYes, rb8FalseFloorNo, rb8FalseFloorYes,
                rb8FireSErviceNo, rb8FireServiceYes, rb8FrontWallOther, rb8FrontWallStainlessSteel, rb8GPOInCarNo, rb8GPOInCarYes, rb8HandrailOther,
                rb8HandRailStainlessSteel, rb8IndependentServiceNo, rb8IndependentServiceYes, rb8LandingDoorFinishOther,
                tb8LandingDoorFinishStainlessSteel, rb8LCDColourBlue, rb8LCDColourRed, rb8LCDColourWhite, rb8LoadWeighingNo, rb8LoadWeighingYes,
                rb8MirrorFullSize, rb8MirrorHalfSize, rb8MirrorOther, rb8OutOFServiceNo, rb8OutOfSErviceYes, rb8PositionIndicatorTypeFlushMount,
                rb8PositionIndicatorTypeSurfaceMount, rb8ProtectiveBlanketsNo, rb8ProtectiveBlanketsYes, rb8RearDoorKeySwitchNo,
                rb8RearDoorKeySwitchYes, rb8RearWallOther, rb8RearWallStainlessSteel, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes,
                rb8SideWallOther, rb8SideWallStainlessSteel, rb8StructureShaftConcrete, rb8StructureShaftOther, rb8TrimmerBeamsNo,
                rbTrimmerBeamsYes, rb8VoiceAnnunicationNo, rb8VoiceAnnunicationYes },

           new Object[] { tb9AuxCOPLocation, tb9CarDepth, tb9CarDoorFinishText, tb9CarHeight, tb9Load, tb9CarWidth, tb9CeilingFinishText,
                tb9ControlerLocationText, tb9COPFinishText, tb9Designations, tb9DoorHeight, tb9DoorTracksText, tb9DoorWidth, tb9FacePlateMaterialText,
                tb9FloorFinish, tb9FrontWallText, tb9HandrailTexrt, tb9Headroom, tb9KeyswitchLocation, tb9LandingDoorFinishText, tb9LiftNumbers,
                tb9LiftRating, tb9MainCOPLocation, tb9MirrorText, tb9CarNotes, tb9NumberOFCarEntraces, tb9NumberOfCOPs, tb9NumberOfLandingDoors,
                tb9NumberOfLandings, tb9NumberOfLEDLights, tb9PitDepth, tb9RearWallText, tb9ShaftDepth, tb9ShaftWidth, tb9SideWallText, tb9Speed,
                tb9StructureShaftText, tb9Travel, tb9TypeOfLift, rb9AdvancedOpeningNo, rb9AdvancedOpeningYes, rb9BumpRailNo, rb9BumpRailYes,
                rb9CarDoorFinishOther, rb9CarDoorFinishStainlessSteel, rb9CeilingFinishMirrorStainlessSteel, rb9CeilingFinishOther,
                rb9CeilingFinishStainlessSteel, rb9CeilingFinishWhite, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther,
                rb9ControlerLocationShaft, rb9ControlrLocationTopLanding, rb9COPFinishOther, rb9COPFinishStainlessSteel, rb9DoorNudgingNo,
                rb9DoorNudgingYes, rb9DoorTracksAluminium, rb9DoorTracksOther, rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening,
                rb9EmergencyLoweringSystemNo, rb9EmergencyLoweringSystemYes, rb9ExclusiveServiceNo, rb9ExclusiveServiceYes,
                rb9FacePlateMaterialOther, rb9FacePlateMaterialStainlessSteel, rb9FalseCeilingNo, rb9FalseCeilingYes, rb9FalseFloorNo, rb9FalseFloorYes,
                rb9FireServiceNo, rb9FireSErviceYes, rb9FrontWallOther, rb9FrontWallStainlessSteel, rb9GPOInCarNo, rb9GPOInCarYes, rb9HandrailOther,
                rb9HandrailStainlessSteel, rb9IndependentServiceNo, rb9IndependentServiceYes, rb9LandingDoorFinishOther,
                rb9LandingDoorFinishStainlessSteel, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite, rb9LoadWeighingNo, rb9LoadWeighingYes,
                rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther, rb9OutOfServiceNo, rb9OutOfServiceYes, rb9PositionIndicatorTypeFlushMount,
                rb9PositionIndicatorTypeSurfaceMount, rb9ProtectiveBlanketsNo, rb9ProtectiveBlanketsYes, rb9RearDoorKeySwitchNo,
                rb9RearDoorKeySwitchYes, rb9RearWallOther, rb9RearWallStainlessSteel, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes,
                rb9SideWallOther, rb9SideWallStainlessSteel, rb9StructureShaftConcrete, rb9StructureShaftOther, rb9TrimmerBeamsNo,
                rb9TrimmerBeamsYes, rb9VoiceAnnunciationNo, rb9VoiceAnnunciationYes },

           new Object[] { tb10AuxCOPLocation, tb10CarDepth, tb10CarDoorFinishText, tb10CarHeight, tb10LiftCarLoad, tb10CarWidth, tb10CEilingFinishText,
                tb10ControlerLocationText, tb10COPFinishText, tb10Desigination, tb10DoorHeight, tb10DoorTracksText, tb10DoorWidth, tb10FacePlateMaterialText,
                tb10FloorFinish, tb10FrontWallText, tb10HandrailText, tb10Headroom, tb10KeyswitchLocation, tb10LandingDoorFinishText, tb10LiftNumbers,
                tb10LiftRating, tb10MainCOPLocation, tb10MirrorText, tb10LiftCarNotes, tb10NumberofCarEntrances, tb10NumberOfCOPs, tb10NumberofLandingDoors,
                tb10NumberofLandings, tb10NumberOfLEDLIghts, tb10PitDepth, tb10RearWallText, tb10ShaftDepth, tb10ShaftWidth, tb10SideWallText, tb10Speed,
                tb10StructureShaftText, tb10Travel, tb10TypeOfLift, rb10AdvancedOpeningNo, rb10AdvancedOpeningYes, rb10BumpRailNo, rb10BumpRaidYes,
                rb10CarDoorFinishOther, rb10CarDoorFinishStainlessSteel, rb10CEilingFinishMirrorStainlessSteel, rb10CEilingFinishOther,
                rb10CeilingFinishStainlessSteel, rb10CeilingFinishWhite, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther,
                rb10ControlerLocationShaft, rb10ControlerLocationTopLanding, rb10COPFinishOther, rb10COPFinishStainlessSteel, rb10DoorNudgingNo,
                rb10DoorNudgingYes, rb10DoorTracksAluminium, rb10DoorTracksOther, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening,
                rb10EmergencyLoweringSystemNo, rb10EmergencyLoweringSystemYes, rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes,
                rb10FacePlateMaterialOther, rb10FacePlateMaterialStainlessSteel, rb10FalseCeilingNo, rb10FalseCeilingYes, rb10FalseFloorNo, rb10FalseFloorYes,
                rb10FireSERviceNo, rb10FireSErviceYes, rb10FrontWallOther, rb10FrontWallStainlessSteel, rb10GPOInCarNo, rb10GPOInCarYes, rb10HandrailOther,
                rb10HandrailStainlessSteel, rb10IndependentServiceNo, rb10IndependentServiceYes, rb10LAndingDoorFinishOtherr,
                rb10LandingDoorFinishStainlessSteel, rb10LCDColourBlue, rb10LCDColourRed, rb10LCDColourWhite, rb10LoadWEighingNo, rb10LoadWeighingYes,
                rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther, rb10OutOfServiceNo, rb10OutOFServiceYes, rb10PositionIndicatorTypeFlushMount,
                rb10PositionIndicatorTypeSurfaceMount, rb10ProtectiveBlanketNo, rb10ProtectiveBlanketYes, rb10RearDoorKeySwitchNo,
                rb10RearDoorKeySwitchYes, rb10RearWallOther, rb10RearWallStainlessSteel, rb10SecurityKeySwitchNo, rb10SecurityKeySwitchYes,
                rb10SideWallOther, rb10SideWallStainlesSteel, rb10StructureShaftConcrete, rb10StructureShaftOther, rb10TrimmerBeamsNo,
                rb10TimmerbeamsYes, rb10VoiceAnnunciationNo, rb10VoiceAnnunciationYes },

           new Object[] { tb11AuxCOPLocation, tb11CarDepth, tb11CarDoorFinishText, tb11CarHeight, tb11LiftCarLoad, tb11CarWidth, tb11CeilingFinishText,
                tb11ControlerLocationText, tb11COPFinishText, tb11Designations, tb11DoorHeight, tb11DoorTracksText, tb11DoorWidth, tb11FaceplateMaterialText,
                tb11FloorFinish, tb11FrontWallText, tb11HandrailText, tb11Headroom, tb11KeyswitchLocation, tb11LandingDoorFInishOther, rb11LiftNumbers,
                tb11LiftRating, tb11MainCOPLocation, tb11MirrorText, tb11LiftCarNote, tb11NumberofCarEntrances, tb11NumberOfCOPs, tb11NumberOfLandingDoors,
                tb11NumberOfLandings, tb11NumberOfLEDLights, tb11PitDepth, tb11RearWallText, tb11ShaftDepth, tb11ShaftWidth, tb11SideWallText, tb11Speed,
                rb11StructureShaftText, tb11Travel, tb11TypeOfLift, rb11AdvancedOpeningNo, rb11AdvancedOpeningYes, rb11BumpRailNo, rb11BumpRailYes,
                rb11CarDoorFinishOther, rb11CarDoorFinishStainlessSteel, rb11CeilingFinishMirrorStainlessSteel, rb11CeilingFinishOther,
                rb11CeilingFinishStainlessSteel, rb11CeilingFinishWhite, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther,
                rb11ControlerLocationShaft, tb11ControlerLocationTopLanding, rb11COPFinishOther, rb11COPFinishStainlessSteel, rb11DoorNudgingNo,
                rb11DoorNudgingYes, rb11DoorTracksAluminium, rb11DoorTracksOther, rb11DoorTypeCentreOpening, rb11DoorTypeSideOpening,
                rb11EmergencyLoweringSystemNo, rb11EmergencyLoweringSystemYes, rb11ExclusiveServiceNo, rb11ExclusiveServiceYes,
                rb11FacePlateMaterialOther, rb11FacePlateMaterialStainlessSTeel, rb11FalseCeilingNo, rb11FalseCEilingYes, rb11FalseFloorNo, rb11FalseFloorYes,
                rb11FireServiceNo, rb11FireServiceYes, rb11FrontWallOther, rb11FrontWallStainlessSteel, rb11GPOInCarNo, rb11GPOInCarYes, rb11HandrailOther,
                rb11HandrailStainlessSteel, rb11IndependentSErviceNO, rb11IndependentServiceYes, rb11LandingDoorFinishOther,
                rb11LandingDoorFinishStainlessSteel, rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite, rb11LoadWeighingNo, rb11LoadWeighingYes,
                rb11MirrorFullSize, rb11MirrorHalfSize, rb11MirrorOther, rb11OutOfServiceNo, rb11OutOfSErviceYes, rb11PositionIndicatorTypeFlushMount,
                rb11PositionIndicatorTypeSurfaceMount, rb11ProtectiveBlanketNo, rb11ProtectiveBlanketsYes, rb11RearDoorKeySwitchNo,
                rb11RearDoorKeySwitchYes, rb11RearWallOther, rb11RearWallStainlessSteel, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes,
                rb11SideWallOther, rb11SideWallStainlessSteel, rb11StructureShaftConcrete, rb11StrructureShaftOther, rb11TrimmerBeamNo,
                rb11TrimmerBeamsYes, rb11VoiceAnnunciationNo, rb11VoiceAnnunciationYes },

           new Object[] { tb12AuxCOPLocation, tb12CarDepth, tb12CarDoorFinishText, tb12CarHeight, tb12CarLoad, tb12CarWidth, tb12CeilingFinishText,
                tb12ControlerLocationText, tb12COPFinishText, tb12Designations, tb12LandingDoorHeight, tb12DoorTracksText, tb12LandingDoorWidth, tb12FacePlateMaterialText,
                tb12FloorFinish, tb12FrontWallText, tb12HandrailText, tb12Headroom, tb12KeyswitchLocation, tb12LandingDoorFinishText, tb12LiftNumbers,
                tb12CarLiftRating, tb12MainCOPLocation, tb12MirrorText, tb12LiftCarNotes, tb12CarNumberOfCarEntrances, tb12NumberOfCOPs, tb12NumberOfLandingDoors,
                tb12NumberOfLandings, tb12NumberOfLEDLights, tb12PitDepth, tb12RearWallText, tb12ShaftDepth, tb12ShaftWidth, tb12SideWallText, tb12CarSpeed,
                tb12StructureShaftText, tb12Travel, tb12TypeOfLift, rb12AdvancedOpeningNo, rb12AdvancedOpeningYes, rb12BumpRailNo, rb12BumpRailYes,
                rb12CarDoorFinishOther, rb12CarDoorFinishStainlessSteel, rb12CeilingFinishMirrorStainlessSteel, rb12CeilingFinishOther,
                rb12CeilingFinishStainlessSteel, rb12CeilingFinishWhite, rb12ControlerLocationBottomLanding, rb12ControlerLocationOther,
                rb12ControlerLocationShaft, rb12ControlerLocationTopLanding, rb12COPFinishOther, rb12COPFinishStainlessSteel, rb12LandingDoorNudgingNo,
                rb12LandingDoorNudgingYes, rb12DoorTracksAluminium, rb12DoorTracksOther, rb12DoorTypeCentreOpening, rb12DoorTypeSideOpening,
                rb12EmergencyLoweringSystemNo, rb12EmergencyLoweringSystemYes, rb12ExclusiveServiceNo, rb12ExclusiveServiceYes,
                rb12FacePlateMaterialOther, rb12FacePlateMaterialStainlessSteel, rb12FalseCeilingNo, rb12FalseCeilingYes, rb12FalseFloorNo, rb12FalseFloorYes,
                rb12FireServiceNo, rb12FireServiceYes, rb12FrontWallOther, rb12FrontWallStainlessSteel, rb12GPOInCarNo, rb12GPOInCarYes, rb12HandrailOther,
                rb12HandrailStainlessSTeel, rb12IndependentServiceNo, rb12IndependentServiceYes, rb12LandingDoorFinishOther,
                rb12LandingDoorFinishStainlessSteel, rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite, rb12LoadWeighingNo, rb12LoadWeighingYes,
                rb12MirrorFullSize, rb12MirrorHalfSize, rb12MirrorOTher, rb12OutOfServiceNo, rb12OutOfServiceYes, rb12PositionIndicatorTypeFlushMount,
                rb12PositionIndicatorTypeSurfaceMount, rb12ProtectiveBlanketsNo, rb12ProectiveBlanketsYes, rb12RearDoorKeySwitchNo,
                rb12RearDoorKeySwitchYes, rb12RearWallOther, rb12RearWallStainlessSteel, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes,
                rb12SideWallOther, rb12SideWallStainlessSteel, rb12StructureShaftConcrete, rb12StructureShaftOther, rb12TrimmerBeamsNo,
                rb12TrimmerBeamsYes, rb12VoicAnnunciationNo, rb12VoiceAnnuniationYes }
        };

            #endregion

            for (int i = 0; i < pageObjects[pageToBeFilled].Length; i++)
            {
                if (pageObjects[pageToBeFilled][i] is TextBox)
                {
                    TextBox sourceTextBox = (TextBox)pageObjects[0][i];
                    TextBox recipientTextBox = (TextBox)pageObjects[pageToBeFilled][i];

                    recipientTextBox.Text = sourceTextBox.Text;
                }
                else if (pageObjects[pageToBeFilled][i] is RadioButton)
                {
                    RadioButton sourceRadioButton = (RadioButton)pageObjects[0][i];
                    RadioButton recipientRadioButton = (RadioButton)pageObjects[pageToBeFilled][i];

                    recipientRadioButton.Checked = sourceRadioButton.Checked;
                }
                else if (pageObjects[pageToBeFilled][i] is CheckBox)
                {
                    CheckBox sourceCheckBox = (CheckBox)pageObjects[0][i];
                    CheckBox recipientCheckBox = (CheckBox)pageObjects[pageToBeFilled][i];

                    recipientCheckBox.Checked = sourceCheckBox.Checked;
                }
            }
        }
        #endregion

        #region Pannel Switching Methods 
        private void PanelMenuChange(Panel panelToBeShown)
        {
            SetValuesForMainMenuLabels();
            Panel[] panels = { panelAdditionalCosts, panelLiftPrices, panelShipping };

            foreach (Panel p in panels)
            {
                if (p != panelToBeShown)
                {
                    PanelDefaultSettings(p, p.Location);
                }
                if (p == panelToBeShown)
                {
                    PanelDefaultSettings(p, p.Location, !p.Visible, !p.Enabled);
                }
            }
        }

        private void liftCostsClose_Click(object sender, EventArgs e)
        {
            PanelMenuChange(null);
        }

        private void extraCostsClose_Click(object sender, EventArgs e)
        {
            PanelMenuChange(null);
        }

        private void shippingCostsClose_Click(object sender, EventArgs e)
        {
            PanelMenuChange(null);
        }
        private void btnLiftCosts_Click(object sender, EventArgs e)
        {
            PanelMenuChange(panelLiftPrices);
        }

        private void btnShippingCosts_Click(object sender, EventArgs e)
        {
            PanelMenuChange(panelShipping);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            PanelMenuChange(panelAdditionalCosts);
        }

        #endregion

        #region Loading Old Quote Data Methods
        private void PullInfo()
        {
            try
            {
                #region TB Load
                LoadPreviousXmlTb(tbfname, tblname, tbphone, tbAddress1, tbAddress2, tbAddress3, tbLiftNumbers, tbTypeofLift,
                    tbShaftDepth, tbShaftWidth, tbPitDepth, tbHeadroom, tbTravel, tbNumofLandingDoors, tbNumofLandings,
                    tbNumberOfCOPS, tbMainCOPLocation, tbAuxCOPLocation, tbKeyswitchLocation, tbDesignations, tbNumOfLEDLights, tbFloorFinish,
                    tbDoorWidth, tbDoorHeight, tbLoad, tbSpeed, tbwidth, tbDepth, tbHeight, tbLiftRating, tbNumofCarEntrances, tbLiftCarNotes,
                    tb2AuxCOPLocation, tb2CarDepth, tb2CarDoorFinishText,
                    tb2CarHeight, tb2CarLoad, tb2CarWidth, tb2CeilingFinishText, tb2ControlerLocationText, tb2COPFinishText, tb2Designations,
                    tb2DoorHeight, tb2DoorTracksText, tb2DoorWidth, tb2FacePlateMaterialText, tb2FloorFinish, tb2FrontWallText,
                    tb2HandrailText, tb2Headroom, tb2KeyswitchLocation, tb2LandingDoorFinishText, tb2LiftNumbers, tb2LiftRating,
                    tb2MainCOPLocation, tb2MirrorText, tb2Note, tb2NumberOfCarEntrances, tb2NumberOfCOPs, tb2NumberOfLAndingDoors,
                    tb2NumberOfLandings, tb2NumberofLEDLights, tb2PitDepth, tb2RearWallText, tb2ShaftDepth, tb2ShaftWidth,
                    tb2SideWallText, tb2Speed, tb2StructureShaftText, tb2Travel, tb2TypeOfLift, tb3AuxCOPLocation, tb3CarDepth, tb3CarDoorFinishText,
                    tb3CarHeight, tb3CarNote, tb3CarWidth, tb3CEilingFinishText, tb3ControlerLocationText, tb3COPFinishText, tb3Designations,
                    tb3DoorHeight, tb3DoorTracksText, tb3DoorWidth, tb3FacePlaterMaterialText, tb3FloorFinish, tb3FrontWallText,
                    tb3HandrailText, tb3HeadRoom, tb3KeyswitchLocation, tb3LandingDoorFinishText, tb3LiftNumbers, tb3LiftRating,
                    tb3Load, tb3MainCOPLocation, tb3MirrorText, tb3NumberOfCarEntrances, tb3NumberOfCOPs, tb3NumberOfLandingDoors,
                    tb3NumberOfLandings, tb3NumberOfLEDLights, tb3PitDepth, tb3RearWallText, tb3ShaftDepth, tb3ShaftWidth,
                    tb3SideWallText, tb3Speed, tb3StructureShaftText, tb3Travel, tb3TypeOfLift, tb4AuxCOPLocation, tb4CarDepth, tb4CarDoorFinish,
                    tb4CarHeight, tb4CarNote, tb4CarWidth, tb4CeilingFinishText, tb4ControlerLocationText, tb4COPFinishText, tb4Designations,
                    tb4DoorHeight, tb4DoorTracksText, tb4DoorWidth, tb4FacePlateMaterialText, tb4FloorFinish, tb4FrontWallText,
                    tb4HandrailText, tb4Headroom, tb4KeyswitchLocations, tb4LandingDoorFinishText, tb4LiftNumbers, tb4LiftRating,
                    tb4Load, tb4MainCOPLocation, tb4MirrorText, tb4NumberOfCarEntrances, tb4NumberOfCOPs, tb4NumberOfLandingDoors,
                    tb4NumberOfLandings, tb4NumbeROfLEDLights, tb4PitDepth, tb4RearWallText, tb4ShaftDepth, tb4ShaftWidth, tb4SideWallText,
                    tb4Speed, tb4StructureShaftText, tb4Travel, tb4TypeOfLift, tb5AuxCOPLocation, tb5CarDepth, tb5CarDoorFinishText,
                    tb5CaRHeight, tb5CarNote, tb5CarWidth, tb5CeilingFinishText, tb5ControlerLocationText, tb5COPFinishText, tb5Designations,
                    tb5DoorHeight, tb5DoorTRacksText, tb5DoorWidth, tb5FacePlateMaterialText, tb5FloorFinish, tb5FrontWallText,
                    tb5HandrailText, tb5Headroom, tb5KetyswitchLocation, tb5LandingDoorFinishText, tb5LiftNumbers, tb5LiftRating,
                    tb5Load, tb5MainCOPLocation, tb5MirrorText, tb5NumberOfCarEntrances, tb5NumberOfCOPs, tb5NumberOfLandingDoors,
                    tb5NumberOfLandings, tb5NumberOIfLEDLights, tb5PitDepth, tb5RearWallText, tb5ShaftDEpth, tb5ShaftWidth,
                    tb5SideWallText, tb5Speed, tb5StructureShaftText, tb5Travel, tb5TypeOfLift, tb6AuxCOPLocation, tb6CarDepth, tb6CarDoorFinishText,
                    tb6CarHeight, tb6CarLoad, tb6CarNote, tb6CarSpeed, tb6CarWidth, tb6CeilingFinishText, tb6ControlerLocationText, tb6COPFinishText,
                    tb6Designations, tb6DoorHeight, tb6DoorTracksOther, tb6DoorWidth, tb6FacePlateMaterialText, tb6FloorFinish,
                    tb6FrontWallText, tb6HAndrailText, tb6Headroom, tb6KeySwitchLocation, tb6LandingDoorFinishText, tb6LiftNumbers,
                    tb6LiftRating, tb6MainCOPLocation, tb6MirrorText, tb6NumberOfCarEntrances, tb6NumberOFCOPs, tb6NumberOfLandingDoors,
                    tb6NumberOfLandings, tb6NumberOfLEDLights, tb6PitDepth, tb6RearWallText, tb6ShaftDepth, tb6ShaftWidth,
                    tb6SideWallText, tb6StructureShaftText, tb6Travel, tb6TypeOfLift, tb7AuzCOPLocation, tb7CarDepth, tb7CarDoorFinishText,
                    tb7CarHeight, tb7CarLoad, tb7CarNotes, tb7CarSpeed, tb7CarWidth, tb7CEilingFinishText, tb7ControlerLocationText, tb7COPFinishText,
                    tb7Designations, tb7DoorHeight, tb7DoorTracksText, tb7DoorWidth, tb7FacePlateMaterialText, tb7FloorFinish, tb7FrontWallText,
                    tb7HandrailText, tb7HeadRoom, tb7KeyswitchLocation, tb7LandingDoorFinishText, tb7LiftNumbers, tb7LiftRating,
                    tb7MainCOPLocation, tb7MirrorText, tb7NumberOfCarEntrances, tb7NumberOfCOPs, tb7NumberOfLandingDoors, tb7NumberOfLandings,
                    tb7NumberOfLEDLights, tb7PitDepth, tb7RearWallText, tb7ShaftDepth, tb7ShaftWidth, tb7SideWallText,
                    tb7StructureShaftText, tb7Travel, tb7TypeOfLift, tb8AuxCOPLocation, tb8CarDEpth, tb8CarDoorFinishText,
                    tb8CarHeight, tb8CarWidth, tb8CeilingFinishText, tb8ControlerLocationText, tb8COPFinishText, tb8Desiginations, tb8DoorHeight,
                    tb8DoorTracksText, tb8DoorWidth, tb8FacePlateMaterialText, tb8FloorFinish, tb8FrontWallText, tb8HandrailText,
                    tb8Headroom, tb8KeyswitchLocations, tb8LandingDoorFinishText, tb8LiftCarNotes,
                    tb8LiftNumbers, tb8LiftRating, tb8Load, tb8MainCOPLocation, tb8MirrorText, tb8NumberOfCarEntrances, tb8NumberOfCOPs,
                    tb8NumberOfLandingDoors, tb8NumberOfLandings, tb8NumberofLEDLights, tb8PitDepth, tb8RearWallText,
                    tb8ShaftDepth, tb8ShaftWidth, tb8SideWallText, tb8Speed, tb8StructureShaftText, tb8Travel, tb8TypeOfLift, tb9AuxCOPLocation, tb9CarDepth, tb9CarDoorFinishText, tb9CarHeight,
                    tb9CarNotes, tb9CarWidth, tb9CeilingFinishText, tb9ControlerLocationText, tb9COPFinishText, tb9Designations, tb9DoorHeight, tb9DoorTracksText,
                    tb9DoorWidth, tb9FacePlateMaterialText, tb9FloorFinish, tb9FrontWallText, tb9HandrailTexrt, tb9Headroom, tb9KeyswitchLocation,
                    tb9LandingDoorFinishText, tb9LiftNumbers, tb9LiftRating, tb9Load, tb9MainCOPLocation, tb9MirrorText, tb9NumberOFCarEntraces,
                    tb9NumberOfCOPs, tb9NumberOfLandingDoors, tb9NumberOfLandings, tb9NumberOfLEDLights, tb9PitDepth, tb9RearWallText,
                    tb9ShaftDepth, tb9ShaftWidth, tb9SideWallText, tb9Speed, tb9StructureShaftText, tb9Travel, tb9TypeOfLift, tb10AuxCOPLocation, tb10CarDepth, tb10CarDoorFinishText,
                    tb10CarHeight, tb10CarWidth, tb10CEilingFinishText, tb10ControlerLocationText, tb10COPFinishText, tb10Desigination, tb10DoorHeight,
                    tb10DoorTracksText, tb10DoorWidth, tb10FacePlateMaterialText, tb10FloorFinish, tb10FrontWallText, tb10HandrailText,
                    tb10Headroom, tb10KeyswitchLocation, tb10LandingDoorFinishText, tb10LiftCarLoad, tb10LiftCarNotes, tb10LiftNumbers,
                    tb10LiftRating, tb10MainCOPLocation, tb10MirrorText, tb10NumberofCarEntrances, tb10NumberOfCOPs, tb10NumberofLandingDoors,
                    tb10NumberofLandings, tb10NumberOfLEDLIghts, tb10PitDepth, tb10RearWallText, tb10ShaftDepth, tb10ShaftWidth,
                    tb10SideWallText, tb10Speed, tb10StructureShaftText, tb10Travel, tb10TypeOfLift, tb11AuxCOPLocation, tb11CarDepth, tb11CarDoorFinishText,
                    tb11CarHeight, tb11CarWidth, tb11CeilingFinishText, tb11ControlerLocationText, tb11COPFinishText, tb11Designations, tb11DoorHeight,
                    tb11DoorTracksText, tb11DoorWidth, tb11FaceplateMaterialText, tb11FloorFinish, tb11FrontWallText, tb11HandrailText, tb11Headroom,
                    tb11KeyswitchLocation, tb11LandingDoorFInishOther, tb11LiftCarLoad, tb11LiftCarNote, tb11LiftRating, tb11MainCOPLocation,
                    tb11MirrorText, tb11NumberofCarEntrances, tb11NumberOfCOPs, tb11NumberOfLandingDoors, tb11NumberOfLandings, tb11NumberOfLEDLights,
                     tb11PitDepth, tb11RearWallText, tb11ShaftDepth, tb11ShaftWidth, tb11SideWallText, tb11Speed, tb11Travel, tb11TypeOfLift,
                    rb11LiftNumbers, rb11StructureShaftText, tb12AuxCOPLocation, tb12CarDepth, tb12CarDoorFinishText,
                    tb12CarHeight, tb12CarLiftRating, tb12CarLoad, tb12CarNumberOfCarEntrances, tb12CarSpeed, tb12CarWidth, tb12CeilingFinishText,
                    tb12ControlerLocationText, tb12COPFinishText, tb12Designations, tb12DoorTracksText, tb12FacePlateMaterialText,
                    tb12FloorFinish, tb12FrontWallText, tb12HandrailText, tb12Headroom, tb12KeyswitchLocation, tb12LandingDoorFinishText,
                    tb12LandingDoorHeight, tb12LandingDoorWidth, tb12LiftCarNotes, tb12LiftNumbers, tb12MainCOPLocation, tb12MirrorText,
                    tb12NumberOfCOPs, tb12NumberOfLandingDoors, tb12NumberOfLandings, tb12NumberOfLEDLights, tb12PitDepth,
                     tb12RearWallText, tb12ShaftDepth, tb12ShaftWidth, tb12SideWallText, tb12StructureShaftText, tb12Travel, tb12TypeOfLift,
                     tbControlerLocation, tbStructureShaft, tbLandingDoorFinish, tbDoorTracks, tbCarDoorFinish, tbCeilingFinish, tbFrontWall,
                     tbMirror, tbHandrail, tbSideWall, tbRearWall, tbCOPFinish, tbFacePlateMaterial
                    );
                #endregion

                #region RB Load
                LoadPreviousXmlRb(rb10AdvancedOpeningNo, rb10AdvancedOpeningYes, rb10BumpRaidYes, rb10BumpRailNo, rb10CarDoorFinishOther,
                    rb10CarDoorFinishStainlessSteel, rb10CEilingFinishMirrorStainlessSteel, rb10CEilingFinishOther, rb10CeilingFinishStainlessSteel,
                    rb10CeilingFinishWhite, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther, rb10ControlerLocationShaft,
                    rb10ControlerLocationTopLanding, rb10COPFinishOther, rb10COPFinishStainlessSteel, rb10DoorNudgingNo, rb10DoorNudgingYes,
                    rb10DoorTracksAluminium, rb10DoorTracksOther, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening, rb10EmergencyLoweringSystemYes,
                    rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes, rb10FacePlateMaterialOther, rb10FacePlateMaterialStainlessSteel, rb10FalseCeilingNo,
                    rb10FalseCeilingYes, rb10FalseFloorNo, rb10FalseFloorYes, rb10FireSERviceNo, rb10FireSErviceYes, rb10FrontWallOther,
                    rb10FrontWallStainlessSteel, rb10GPOInCarNo, rb10GPOInCarYes, rb10HandrailOther, rb10HandrailStainlessSteel, rb10IndependentServiceNo,
                    rb10IndependentServiceYes, rb10LAndingDoorFinishOtherr, rb10LandingDoorFinishStainlessSteel, rb10LCDColourBlue, rb10LCDColourRed,
                    rb10LCDColourWhite, rb10LoadWEighingNo, rb10LoadWeighingYes, rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther,
                    rb10OutOfServiceNo, rb10OutOFServiceYes, rb10PositionIndicatorTypeFlushMount, rb10PositionIndicatorTypeSurfaceMount,
                    rb10ProtectiveBlanketNo, rb10ProtectiveBlanketYes, rb10RearDoorKeySwitchNo, rb10RearDoorKeySwitchYes, rb10RearWallOther,
                    rb10RearWallStainlessSteel, rb10SecurityKeySwitchNo, rb10SecurityKeySwitchYes, rb10SideWallOther, rb10SideWallStainlesSteel,
                    rb10StructureShaftConcrete, rb10StructureShaftOther, rb10TimmerbeamsYes, rb10TrimmerBeamsNo, rb10VoiceAnnunciationNo,
                    rb10VoiceAnnunciationYes, rb11AdvancedOpeningNo, rb11AdvancedOpeningYes, rb11BumpRailNo, rb11BumpRailYes, rb11CarDoorFinishOther,
                    rb11CarDoorFinishStainlessSteel, rb11CeilingFinishMirrorStainlessSteel, rb11CeilingFinishOther, rb11CeilingFinishStainlessSteel,
                    rb11CeilingFinishWhite, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther, rb11ControlerLocationShaft, rb11COPFinishOther,
                    rb11COPFinishStainlessSteel, rb11DoorNudgingNo, rb11DoorNudgingYes, rb11DoorTracksAluminium, rb11DoorTracksOther,
                    rb11DoorTypeCentreOpening, rb11DoorTypeSideOpening, rb11EmergencyLoweringSystemNo, rb11EmergencyLoweringSystemYes,
                    rb11ExclusiveServiceNo, rb11ExclusiveServiceYes, rb11FacePlateMaterialOther, rb11FacePlateMaterialStainlessSTeel, rb11FalseCeilingNo,
                    rb11FalseCEilingYes, rb11FalseFloorNo, rb11FalseFloorYes, rb11FireServiceNo, rb11FireServiceYes, rb11FrontWallOther,
                    rb11FrontWallStainlessSteel, rb11GPOInCarNo, rb11GPOInCarYes, rb11HandrailOther, rb11HandrailStainlessSteel,
                    rb11IndependentSErviceNO, rb11IndependentServiceYes, rb11LandingDoorFinishOther, rb11LandingDoorFinishStainlessSteel,
                    rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite, rb11LoadWeighingNo, rb11LoadWeighingYes,
                    rb11MirrorFullSize, rb11MirrorHalfSize, rb11MirrorOther, rb11OutOfServiceNo, rb11OutOfSErviceYes, rb11PositionIndicatorTypeFlushMount,
                    rb11PositionIndicatorTypeSurfaceMount, rb11ProtectiveBlanketNo, rb11ProtectiveBlanketsYes, rb11RearDoorKeySwitchNo,
                    rb11RearDoorKeySwitchYes, rb11RearWallOther, rb11RearWallStainlessSteel, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes,
                    rb11SideWallOther, rb11SideWallStainlessSteel, rb11StrructureShaftOther, rb11StructureShaftConcrete,
                    rb11TrimmerBeamNo, rb11TrimmerBeamsYes, rb11VoiceAnnunciationNo, rb11VoiceAnnunciationYes, rb12AdvancedOpeningNo,
                    rb12AdvancedOpeningYes, rb12BumpRailNo, rb12BumpRailYes, rb12CarDoorFinishOther, rb12CarDoorFinishStainlessSteel,
                    rb12CeilingFinishMirrorStainlessSteel, rb12CeilingFinishOther, rb12CeilingFinishStainlessSteel, rb12CeilingFinishWhite,
                    rb12ControlerLocationBottomLanding, rb12ControlerLocationOther, rb12ControlerLocationShaft, rb12ControlerLocationTopLanding,
                    rb12COPFinishOther, rb12COPFinishStainlessSteel, rb12DoorTracksAluminium, rb12DoorTracksOther, rb12DoorTypeCentreOpening,
                    rb12DoorTypeSideOpening, rb12EmergencyLoweringSystemNo, rb12EmergencyLoweringSystemYes, rb12ExclusiveServiceNo,
                    rb12ExclusiveServiceYes, rb12FacePlateMaterialOther, rb12FacePlateMaterialStainlessSteel, rb12FalseCeilingNo, rb12FalseCeilingYes,
                    rb12FalseFloorNo, rb12FalseFloorYes, rb12FireServiceNo, rb12FireServiceYes, rb12FrontWallOther, rb12FrontWallStainlessSteel,
                    rb12GPOInCarNo, rb12GPOInCarYes, rb12HandrailOther, rb12HandrailStainlessSTeel, rb12IndependentServiceNo, rb12IndependentServiceYes,
                    rb12LandingDoorFinishOther, rb12LandingDoorFinishStainlessSteel, rb12LandingDoorNudgingNo, rb12LandingDoorNudgingYes,
                    rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite, rb12LoadWeighingNo, rb12LoadWeighingYes, rb12MirrorFullSize,
                    rb12MirrorHalfSize, rb12MirrorOTher, rb12OutOfServiceNo, rb12OutOfServiceYes, rb12PositionIndicatorTypeFlushMount,
                    rb12PositionIndicatorTypeSurfaceMount, rb12ProectiveBlanketsYes, rb12ProtectiveBlanketsNo, rb12RearDoorKeySwitchNo,
                    rb12RearDoorKeySwitchYes, rb12RearWallOther, rb12RearWallStainlessSteel, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes,
                    rb12SideWallOther, rb12SideWallStainlessSteel, rb12StructureShaftConcrete, rb12StructureShaftOther, rb12TrimmerBeamsNo,
                    rb12TrimmerBeamsYes, rb12VoicAnnunciationNo, rb12VoiceAnnuniationYes, rb2AdvancedOpeningNo, rb2AdvancedOpeningYes,
                    rb2BumpRailNo, rb2BumpRailYes, rb2CarDoorFinishOther, rb2CarDoorFinishStainlessSteel, rb2CeilingFinishMirrorStainlessSteel,
                    rb2CeilingFinishOther, rb2CeilingFinishStainlessSteel, rb2CeilingFinishWhite, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther,
                    rb2ControlerLocationShaft, rb2ControlerLocationTopLanding, rb2COPFinishOther, rb2COPFinishStainlessSTeell, rb2DoorNudgingNo,
                    rb2DoorNudgingYes, rb2DoorTracksAluminium, rb2DoorTracksOther, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening,
                    rb2EmergemncyLoweringSystemNo, rb2EmergencyLoweringSystemYes, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes,
                    rb2FacePlateMaterialOther, rb2FacePlateMaterialStainlessSteel, rb2FalseCeilingNo, rb2FalseCeilingYes, rb2FalseFloorNo,
                    rb2FalseFloorYes, rb2FireSErviceNo, rb2FireSErviceYes, rb2FrontWallOther, rb2FrontWallStainlessSteel, rb2GPOInCarNo,
                    rb2GPOInCarYes, rb2HandrailOther, rb2HandRailStainlessSteel, rb2IndependentServiceNo, rb2IndependentServiceYes,
                    rb2LandingDoorFinishOther, rb2LandingDoorFinishStainlessSteel, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite,
                    rb2LoadWeighingNo, rb2LoadWeighingYes, rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther, rb2OutOfServiceNo,
                    rb2OutOfServiceYes, rb2PositionIndicatorTypeFlushMount, rb2PositionIndicatorTypeSurfaceMount, rb2ProtectiveBlanketsNo,
                    rb2ProtectiveBlanketsYes, rb2RearDoorKeySwitchNo, rb2RearDoorKeySwitchYes, rb2RearWallOther, rb2RearWallStainlessSteel,
                    rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes, rb2SideWallOther, rb2SideWallStainlessSteel, rb2StructureShaftConcrete,
                    rb2StructureShaftOther, rb2TrimmerBeamsNo, rb2TrimmerBeamsYes, rb2VoiceAnnunciationNo, rb2VoiceAnnunciationYes,
                    rb3AdvancedOpeningNo, rb3AdvancedOpeningYes, rb3BumpRailNo, rb3BumpRailYes, rb3CarDoorFinishOther, rb3CarDoorFinishStainlessSteel,
                    rb3CeilingFinishOther, rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite, rb3ControleRLocationBottomLanding, rb3ControlerLocationOther,
                    rb3ControlerLocationShaft, rb3ControlerLocationTopLanding, rb3COPFinishOther, rb3COPFinishStainlessSteel, rb3DoorNudgingNo,
                    rb3DoorNudgingYes, rb3DoorTracksAluminium, rb3DoorTracksOther, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening,
                    rb3EmergencyLoweringSystemNo, rb3EmergencyLoweringSystemYes, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes,
                    rb3FacePlateMaterialOther, rb3FacePlateMaterialStainlessSteel, rb3FalseCeilingNo, rb3FalseCeilingYes, rb3FalseFloorNo, rb3FalseFloorYes,
                    rb3FireServiceYes, rb3FireServieNo, rb3FrontWallOther, rb3FrontWallStainlessSteel, rb3GPOInCarNo, rb3GPOInCarYes,
                    rb3HandrailOther, rb3HandrailStainlessSteel, rb3IndependentServiceNo, rb3IndependentServiceYes, rb3landingDoorFinishOther,
                    rb3LandingDoorFinishStainlessSteel, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite, rb3LoadWeighingNo, rb3LoadWeighingYes,
                    rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther, rb3MirrorStainlessSteel, rb3OutOfServiceNo, rb3OutOfSErviceYes,
                    rb3PositionIndicatorTypeFlushMount, rb3PositionIndicatorTypeSurfaceMount, rb3ProtectiveBlanketsNo, rb3ProtectiveBlanketsYes,
                    rb3RearDoorKeySwitchNo, rb3RearDoorKeySwitchYes, rb3RearWallOther, rb3RearWallStainlessSteel, rb3SecurityKeySwitchNo,
                    rb3SecurityKeySwitchYes, rb3SideWallOther, rb3SideWallStainlessSteel, rb3StructureShaftConcrete, rb3StructureShaftOther,
                    rb3TrimmerBeamsNo, rb3TrimmerBeamsYes, rb3VoiceAnnunciationNo, rb3VoiceAnnunciationYes, rb4AdvancedOpeningNo,
                    rb4AdvancedOpeningYes, rb4BumpRailNo, rb4BumpRailYes, rb4CarDoorFinishOther, rb4CarDoorFinishStainlessSteel,
                    rb4CeilingFinishMirrorStainlessSteel, rb4CeilingFinishOther, rb4CeilingFinishStainlessSteel, rb4CeilingFinishWhite,
                    rb4ControelrLocationTopLanding, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther, rb4ControlerLocationShaft,
                    rb4COPFinishOther, rb4COPFinishStainlessSteel, rb4DoorNudgingNo, rb4DoorNudgingYes, rb4DoorTracksAluminium, rb4DoorTracksOther,
                    rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening, rb4EmergencyLoweringSystemNo, rb4EmergencyLoweringSystemYes,
                    rb4ExclusiveServiceNo, rb4ExclusiveServiceYes, rb4FacePlateMaterialOther, rb4FacePlateMaterialStainlessSteel, rb4FalseCeilingNO,
                    rb4FalseCeilingYes, rb4FalseFloorNo, rb4FalseFloorYes, rb4FireSErviceNo, rb4FireServiceYes, rb4FrontWallStainlessSteel, rb4FrotnWallOther,
                    rb4GPOInCarNo, rb4GPOInCarYes, rb4HandrailOther, rb4HandRailStainlesSteel, rb4IndependentServiceYes, rb4LandingDoorFinishOther,
                    rb4LandingDoorFinishStainlessSteel, rb4LCDColourBlue, rb4LCDColourRed, rb4LCDColourWhite, rb4LoadWeighingNo, rb4LoadWeighingYes,
                    rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther, rb4OutOfServiceNo, rb4OutOfServiceYes, rb4PositionIndicatorTypeFlushMount,
                    rb4PositionIndicatorTypeSurfaceMount, rb4ProtectiveBlanketsYes, rb4ProtetiveBlanketsNo, rb4RearDoorKeySwitchNo,
                    rb4RearDoorKeySwitchYes, rb4RearWallOther, rb4RearWallStainlessSteel, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes,
                    rb4SideWallOther, rb4SideWallStainlessSteel, rb4StructureShaftConcrete, rb4StructureShaftOther, rb4TrimmerBeamsNo,
                    rb4TrimmerBeamsYes, rb4VoiceAnnunciationNo, rb4VoiceAnnunciationYes, rb5AdvancedOpeningNo, rb5AdvancedOpeningYes, rb5BumpRailNo,
                    rb5BumpRailYes, rb5CarDoorFinishOther, rb5CarDoorFinishStainlessSteel, rb5CeilingFinishMirrorStainlessSTeel, rb5CeilingFinishOther,
                    rb5CeilingFinishStainlessSteel, rb5CeilingFinishWhite, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther,
                    rb5ControlerLocationShaft, rb5ControlerLocationTopLanding, rb5COPFinishOther, rb5COPFinishStainlessSteel, rb5DoorNudgingNo,
                    rb5DoorNudgingYes, rb5DoorTracksAluminium, rb5DoorTracksOther, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening,
                    rb5EmergencyLoweringSystemNo, rb5EmergencyLoweringSystemYes, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes,
                    rb5FacePlateMAterialOtjer, rb5FacePlateMaterialStainlessSteel, rb5FalseCeilingNo, rb5FalseCeilingYes, rb5FalseFloorNo, rb5FalseFloorYes,
                    rb5FireSErviceNo, rb5FireServiceYes, rb5FrontWallOther, rb5FrontWallStainlessSteel, rb5GPOInCarNo, rb5GPOInCarYes,
                    rb5HandRailOther, rb5HandRailStainlessSteel, rb5IndependentServiceNo, rb5IndependentServiceYes, rb5LAndingDoorFinishOther,
                    rb5LandingDoorFinishStainlessSteel, rb5LCDColoiurWHite, rb5LCDColourBlue, rb5LCDColourRed, rb5LoadWeighingNo, rb5LoadWeighingYes,
                    rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther, rb5OutOfServiceNo, rb5OutOfServiceYes, rb5PositionIndicatorTypeSurfaceMount,
                    rb5PostionIndicatorTypeFlushMount, rb5ProtectiveBlanketsNo, rb5ProtectiveBlanketsYes, rb5RearDoorKeySwitchNo, rb5RearDoorKeySwitchYes,
                    rb5RearWallOther, rb5RearWallStainlesSteel, rb5SecurityKeySwitchNo, rb5SecurityServiceYes, rb5SideWallOther, rb5SideWallStainlessSTeel,
                    rb5StructureShaftConcrete, rb5StructureShaftOther, rb5TrimmerBeamsNo, rb5TrimmerBeamsYes, rb5VoiceAnnunciationNo,
                    rb5VoiceAnnunciationYes, rb6AdvancedOpeningNo, rb6AdvancedOpeningYes, rb6BumpRailNo, rb6BumpRailYes, rb6CarDoorFinishOther,
                    rb6CarDoorFinishStainlessSteel, rb6CeilingFinishMirrorStainlessSteel, rb6CeilingFinishOther, rb6CeilingFinishStainlessSteel,
                    rb6CeilingFinishWhite, rb6ControlerLocationBottomLanding, rb6ControlerLocationOther, rb6ControlerLocationShaft,
                    rb6ControlerLocationTopLanding, rb6COPFinishOther, rb6COPFinishStainlessSteel, rb6DoorNudgingNo, rb6DoorNudgingYes,
                    rb6DoorTracksAluminium, rb6DoorTracksOther, rb6DoorTypeCentreOpening, rb6DoorTypeSideOpening, rb6EmergencyLoweringSystemNo,
                    rb6EmergencyLoweringSystemYes, rb6ExclusiveServiceNo, rb6ExclusiveServiceYes, rb6FacePlateMaterialOther,
                    rb6FacvePlateMaterialStainlessSteel, rb6FalseCeilingNo, rb6FalseCeilingYes, rb6FalseFloorNo, rb6FalseFloorYes, rb6FireServiceNo,
                    rb6FireServiceYes, rb6FrontWallOther, rb6FrontWallStainlessSteel, rb6GPOInCarNo, rb6GPOInCarYes, rb6HandrailOther,
                    rb6HandrailStainlessSteel, rb6IndependentNo, rb6IndependentServiceYes, rb6LandingDoorFinishOther, rb6LandingDoorFinishStainlessSteel,
                    rb6LCDColourBlue, rb6LCDColourRed, rb6LCDColourWhite, rb6LoadWeighingNo, rb6LoadWeighingYes, rb6MirrorFullSize, rb6MirrorHalfSize,
                    rb6MirrorOther, rb6OutOfServiceNo, rb6OutOfServiceYes, rb6PositionIndicatorTypeFlushMount, rb6PositionIndicatorTypeSurfaceMount,
                    rb6ProtectiveBlanketsNo, rb6ProtectiveBlanketsYes, rb6RearDoorKeySwitchNo, rb6RearDoorKeySwitchYes, rb6RearWallOther,
                    rb6RearWallStainlessSteel, rb6SecurityKeySwitchNo, rb6SecurityKeySwitchYes, rb6SideWallOther, rb6SideWallStainlessSteel,
                    rb6StructureShaftCOncrete, rb6StructureShaftOther, rb6TrimmerBeamsNo, rb6TrimmerBeamsYes, rb6VoiceAnnunciationNo,
                    rb6VoiceAnnunciationYes, rb7AdvancedOpeningNo, rb7AdvancedOpeningYes, rb7BumpRailNo, rb7BumpRailYes, rb7CarDoorFinishOther,
                    rb7CarDoorFinishStainlessSteel, rb7CeilingFinishMirrorStainlessSteel, rb7CeilingFinishOther, rb7CeilingFinishStainlessSteel,
                    rb7CEilingFinishWhite, rb7ControlerLocationBottomLAnding, rb7ControlerLocationOther, rb7ControlerLocationShaft,
                    rb7ControlerLocationTopLanding, rb7COPFinishOther, rb7COPFinishStainlessSteel, rb7DoorNudgingNo, rb7DoorNudgingYes,
                    rb7DoorTracksAluminium, rb7DoorTracksOther, rb7DoorTypeCentreOpening, rb7DoorTypeSideOpening, rb7EmergencyLoweringSystemNo,
                    rb7EmergencyLoweringSystemYes, rb7ExclusiveServiceNo, rb7ExclusiveServiceYes, rb7FacePlateMaterialOther,
                    rb7FacePlateMaterialStainlessSteel, rb7FalseCeilingNo, rb7FalseCeilingYes, rb7FalseFloorNo, rb7FalseFloorYes, rb7FireServiceNo,
                    rb7FireSErviceYes, rb7FrontWallOther, rb7FrontWallStainlessSteel, rb7GPOInCarNo, rb7GPOInCarYes, rb7HandrailOther,
                    rb7HandrailStainlessSteel, rb7IndependentServiceNo, rb7IndpendentServiceYes, rb7LandingDoorFinishOther,
                    rb7LandingDoorFinishStainlessSteel, rb7LCDColourBlue, rb7LCDColourRed, rb7LCDColourWhite, rb7LoadWeighingNo, rb7LoadWeighingYes,
                    rb7MirrorFullSize, rb7MirrorHalfSize, rb7MirrorOther, rb7OutOfSErviceNo, rb7OutOfServiceYes, rb7PositionIndicatorTypeFlushMount,
                    rb7PositionIndicatorTypeSurfaceMount, rb7ProtctiveBlanketsNo, rb7ProtectiveBlanketsYes, rb7RearDoorKeySwitchNo,
                    rb7RearDoorKeySwitchYes, rb7RearWallOther, rb7RearWallStainlessSteel, rb7SecurityKeySwitchNo, rb7SecurityKeySwitchYes,
                    rb7SideWallOther, rb7SideWallStainlessSteel, rb7StructureShaftConcrete, rb7StructureShaftOther, rb7TrimmerBeamsNo,
                    rb7TrimmerBeamsYes, rb7VoiceAnnunciationYes, rb7VoiceAnunciationNo, rb8AdvancedOpeningYes, rb8AdvncedOpeningNo, rb8BumpRailNo,
                    rb8BumpRailYes, rb8CarDoorFinishOther, rb8CarDoorFinishStainlessSteel, rb8CeilingFinishMirrorStainlessSTeel, rb8CeilingFinishOther,
                    rb8CeilingFinishStainlessSteel, rb8CeilingFinishWhite, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther,
                    rb8ControlerLocationShaft, rb8ControlerLocationTopLanding, rb8COPFinishOther, rb8COPFinishStainlessSteel, rb8DoorNudgingYes,
                    rb8DoorTracksAluminium, rb8DoorTracksOther, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening, rb8EmergencyLoweringSystemNo,
                    rb8EmergencyLoweringSystemYes, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes, rb8FacePlateMaterialOther,
                    rb8FacePlateMaterialStainlessSteel, rb8FalseCeilingNo, rb8FalseCeilingYes, rb8FalseFloorNo, rb8FalseFloorYes, rb8FireSErviceNo,
                    rb8FireServiceYes, rb8FrontWallOther, rb8FrontWallStainlessSteel, rb8GPOInCarNo, rb8GPOInCarYes, rb8HandrailOther,
                    rb8HandRailStainlessSteel, rb8IndependentServiceNo, rb8IndependentServiceYes, rb8LandingDoorFinishOther, rb8LCDColourBlue,
                    rb8LCDColourRed, rb8LCDColourWhite, rb8LoadWeighingNo, rb8LoadWeighingYes, rb8MirrorFullSize, rb8MirrorHalfSize,
                    rb8MirrorOther, rb8OutOFServiceNo, rb8OutOfSErviceYes, rb8PositionIndicatorTypeFlushMount, rb8PositionIndicatorTypeSurfaceMount,
                    rb8ProtectiveBlanketsNo, rb8ProtectiveBlanketsYes, rb8RearDoorKeySwitchNo, rb8RearDoorKeySwitchYes, rb8RearWallOther,
                    rb8RearWallStainlessSteel, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes, rb8SideWallOther, rb8SideWallStainlessSteel,
                    rb8StructureShaftConcrete, rb8StructureShaftOther, rb8TimmemrBeamsYes, rb8TrimmerBeamsNo, rb8VoiceAnnunicationNo,
                    rb8VoiceAnnunicationYes, rb9AdvancedOpeningNo, rb9AdvancedOpeningYes, rb9BumpRailNo, rb9BumpRailYes, rb9CarDoorFinishOther,
                    rb9CarDoorFinishStainlessSteel, rb9CeilingFinishMirrorStainlessSteel, rb9CeilingFinishOther, rb9CeilingFinishStainlessSteel,
                    rb9CeilingFinishWhite, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther, rb9ControlerLocationShaft,
                    rb9ControlrLocationTopLanding, rb9COPFinishOther, rb9COPFinishStainlessSteel, rb9DoorNudgingNo, rb9DoorNudgingYes,
                    rb9DoorTracksAluminium, rb9DoorTracksOther, rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening, rb9EmergencyLoweringSystemNo,
                    rb9EmergencyLoweringSystemYes, rb9ExclusiveServiceNo, rb9ExclusiveServiceYes, rb9FacePlateMaterialOther,
                    rb9FacePlateMaterialStainlessSteel, rb9FalseCeilingNo, rb9FalseCeilingYes, rb9FalseFloorNo, rb9FalseFloorYes, rb9FireServiceNo,
                    rb9FireSErviceYes, rb9FrontWallOther, rb9FrontWallStainlessSteel, rb9GPOInCarNo, rb9GPOInCarYes, rb9HandrailOther,
                    rb9HandrailStainlessSteel, rb9IndependentServiceNo, rb9IndependentServiceYes, rb9LandingDoorFinishOther,
                    rb9LandingDoorFinishStainlessSteel, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite, rb9LoadWeighingNo, rb9LoadWeighingYes,
                    rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther, rb9OutOfServiceNo, rb9OutOfServiceYes, rb9PositionIndicatorTypeFlushMount,
                    rb9PositionIndicatorTypeSurfaceMount, rb9ProtectiveBlanketsNo, rb9ProtectiveBlanketsYes, rb9RearDoorKeySwitchNo,
                    rb9RearDoorKeySwitchYes, rb9RearWallOther, rb9RearWallStainlessSteel, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes,
                    rb9SideWallOther, rb9SideWallStainlessSteel, rb9StructureShaftConcrete, rb9StructureShaftOther, rb9TrimmerBeamsNo,
                    rb9TrimmerBeamsYes, rb9VoiceAnnunciationNo, rb9VoiceAnnunciationYes, rbAdvancedOpeningNo, rbAdvancedOpeningYes, rbBumpRailNo,
                    rbBumpRailYes, rbCarDoorFInishBrushedStainlessSteel, rbCarDoorFinishOther, rbCeilingFinishBrushedStasinlessSteel,
                    rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther, rbCeilingFinishWhite, rbControlerLoactionTopLanding,
                    rbControlerLocationBottomLanding, rbControlerLocationOther, rbControlerlocationShaft, rbCOPFinishOther, rbCOPFinishSatinStainlessSteel,
                    rbDoorNudgingNo, rbDoorNudgingYes, rbDoorTracksAnodisedAluminium, rbDoorTracksOther, rbDoorTypeCentreOpening,
                    rbDoorTypeSideOpening, rbEmergencyLoweringSystemNo, rbEmergencyLoweringSystemYes, rbExclusiveServiceNo, rbExclusiveServiceYes,
                    rbFacePlateMaterialOther, rbFacePlateMaterialSatinStainlessSteel, rbFalseCeilingNo, rbFalseCeilingYes, rbFalseFloorNo, rbFalseFloorYes,
                    rbFireServiceNo, rbFireServiceYes, rbFrontWallBrushedStainlessSteel, rbGPOInCarNo, rbGPOInCarYes, rbHandrailBrushedStainlessSTeel,
                    rbHandrailOther, rbIndependentServiceNo, rbIndependentServiceYes, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel,
                    rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite, rbLoadWeighingNo, rbLoadWeighingYes, rbMirrorFullSize, rbMirrorHalfSize,
                    rbMirrorOther, rbOutofServiceNo, rbOutofServiceYes, rbPositionIndicatorTypeFlushMount, rbPositionIndicatorTypeSurfaceMount,
                    rbProtectiveBlanketsNo, rbProtectriveBlanketsYes, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes, rbRearWallBrushedStainlessSteel,
                    rbRearWallOther, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes, rbSideWallBrushedStainlessSteel, rbSideWallOther, rbSL,
                    rbStructureShaftConcrete, rbStructureShaftOther, rbSumasa, rbTrimmerBeamsNo, rbTrimmerBeamsYes, rbVoiceAnnunciationNo,
                    rbVoiceAnnunciationYes, rbWittur);
                #endregion
            }
            catch (Exception)
            {
                return;
            }
        }

        #endregion

        #region Save Quote Data Methods

        private void SaveData()
        {
            string prefix = "P0";
            saveData["NumberOfPagesOpen"] = pageTracker.ToString();
            // QuoteInfo3 nF = new QuoteInfo3();
            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 

            if (pageTracker >= 1)
            {
                prefix = "P1";
                #region Page 1 Saving
                SaveTbToXML(tbAuxCOPLocation, tbCOPFinish, tbDesignations, tbKeyswitchLocation, tbMainCOPLocation, tbNumberOfCOPS,
                    tbCarDoorFinish, tbCeilingFinish, tbFloorFinish, tbFrontWall, tbHandrail, tbMirror, tbNumOfLEDLights, tbRearWall, tbSideWall,
                    tbDoorHeight, tbDoorTracks, tbDoorWidth, tbLandingDoorFinish, tbDepth, tbHeight, tbLiftCarNotes, tbLiftRating, tbLoad, tbNumofCarEntrances, tbSpeed,
                    tbwidth, tbHeadroom, tbNumofLandingDoors, tbNumofLandings, tbPitDepth, tbTypeofLift, tbLiftNumbers,
                    tbShaftDepth, tbShaftWidth, tbStructureShaft, tbTravel, tbControlerLocation, tbfname, tblname, tbphone, tbAddress1, tbAddress2, tbAddress3
                    );

                SaveRbToXML(rbCOPFinishOther, rbCOPFinishSatinStainlessSteel, rbExclusiveServiceNo, rbExclusiveServiceYes,
                     rbGPOInCarNo, rbGPOInCarYes, rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite, rbPositionIndicatorTypeFlushMount,
                    rbPositionIndicatorTypeSurfaceMount, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes,
                    rbVoiceAnnunciationNo, rbVoiceAnnunciationYes, tbFrontWallOther, rbBumpRailNo, rbBumpRailYes, rbCarDoorFInishBrushedStainlessSteel, rbCarDoorFinishOther,
                    rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther, rbCeilingFinishWhite, rbFalseCeilingNo,
                    rbFalseCeilingYes, rbFrontWallBrushedStainlessSteel, rbHandrailBrushedStainlessSTeel, rbHandrailOther, rbMirrorFullSize, rbMirrorHalfSize,
                    rbMirrorOther, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes, rbRearWallBrushedStainlessSteel, rbRearWallOther,
                    rbSideWallBrushedStainlessSteel, rbSideWallOther, rbAdvancedOpeningNo, rbAdvancedOpeningYes, rbDoorNudgingNo, rbDoorNudgingYes,
                    rbDoorTracksAnodisedAluminium, rbDoorTracksOther, rbDoorTypeCentreOpening, rbDoorTypeSideOpening, rbLandingDoorFinishOther,
                    rbLandingDoorFinishStainlessSteel, rbFalseFloorNo, rbFalseFloorYes, rbStructureShaftConcrete, rbStructureShaftOther, rbTrimmerBeamsNo,
                    rbTrimmerBeamsYes, rbControlerLoactionTopLanding, rbControlerLocationBottomLanding, rbControlerLocationOther, rbControlerlocationShaft,
                    rbFireServiceNo, rbFireServiceYes, rbIndependentServiceNo, rbIndependentServiceYes, rbLoadWeighingNo, rbLoadWeighingYes,
                    rbControlerLoactionTopLanding, rbControlerLocationBottomLanding, rbControlerLocationOther, rbControlerlocationShaft,
                    rbFireServiceNo, rbFireServiceYes, rbIndependentServiceNo, rbIndependentServiceYes, rbLoadWeighingNo, rbLoadWeighingYes,
                    rbSL, rbSumasa, rbWittur
                    );
                #endregion
                #region Page 1 Word Export
                WordData("AE105", tbfname.Text); //first name
                WordData("AE106", tblname.Text);//last name
                WordData("AE107", tbphone.Text);//phone number
                WordData("AE108", tbAddress1.Text);//address 1
                WordData("AE109", tbAddress2.Text);//address 2
                WordData("AE110", tbAddress3.Text);//address 3
                WordData(prefix + "AE111", RadioButtonHandeler(null, rbSL, rbWittur, rbSumasa)); //supplier
                WordData("AE114", TotalLifts().ToString());//lift number
                WordData("AE115", tbTypeofLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE118", RadioButtonHandeler(tbControlerLocation, rbControlerLoactionTopLanding, rbControlerlocationShaft, rbControlerLocationBottomLanding, rbControlerLocationOther));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rbFireServiceNo, rbFireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tbShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tbShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tbPitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tbHeadroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tbTravel.Text, "mm"));//travel
                WordData(prefix + "AE128", tbNumofLandings.Text);// number of landings
                WordData(prefix + "AE129", tbNumofLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tbStructureShaft, rbStructureShaftConcrete, rbStructureShaftOther)); //structure shaft 
                WordData(prefix + "AE134", MeasureStringChecker(tbLoad.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tbSpeed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tbwidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tbDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tbHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tbLiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tbLiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tbNumofCarEntrances.Text);//number of car entraces
                if (tbLiftCarNotes.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tbLiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tbDoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tbDoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tbLandingDoorFinish, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rbDoorTypeCentreOpening, rbDoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tbDoorTracks, rbDoorTracksAnodisedAluminium, rbDoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rbAdvancedOpeningNo, rbAdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rbDoorNudgingNo, rbDoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tbCarDoorFinish, rbCarDoorFinishOther, rbCarDoorFInishBrushedStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tbCeilingFinish, rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishWhite, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rbFalseCeilingNo, rbFalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rbBumpRailYes, rbBumpRailNo));//bump rail
                WordData(prefix + "AE159", tbFloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tbFrontWall, rbFrontWallBrushedStainlessSteel, tbFrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tbMirror, rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tbHandrail, rbHandrailBrushedStainlessSTeel, rbHandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tbSideWall, rbSideWallBrushedStainlessSteel, rbSideWallOther)); //side wall 
                WordData(prefix + "AE165", tbNumOfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE216", RadioButtonHandeler(tbRearWall, rbRearWallOther, rbRearWallBrushedStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tbNumberOfCOPS.Text); // number of COPS
                WordData(prefix + "AE169", tbMainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tbAuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tbDesignations.Text); // designations 
                WordData(prefix + "AE191", tbKeyswitchLocation.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tbCOPFinish, rbCOPFinishSatinStainlessSteel));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rbLEDColourRed, rbLEDColourBlue, rbLEDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rbExclusiveServiceNo, rbExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rbGPOInCarNo, rbGPOInCarYes));//GPO in car
                WordData(prefix + "AE190", RadioButtonHandeler(null, rbPositionIndicatorTypeSurfaceMount, rbPositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tbFacePlateMaterial, rbFacePlateMaterialSatinStainlessSteel, rbFacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rbEmergencyLoweringSystemYes, rbEmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE178", CheckboxTrueToYes(cbMainSecurity));//security cabiling only 

                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rbIndependentServiceYes, rbIndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rbLoadWeighingYes, rbLoadWeighingNo));//load weighing
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rbFalseFloorYes, rbFalseFloorNo));//false floor
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rbTrimmerBeamsYes, rbTrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rbProtectriveBlanketsYes, rbProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rbVoiceAnnunciationYes, rbVoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rbOutofServiceYes, rbOutofServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 2)
            {
                prefix = "P2";
                #region Page 2 Saving
                SaveTbToXML(tb2AuxCOPLocation, tb2CarDepth, tb2CarDoorFinishText,
                    tb2CarHeight, tb2CarLoad, tb2CarWidth, tb2CeilingFinishText, tb2ControlerLocationText, tb2COPFinishText, tb2Designations,
                    tb2DoorHeight, tb2DoorTracksText, tb2DoorWidth, tb2FacePlateMaterialText, tb2FloorFinish, tb2FrontWallText,
                    tb2HandrailText, tb2Headroom, tb2KeyswitchLocation, tb2LandingDoorFinishText, tb2LiftNumbers, tb2LiftRating,
                    tb2MainCOPLocation, tb2MirrorText, tb2Note, tb2NumberOfCarEntrances, tb2NumberOfCOPs, tb2NumberOfLAndingDoors,
                    tb2NumberOfLandings, tb2NumberofLEDLights, tb2PitDepth, tb2RearWallText, tb2ShaftDepth, tb2ShaftWidth,
                    tb2SideWallText, tb2Speed, tb2StructureShaftText, tb2Travel, tb2TypeOfLift
                    );

                SaveRbToXML(rb2AdvancedOpeningNo, rb2AdvancedOpeningYes, rb2BumpRailNo, rb2BumpRailYes, rb2CarDoorFinishOther,
                    rb2CarDoorFinishStainlessSteel, rb2CeilingFinishMirrorStainlessSteel, rb2CeilingFinishOther, rb2CeilingFinishStainlessSteel,
                    rb2CeilingFinishWhite, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther, rb2ControlerLocationShaft,
                    rb2ControlerLocationTopLanding, rb2COPFinishOther, rb2COPFinishStainlessSTeell, rb2DoorNudgingNo, rb2DoorNudgingYes,
                    rb2DoorTracksAluminium, rb2DoorTracksOther, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening,
                    rb2EmergemncyLoweringSystemNo, rb2EmergencyLoweringSystemYes, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes,
                    rb2FacePlateMaterialOther, rb2FacePlateMaterialStainlessSteel, rb2FalseCeilingNo, rb2FalseCeilingYes, rb2FalseFloorNo,
                    rb2FalseFloorYes, rb2FireSErviceNo, rb2FireSErviceYes, rb2FrontWallOther, rb2FrontWallStainlessSteel, rb2GPOInCarNo,
                    rb2GPOInCarYes, rb2HandrailOther, rb2HandRailStainlessSteel, rb2IndependentServiceNo, rb2IndependentServiceYes,
                    rb2LandingDoorFinishOther, rb2LandingDoorFinishStainlessSteel, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite,
                    rb2LoadWeighingNo, rb2LoadWeighingYes, rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther, rb2OutOfServiceNo,
                    rb2OutOfServiceYes, rb2PositionIndicatorTypeFlushMount, rb2PositionIndicatorTypeSurfaceMount, rb2ProtectiveBlanketsNo,
                    rb2ProtectiveBlanketsYes, rb2RearDoorKeySwitchNo, rb2RearDoorKeySwitchYes, rb2RearWallOther, rb2RearWallStainlessSteel,
                    rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes, rb2SideWallOther, rb2SideWallStainlessSteel, rb2StructureShaftConcrete,
                    rb2StructureShaftOther, rb2TrimmerBeamsNo, rb2TrimmerBeamsYes,
                    rb2VoiceAnnunciationNo, rb2VoiceAnnunciationYes
                    );
                #endregion
                #region Page 2 Word Export
                WordData(prefix + "AE114", tb2LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb2TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb2IndependentServiceYes, rb2IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb2LoadWeighingYes, rb2LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb2ControlerLocationText, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther, rb2ControlerLocationShaft, rb2ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb2FireSErviceNo, rb2FireSErviceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb2ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb2ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb2PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb2Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb2Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb2NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb2NumberOfLAndingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb2StructureShaftText, rb2StructureShaftConcrete, rb2StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb2TrimmerBeamsYes, rb2TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb2FalseFloorYes, rb2FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb2CarLoad.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb2Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb2CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb2CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb2CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb2LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb2LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb2NumberOfCarEntrances.Text);//number of car entraces
                if (tb2Note.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb2Note.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb2DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb2DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb2LandingDoorFinishText, rb2LandingDoorFinishOther, rb2LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb2DoorTracksText, rb2DoorTracksAluminium, rb2DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb2AdvancedOpeningNo, rb2AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb2DoorNudgingNo, rb2DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb2CarDoorFinishText, rb2CarDoorFinishOther, rb2CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb2CeilingFinishText, rb2CeilingFinishStainlessSteel, rb2CeilingFinishWhite, rb2CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb2FalseCeilingNo, rb2FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb2BumpRailYes, rb2BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb2FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb2FrontWallText, rb2FrontWallStainlessSteel, rb2FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb2MirrorText, rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb2HandrailText, rb2HandRailStainlessSteel, rb2HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb2SideWallText, rb2SideWallStainlessSteel, rb2SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb2NumberofLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb2ProtectiveBlanketsYes, rb2ProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb2RearWallText, rb2RearWallOther, rb2RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb2NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb2MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb2AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb2Designations.Text); // designations 
                WordData(prefix + "AE191", tb2KeyswitchLocation.Text); //key switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb2COPFinishText, rb2COPFinishStainlessSTeell, rb2COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb2RearDoorKeySwitchNo, rb2RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb2GPOInCarNo, rb2GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb2VoiceAnnunciationYes, rb2VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb2PositionIndicatorTypeSurfaceMount, rb2PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb2FacePlateMaterialText, rb2FacePlateMaterialStainlessSteel, rb2FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb2EmergencyLoweringSystemYes, rb2EmergemncyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb2OutOfServiceYes, rb2OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 3)
            {
                prefix = "P3";
                #region Page 3 Saving
                SaveTbToXML(tb3AuxCOPLocation, tb3CarDepth, tb3CarDoorFinishText,
                    tb3CarHeight, tb3CarNote, tb3CarWidth, tb3CEilingFinishText, tb3ControlerLocationText, tb3COPFinishText, tb3Designations,
                    tb3DoorHeight, tb3DoorTracksText, tb3DoorWidth, tb3FacePlaterMaterialText, tb3FloorFinish, tb3FrontWallText,
                    tb3HandrailText, tb3HeadRoom, tb3KeyswitchLocation, tb3LandingDoorFinishText, tb3LiftNumbers, tb3LiftRating,
                    tb3Load, tb3MainCOPLocation, tb3MirrorText, tb3NumberOfCarEntrances, tb3NumberOfCOPs, tb3NumberOfLandingDoors,
                    tb3NumberOfLandings, tb3NumberOfLEDLights, tb3PitDepth, tb3RearWallText, tb3ShaftDepth, tb3ShaftWidth,
                    tb3SideWallText, tb3Speed, tb3StructureShaftText, tb3Travel, tb3TypeOfLift
                    );

                SaveRbToXML(rb3AdvancedOpeningNo, rb3AdvancedOpeningYes, rb3BumpRailNo, rb3BumpRailYes, rb3CarDoorFinishOther,
                    rb3CarDoorFinishStainlessSteel, rb3CeilingFinishOther, rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite,
                    rb3ControleRLocationBottomLanding, rb3ControlerLocationOther, rb3ControlerLocationShaft, rb3ControlerLocationTopLanding,
                    rb3COPFinishOther, rb3COPFinishStainlessSteel, rb3DoorNudgingNo, rb3DoorNudgingYes, rb3DoorTracksAluminium,
                    rb3DoorTracksOther, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening, rb3EmergencyLoweringSystemNo,
                    rb3EmergencyLoweringSystemYes, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes, rb3FacePlateMaterialOther,
                    rb3FacePlateMaterialStainlessSteel, rb3FalseCeilingNo, rb3FalseCeilingYes, rb3FalseFloorNo, rb3FalseFloorYes,
                    rb3FireServiceYes, rb3FireServieNo, rb3FrontWallOther, rb3FrontWallStainlessSteel, rb3GPOInCarNo, rb3GPOInCarYes,
                    rb3HandrailOther, rb3HandrailStainlessSteel, rb3IndependentServiceNo, rb3IndependentServiceYes, rb3landingDoorFinishOther,
                    rb3LandingDoorFinishStainlessSteel, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite, rb3LoadWeighingNo,
                    rb3LoadWeighingYes, rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther, rb3MirrorStainlessSteel, rb3OutOfServiceNo,
                    rb3OutOfSErviceYes, rb3PositionIndicatorTypeFlushMount, rb3PositionIndicatorTypeSurfaceMount, rb3ProtectiveBlanketsNo,
                    rb3ProtectiveBlanketsYes, rb3RearDoorKeySwitchNo, rb3RearDoorKeySwitchYes, rb3RearWallOther, rb3RearWallStainlessSteel,
                    rb3SecurityKeySwitchNo, rb3SecurityKeySwitchYes, rb3SideWallOther, rb3SideWallStainlessSteel, rb3StructureShaftConcrete,
                    rb3StructureShaftOther, rb3TrimmerBeamsNo, rb3TrimmerBeamsYes,
                    rb3VoiceAnnunciationNo, rb3VoiceAnnunciationYes
                    );
                #endregion
                #region Page 3 Word Export
                WordData(prefix + "AE114", tb3LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb3TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb3IndependentServiceYes, rb3IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb3LoadWeighingYes, rb3LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb3ControlerLocationText, rb3ControleRLocationBottomLanding, rb3ControlerLocationOther, rb3ControlerLocationShaft, rb3ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb3FireServieNo, rb3FireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb3ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb3ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb3PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb3HeadRoom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb3Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb3NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb3NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb3StructureShaftText, rb3StructureShaftConcrete, rb3StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb3TrimmerBeamsYes, rb3TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb3FalseFloorYes, rb3FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb3Load.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb3Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb3CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb3CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb3CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb3LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb3LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb3NumberOfCarEntrances.Text);//number of car entraces
                if (tb3CarNote.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb3CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb3DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb3DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb3LandingDoorFinishText, rb3landingDoorFinishOther, rb3LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb3DoorTracksText, rb3DoorTracksAluminium, rb3DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb3AdvancedOpeningNo, rb3AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb3DoorNudgingNo, rb3DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb3CarDoorFinishText, rb3CarDoorFinishOther, rb3CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb3CEilingFinishText, rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite, rb3MirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb3FalseCeilingNo, rb3FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb3BumpRailYes, rb3BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb3FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb3FrontWallText, rb3FrontWallStainlessSteel, rb3FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb3MirrorText, rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb3HandrailText, rb3HandrailStainlessSteel, rb3HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb3SideWallText, rb3SideWallStainlessSteel, rb3SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb3NumberOfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb3ProtectiveBlanketsYes, rb3ProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb3RearWallText, rb3RearWallOther, rb3RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb3NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb3MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb3AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb3Designations.Text); // designations 
                WordData(prefix + "AE191", tb3KeyswitchLocation.Text); //key switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb3COPFinishText, rb3COPFinishStainlessSteel, rb3COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb3RearDoorKeySwitchNo, rb3RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb3SecurityKeySwitchNo, rb3SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb3GPOInCarNo, rb3GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb3VoiceAnnunciationYes, rb3VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb3PositionIndicatorTypeSurfaceMount, rb3PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb3FacePlaterMaterialText, rb3FacePlateMaterialStainlessSteel, rb3FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb3EmergencyLoweringSystemYes, rb3EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb3OutOfSErviceYes, rb3OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 4)
            {
                prefix = "P4";
                #region Page 4 Saving
                SaveTbToXML(tb4AuxCOPLocation, tb4CarDepth, tb4CarDoorFinish,
                    tb4CarHeight, tb4CarNote, tb4CarWidth, tb4CeilingFinishText, tb4ControlerLocationText, tb4COPFinishText, tb4Designations,
                    tb4DoorHeight, tb4DoorTracksText, tb4DoorWidth, tb4FacePlateMaterialText, tb4FloorFinish, tb4FrontWallText,
                    tb4HandrailText, tb4Headroom, tb4KeyswitchLocations, tb4LandingDoorFinishText, tb4LiftNumbers, tb4LiftRating,
                    tb4Load, tb4MainCOPLocation, tb4MirrorText, tb4NumberOfCarEntrances, tb4NumberOfCOPs, tb4NumberOfLandingDoors,
                    tb4NumberOfLandings, tb4NumbeROfLEDLights, tb4PitDepth, tb4RearWallText, tb4ShaftDepth, tb4ShaftWidth, tb4SideWallText,
                    tb4Speed, tb4StructureShaftText, tb4Travel, tb4TypeOfLift
                    );

                SaveRbToXML(rb4AdvancedOpeningNo, rb4AdvancedOpeningYes, rb4BumpRailNo, rb4BumpRailYes, rb4CarDoorFinishOther,
                    rb4CarDoorFinishStainlessSteel, rb4CeilingFinishMirrorStainlessSteel, rb4CeilingFinishOther, rb4CeilingFinishStainlessSteel,
                    rb4CeilingFinishWhite, rb4ControelrLocationTopLanding, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther,
                    rb4ControlerLocationShaft, rb4COPFinishOther, rb4COPFinishStainlessSteel, rb4DoorNudgingNo, rb4DoorNudgingYes,
                    rb4DoorTracksAluminium, rb4DoorTracksOther, rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening, rb4EmergencyLoweringSystemNo,
                    rb4EmergencyLoweringSystemYes, rb4ExclusiveServiceNo, rb4ExclusiveServiceYes, rb4FacePlateMaterialOther,
                    rb4FacePlateMaterialStainlessSteel, rb4FalseCeilingNO, rb4FalseCeilingYes, rb4FalseFloorNo, rb4FalseFloorYes, rb4FireSErviceNo,
                    rb4FireServiceYes, rb4FrontWallStainlessSteel, rb4FrotnWallOther, rb4GPOInCarNo, rb4GPOInCarYes, rb4HandrailOther,
                    rb4HandRailStainlesSteel, rb4IndependentServiceYes, rb4LandingDoorFinishOther, rb4LandingDoorFinishStainlessSteel, rb4LCDColourBlue,
                    rb4LCDColourRed, rb4LCDColourWhite, rb4LoadWeighingNo, rb4LoadWeighingYes, rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther,
                    rb4OutOfServiceNo, rb4OutOfServiceYes, rb4PositionIndicatorTypeFlushMount, rb4PositionIndicatorTypeSurfaceMount,
                    rb4ProtectiveBlanketsYes, rb4ProtetiveBlanketsNo, rb4RearDoorKeySwitchNo, rb4RearDoorKeySwitchYes, rb4RearWallOther,
                    rb4RearWallStainlessSteel, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes, rb4SideWallOther, rb4SideWallStainlessSteel,
                    rb4StructureShaftConcrete, rb4StructureShaftOther, rb4TrimmerBeamsNo,
                    rb4TrimmerBeamsYes, rb4VoiceAnnunciationNo, rb4VoiceAnnunciationYes
                    );
                #endregion
                #region Page 4 Word Export
                WordData(prefix + "AE114", tb4LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb4TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb4IndependentServiceYes, IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb4LoadWeighingYes, rb4LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb4ControlerLocationText, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther, rb4ControlerLocationShaft, rb4ControelrLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb4FireSErviceNo, rb4FireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb4ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb4ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb4PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb4Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb4Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb4NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb4NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb4StructureShaftText, rb4StructureShaftConcrete, rb4StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb4TrimmerBeamsYes, rb4TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb4FalseFloorYes, rb4FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb4Load.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb4Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb4CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb4CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb4CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb4LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb4LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb4NumberOfCarEntrances.Text);//number of car entraces
                if (tb4CarNote.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb4CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb4DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb4DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb4LandingDoorFinishText, rb4LandingDoorFinishOther, rb4LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb4DoorTracksText, rb4DoorTracksAluminium, rb4DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb4AdvancedOpeningNo, rb4AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb4DoorNudgingNo, rb4DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb4CarDoorFinish, rb4CarDoorFinishOther, rb4CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb4CeilingFinishText, rb4CeilingFinishStainlessSteel, rb4CeilingFinishWhite, rb4CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb4FalseCeilingNO, rb4FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb4BumpRailYes, rb4BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb4FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb4FrontWallText, rb4FrontWallStainlessSteel, rb4FrotnWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb4MirrorText, rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb4HandrailText, rb4HandRailStainlesSteel, rb4HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb4SideWallText, rb4SideWallStainlessSteel, rb4SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb4NumbeROfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb4ProtectiveBlanketsYes, rb4ProtetiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb4RearWallText, rb4RearWallOther, rb4RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb4NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb4MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb4AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb4Designations.Text); // designations 
                WordData(prefix + "AE191", tb4KeyswitchLocations.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb4COPFinishText, rb4COPFinishStainlessSteel, rb4COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb4LCDColourBlue, rb4LCDColourRed, rb4LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb4ExclusiveServiceNo, rb4ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb4RearDoorKeySwitchNo, rb4RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb4GPOInCarNo, rb4GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb4VoiceAnnunciationYes, rb4VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb4PositionIndicatorTypeSurfaceMount, rb4PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb4FacePlateMaterialText, rb4FacePlateMaterialStainlessSteel, rb4FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb4EmergencyLoweringSystemYes, rb4EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb4OutOfServiceYes, rb4OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 5)
            {
                prefix = "P5";
                #region Page 5 Saving
                SaveTbToXML(tb5AuxCOPLocation, tb5CarDepth, tb5CarDoorFinishText,
                    tb5CaRHeight, tb5CarNote, tb5CarWidth, tb5CeilingFinishText, tb5ControlerLocationText, tb5COPFinishText, tb5Designations,
                    tb5DoorHeight, tb5DoorTRacksText, tb5DoorWidth, tb5FacePlateMaterialText, tb5FloorFinish, tb5FrontWallText,
                    tb5HandrailText, tb5Headroom, tb5KetyswitchLocation, tb5LandingDoorFinishText, tb5LiftNumbers, tb5LiftRating,
                    tb5Load, tb5MainCOPLocation, tb5MirrorText, tb5NumberOfCarEntrances, tb5NumberOfCOPs, tb5NumberOfLandingDoors,
                    tb5NumberOfLandings, tb5NumberOIfLEDLights, tb5PitDepth, tb5RearWallText, tb5ShaftDEpth, tb5ShaftWidth,
                    tb5SideWallText, tb5Speed, tb5StructureShaftText, tb5Travel, tb5TypeOfLift
                    );

                SaveRbToXML(rb5AdvancedOpeningNo, rb5AdvancedOpeningYes, rb5BumpRailNo, rb5BumpRailYes, rb5CarDoorFinishOther,
                    rb5CarDoorFinishStainlessSteel, rb5CeilingFinishMirrorStainlessSTeel, rb5CeilingFinishOther, rb5CeilingFinishStainlessSteel,
                    rb5CeilingFinishWhite, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther, rb5ControlerLocationShaft,
                    rb5ControlerLocationTopLanding, rb5COPFinishOther, rb5COPFinishStainlessSteel, rb5DoorNudgingNo, rb5DoorNudgingYes,
                    rb5DoorTracksAluminium, rb5DoorTracksOther, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening, rb5EmergencyLoweringSystemNo,
                    rb5EmergencyLoweringSystemYes, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes, rb5FacePlateMAterialOtjer,
                    rb5FacePlateMaterialStainlessSteel, rb5FalseCeilingNo, rb5FalseCeilingYes, rb5FalseFloorNo, rb5FalseFloorYes, rb5FireSErviceNo,
                    rb5FireServiceYes, rb5FrontWallOther, rb5FrontWallStainlessSteel, rb5GPOInCarNo, rb5GPOInCarYes, rb5HandRailOther,
                    rb5HandRailStainlessSteel, rb5IndependentServiceNo, rb5IndependentServiceYes, rb5LAndingDoorFinishOther,
                    rb5LandingDoorFinishStainlessSteel, rb5LCDColoiurWHite, rb5LCDColourBlue, rb5LCDColourRed, rb5LoadWeighingNo,
                    rb5LoadWeighingYes, rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther, rb5OutOfServiceNo, rb5OutOfServiceYes,
                    rb5PositionIndicatorTypeSurfaceMount, rb5PostionIndicatorTypeFlushMount, rb5ProtectiveBlanketsNo, rb5ProtectiveBlanketsYes,
                    rb5RearDoorKeySwitchNo, rb5RearDoorKeySwitchYes, rb5RearWallOther, rb5RearWallStainlesSteel, rb5SecurityKeySwitchNo,
                    rb5SecurityServiceYes, rb5SideWallOther, rb5SideWallStainlessSTeel, rb5StructureShaftConcrete, rb5StructureShaftOther,
                       rb5TrimmerBeamsNo, rb5TrimmerBeamsYes,
                    rb5VoiceAnnunciationNo, rb5VoiceAnnunciationYes
                    );
                #endregion
                #region Page 5 Word Export
                WordData(prefix + "AE114", tb5LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb5TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb5IndependentServiceYes, rb5IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb5LoadWeighingYes, rb5LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb5ControlerLocationText, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther, rb5ControlerLocationShaft, rb5ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb5FireSErviceNo, rb5FireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb5ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb5PitDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb5PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb5Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb5Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb5NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb5NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb5StructureShaftText, rb5StructureShaftConcrete, rb5StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb5TrimmerBeamsYes, rb5TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb5FalseFloorYes, rb5FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb5Load.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb5Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb5CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb5CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb5CaRHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb5LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb5LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb5NumberOfCarEntrances.Text);//number of car entraces
                if (tb5CarNote.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb5CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb5DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb5DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb5LandingDoorFinishText, rb5LAndingDoorFinishOther, rb5LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb5DoorTRacksText, rb5DoorTracksAluminium, rb5DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb5AdvancedOpeningNo, rb5AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb5DoorNudgingNo, rb5DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb5CarDoorFinishText, rb5CarDoorFinishOther, rb5CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb5CeilingFinishText, rb5CeilingFinishStainlessSteel, rb5CeilingFinishWhite, rb5CeilingFinishMirrorStainlessSTeel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb5FalseCeilingNo, rb5FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb5BumpRailYes, rb5BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb5FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb5FrontWallText, rb5FrontWallStainlessSteel, rb5FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb5MirrorText, rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb5HandrailText, rb5HandRailOther, rb5HandRailStainlessSteel));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb5SideWallText, rb5SideWallStainlessSTeel, rb5SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb5NumberOIfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb5ProtectiveBlanketsYes, rb5ProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb5RearWallText, rb5RearWallOther, rb5RearWallStainlesSteel)); //  rear wall
                WordData(prefix + "AE168", tb5NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb5MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb5AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb5Designations.Text); // designations 
                WordData(prefix + "AE191", tb5KetyswitchLocation.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb5COPFinishText, rb5COPFinishStainlessSteel, rb5COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb5LCDColourBlue, rb5LCDColourRed, rb5LCDColoiurWHite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb5RearDoorKeySwitchNo, rb5RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb5SecurityKeySwitchNo, rb5SecurityServiceYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb5GPOInCarNo, rb5GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb5VoiceAnnunciationYes, rb5VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb5PositionIndicatorTypeSurfaceMount, rb5PostionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb5FacePlateMaterialText, rb5FacePlateMaterialStainlessSteel, rb5FacePlateMAterialOtjer));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb5EmergencyLoweringSystemYes, rb5EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb5OutOfServiceYes, rb5OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 6)
            {
                prefix = "P6";
                #region Page 6 Saving
                SaveTbToXML(tb6AuxCOPLocation, tb6CarDepth, tb6CarDoorFinishText,
                    tb6CarHeight, tb6CarLoad, tb6CarNote, tb6CarSpeed, tb6CarWidth, tb6CeilingFinishText, tb6ControlerLocationText, tb6COPFinishText,
                    tb6Designations, tb6DoorHeight, tb6DoorTracksOther, tb6DoorWidth, tb6FacePlateMaterialText, tb6FloorFinish,
                    tb6FrontWallText, tb6HAndrailText, tb6Headroom, tb6KeySwitchLocation, tb6LandingDoorFinishText, tb6LiftNumbers,
                    tb6LiftRating, tb6MainCOPLocation, tb6MirrorText, tb6NumberOfCarEntrances, tb6NumberOFCOPs, tb6NumberOfLandingDoors,
                    tb6NumberOfLandings, tb6NumberOfLEDLights, tb6PitDepth, tb6RearWallText, tb6ShaftDepth, tb6ShaftWidth,
                    tb6SideWallText, tb6StructureShaftText, tb6Travel, tb6TypeOfLift
                    );

                SaveRbToXML(rb6AdvancedOpeningNo, rb6AdvancedOpeningYes, rb6BumpRailNo, rb6BumpRailYes, rb6CarDoorFinishOther,
                    rb6CarDoorFinishStainlessSteel, rb6CeilingFinishMirrorStainlessSteel, rb6CeilingFinishOther, rb6CeilingFinishStainlessSteel,
                    rb6CeilingFinishWhite, rb6ControlerLocationBottomLanding, rb6ControlerLocationOther, rb6ControlerLocationShaft,
                    rb6ControlerLocationTopLanding, rb6COPFinishOther, rb6COPFinishStainlessSteel, rb6DoorNudgingNo, rb6DoorNudgingYes,
                    rb6DoorTracksAluminium, rb6DoorTracksOther, rb6DoorTypeCentreOpening, rb6DoorTypeSideOpening, rb6EmergencyLoweringSystemNo,
                    rb6EmergencyLoweringSystemYes, rb6ExclusiveServiceNo, rb6ExclusiveServiceYes, rb6FacePlateMaterialOther,
                    rb6FacvePlateMaterialStainlessSteel, rb6FalseCeilingNo, rb6FalseCeilingYes, rb6FalseFloorNo, rb6FalseFloorYes, rb6FireServiceNo,
                    rb6FireServiceYes, rb6FrontWallOther, rb6FrontWallStainlessSteel, rb6GPOInCarNo, rb6GPOInCarYes, rb6HandrailOther,
                    rb6HandrailStainlessSteel, rb6IndependentNo, rb6IndependentServiceYes, rb6LandingDoorFinishOther, rb6LandingDoorFinishStainlessSteel,
                    rb6LCDColourBlue, rb6LCDColourRed, rb6LCDColourWhite, rb6LoadWeighingNo, rb6LoadWeighingYes, rb6MirrorFullSize, rb6MirrorHalfSize,
                    rb6MirrorOther, rb6OutOfServiceNo, rb6OutOfServiceYes, rb6PositionIndicatorTypeFlushMount, rb6PositionIndicatorTypeSurfaceMount,
                    rb6ProtectiveBlanketsNo, rb6ProtectiveBlanketsYes, rb6RearDoorKeySwitchNo, rb6RearDoorKeySwitchYes, rb6RearWallOther,
                    rb6RearWallStainlessSteel, rb6SecurityKeySwitchNo, rb6SecurityKeySwitchYes, rb6SideWallOther, rb6SideWallStainlessSteel,
                    rb6StructureShaftCOncrete, rb6StructureShaftOther, rb6TrimmerBeamsNo,
                    rb6TrimmerBeamsYes, rb6VoiceAnnunciationNo, rb6VoiceAnnunciationYes
                    );
                #endregion
                #region Page 6 Word Export
                WordData(prefix + "AE114", tb6LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb6TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb6IndependentServiceYes, rb6IndependentNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb6LoadWeighingYes, rb6LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb6ControlerLocationText, rb6ControlerLocationOther, rb6ControlerLocationBottomLanding, rb6ControlerLocationShaft, rb6ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb6FireServiceNo, rb6FireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb6ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb6ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb6PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb6Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb6Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb6NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb6NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb6StructureShaftText, rb6StructureShaftCOncrete, rb6StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb6TrimmerBeamsYes, rb6TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb6FalseFloorYes, rb6FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb6CarLoad.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb6CarSpeed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb6CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb6CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb6CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb6LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb6LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb6NumberOfCarEntrances.Text);//number of car entraces
                if (tb6CarNote.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb6CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb6DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb6DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb6LandingDoorFinishText, rb6LandingDoorFinishOther, rb6LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb6DoorTypeCentreOpening, rb6DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb6DoorTracksOther, rb6DoorTracksAluminium, rb6DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb6AdvancedOpeningNo, rb6AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb6DoorNudgingNo, rb6DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb6CarDoorFinishText, rb6CarDoorFinishOther, rb6CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb6CeilingFinishText, rb6CeilingFinishStainlessSteel, rb6CeilingFinishWhite, rb6CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb6FalseCeilingNo, rb6FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb6BumpRailYes, rb6BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb6FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb6FrontWallText, rb6FrontWallStainlessSteel, rb6FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb6MirrorText, rb6MirrorFullSize, rb6MirrorHalfSize, rb6MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb6HAndrailText, rb6HandrailStainlessSteel, rb6HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb6SideWallText, rb6SideWallStainlessSteel, rb6SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb6NumberOfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb6ProtectiveBlanketsYes, rb6ProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb6RearWallText, rb6RearWallOther, rb6RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb6NumberOFCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb6MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb6AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb6Designations.Text); // designations 
                WordData(prefix + "AE191", tb6KeySwitchLocation.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb6COPFinishText, rb6COPFinishStainlessSteel, rb6COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb6LCDColourBlue, rb6LCDColourRed, rb6LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb6ExclusiveServiceNo, rb6ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb6RearDoorKeySwitchNo, rb6RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb6SecurityKeySwitchNo, rb6SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb6GPOInCarNo, rb6GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb6VoiceAnnunciationYes, rb6VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb6PositionIndicatorTypeSurfaceMount, rb6PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb6FacePlateMaterialText, rb6FacvePlateMaterialStainlessSteel, rb6FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb6EmergencyLoweringSystemYes, rb6EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb6OutOfServiceYes, rb6OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 7)
            {
                prefix = "P7";
                #region Page 7 Saving
                SaveTbToXML(tb7AuzCOPLocation, tb7CarDepth, tb7CarDoorFinishText,
                    tb7CarHeight, tb7CarLoad, tb7CarNotes, tb7CarSpeed, tb7CarWidth, tb7CEilingFinishText, tb7ControlerLocationText, tb7COPFinishText,
                    tb7Designations, tb7DoorHeight, tb7DoorTracksText, tb7DoorWidth, tb7FacePlateMaterialText, tb7FloorFinish, tb7FrontWallText,
                    tb7HandrailText, tb7HeadRoom, tb7KeyswitchLocation, tb7LandingDoorFinishText, tb7LiftNumbers, tb7LiftRating,
                    tb7MainCOPLocation, tb7MirrorText, tb7NumberOfCarEntrances, tb7NumberOfCOPs, tb7NumberOfLandingDoors, tb7NumberOfLandings,
                    tb7NumberOfLEDLights, tb7PitDepth, tb7RearWallText, tb7ShaftDepth, tb7ShaftWidth, tb7SideWallText,
                    tb7StructureShaftText, tb7Travel, tb7TypeOfLift
                    );

                SaveRbToXML(rb7AdvancedOpeningNo, rb7AdvancedOpeningYes, rb7BumpRailNo, rb7BumpRailYes, rb7CarDoorFinishOther,
                    rb7CarDoorFinishStainlessSteel, rb7CeilingFinishMirrorStainlessSteel, rb7CeilingFinishOther, rb7CeilingFinishStainlessSteel,
                    rb7CEilingFinishWhite, rb7ControlerLocationBottomLAnding, rb7ControlerLocationOther, rb7ControlerLocationShaft,
                    rb7ControlerLocationTopLanding, rb7COPFinishOther, rb7COPFinishStainlessSteel, rb7DoorNudgingNo, rb7DoorNudgingYes,
                    rb7DoorTracksAluminium, rb7DoorTracksOther, rb7DoorTypeCentreOpening, rb7DoorTypeSideOpening, rb7EmergencyLoweringSystemNo,
                    rb7EmergencyLoweringSystemYes, rb7ExclusiveServiceNo, rb7ExclusiveServiceYes, rb7FacePlateMaterialOther,
                    rb7FacePlateMaterialStainlessSteel, rb7FalseCeilingNo, rb7FalseCeilingYes, rb7FalseFloorNo, rb7FalseFloorYes, rb7FireServiceNo,
                    rb7FireSErviceYes, rb7FrontWallOther, rb7FrontWallStainlessSteel, rb7GPOInCarNo, rb7GPOInCarYes, rb7HandrailOther,
                    rb7HandrailStainlessSteel, rb7IndependentServiceNo, rb7IndpendentServiceYes, rb7LandingDoorFinishOther,
                    rb7LandingDoorFinishStainlessSteel, rb7LCDColourBlue, rb7LCDColourRed, rb7LCDColourWhite, rb7LoadWeighingNo, rb7LoadWeighingYes,
                    rb7MirrorFullSize, rb7MirrorHalfSize, rb7MirrorOther, rb7OutOfSErviceNo, rb7OutOfServiceYes, rb7PositionIndicatorTypeFlushMount,
                    rb7PositionIndicatorTypeSurfaceMount, rb7ProtctiveBlanketsNo, rb7ProtectiveBlanketsYes, rb7RearDoorKeySwitchNo,
                    rb7RearDoorKeySwitchYes, rb7RearWallOther, rb7RearWallStainlessSteel, rb7SecurityKeySwitchNo, rb7SecurityKeySwitchYes,
                    rb7SideWallOther, rb7SideWallStainlessSteel, rb7StructureShaftConcrete, rb7StructureShaftOther,
                     rb7TrimmerBeamsNo, rb7TrimmerBeamsYes, rb7VoiceAnnunciationYes, rb7VoiceAnunciationNo
                    );
                #endregion
                #region Page 7 Word Export

                WordData(prefix + "AE114", tb7LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb7TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb7IndpendentServiceYes, rb7IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb7LoadWeighingYes, rb7LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb7ControlerLocationText, rb7ControlerLocationBottomLAnding, rb7ControlerLocationOther, rb7ControlerLocationShaft, rb7ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb7FireServiceNo, rb7FireSErviceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb7ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb7ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb7PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb7HeadRoom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb7Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb7NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb7NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb7StructureShaftText, rb7StructureShaftConcrete, rb7StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb7TrimmerBeamsYes, rb7TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb7FalseFloorYes, rb7FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb7CarLoad.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb7CarSpeed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb7CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb7CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb7CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb7LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb7LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb7NumberOfCarEntrances.Text);//number of car entraces
                if (tb7CarNotes.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb7CarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb7DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb7DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb7LandingDoorFinishText, rb7LandingDoorFinishOther, rb7LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb7DoorTypeCentreOpening, rb7DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb7DoorTracksText, rb7DoorTracksAluminium, rb7DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb7AdvancedOpeningNo, rb7AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb7DoorNudgingNo, rb7DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb7CarDoorFinishText, rb7CarDoorFinishOther, rb7CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb7CEilingFinishText, rb7CeilingFinishStainlessSteel, rb7CEilingFinishWhite, rb7CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb7FalseCeilingNo, rb7FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb7BumpRailYes, rb7BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb7FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb7FrontWallText, rb7FrontWallStainlessSteel, rb7FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb7MirrorText, rb7MirrorFullSize, rb7MirrorHalfSize, rb7MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb7HandrailText, rb7HandrailStainlessSteel, rb7HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb7SideWallText, rb7SideWallStainlessSteel, rb7SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb7NumberOfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb7ProtectiveBlanketsYes, rb7ProtctiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb7RearWallText, rb7RearWallOther, rb7RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb7NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb7MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb7AuzCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb7Designations.Text); // designations 
                WordData(prefix + "AE191", tb7KeyswitchLocation.Text); //key switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb7COPFinishText, rb7COPFinishStainlessSteel, rb7COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb7LCDColourBlue, rb7LCDColourRed, rb7LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb7ExclusiveServiceNo, rb7ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb7RearDoorKeySwitchNo, rb7RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb7SecurityKeySwitchNo, rb7SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb7GPOInCarNo, rb7GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb7VoiceAnnunciationYes, rb7VoiceAnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb7PositionIndicatorTypeSurfaceMount, rb7PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb7FacePlateMaterialText, rb7FacePlateMaterialStainlessSteel, rb7FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb7EmergencyLoweringSystemYes, rb7EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb7OutOfServiceYes, rb7OutOfSErviceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 8)
            {
                prefix = "P8";
                #region Page 8 Saving
                SaveRbToXML(rb8AdvancedOpeningYes, rb8AdvncedOpeningNo, rb8BumpRailNo, rb8BumpRailYes, rb8CarDoorFinishOther,
                    rb8CarDoorFinishStainlessSteel, rb8CeilingFinishMirrorStainlessSTeel, rb8CeilingFinishOther, rb8CeilingFinishStainlessSteel,
                    rb8CeilingFinishWhite, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther, rb8ControlerLocationShaft,
                    rb8ControlerLocationTopLanding, rb8COPFinishOther, rb8COPFinishStainlessSteel, rb8DoorNudgingYes, rb8DoorTracksAluminium,
                    rb8DoorTracksOther, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening, rb8EmergencyLoweringSystemNo,
                    rb8EmergencyLoweringSystemYes, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes, rb8FacePlateMaterialOther,
                    rb8FacePlateMaterialStainlessSteel, rb8FalseCeilingNo, rb8FalseCeilingYes, rb8FalseFloorNo, rb8FalseFloorYes, rb8FireSErviceNo,
                    rb8FireServiceYes, rb8FrontWallOther, rb8FrontWallStainlessSteel, rb8GPOInCarNo, rb8GPOInCarYes, rb8HandrailOther,
                    rb8HandRailStainlessSteel, rb8IndependentServiceNo, rb8IndependentServiceYes, rb8LandingDoorFinishOther, rb8LCDColourBlue,
                    rb8LCDColourRed, rb8LCDColourWhite, rb8LoadWeighingNo, rb8LoadWeighingYes, rb8MirrorFullSize, rb8MirrorHalfSize,
                    rb8MirrorOther, rb8OutOFServiceNo, rb8OutOfSErviceYes, rb8PositionIndicatorTypeFlushMount, rb8PositionIndicatorTypeSurfaceMount,
                    rb8ProtectiveBlanketsNo, rb8ProtectiveBlanketsYes, rb8RearDoorKeySwitchNo, rb8RearDoorKeySwitchYes, rb8RearWallOther,
                    rb8RearWallStainlessSteel, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes, rb8SideWallOther, rb8SideWallStainlessSteel,
                    rb8StructureShaftConcrete, rb8StructureShaftOther, rb8TimmemrBeamsYes,
                    rb8TrimmerBeamsNo, rb8VoiceAnnunicationNo, rb8VoiceAnnunicationYes, tb8DoorNudgingNo, tb8LandingDoorFinishStainlessSteel
                    );

                SaveTbToXML(tb8AuxCOPLocation, tb8CarDEpth, tb8CarDoorFinishText,
                    tb8CarHeight, tb8CarWidth, tb8CeilingFinishText, tb8ControlerLocationText, tb8COPFinishText, tb8Desiginations, tb8DoorHeight,
                    tb8DoorTracksText, tb8DoorWidth, tb8FacePlateMaterialText, tb8FloorFinish, tb8FrontWallText, tb8HandrailText,
                    tb8Headroom, tb8KeyswitchLocations, tb8LandingDoorFinishText, tb8LiftCarNotes,
                    tb8LiftNumbers, tb8LiftRating, tb8Load, tb8MainCOPLocation, tb8MirrorText, tb8NumberOfCarEntrances, tb8NumberOfCOPs,
                    tb8NumberOfLandingDoors, tb8NumberOfLandings, tb8NumberofLEDLights, tb8PitDepth, tb8RearWallText,
                    tb8ShaftDepth, tb8ShaftWidth, tb8SideWallText, tb8Speed, tb8StructureShaftText, tb8Travel, tb8TypeOfLift
                    );
                #endregion
                #region Page 8 Word Export
                WordData(prefix + "AE114", tb8LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb8TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb8IndependentServiceYes, rb8IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb8LoadWeighingYes, rb8LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb8ControlerLocationText, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther, rb8ControlerLocationShaft, rb8ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb8FireSErviceNo, rb8FireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb8ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb8ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb8PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb8Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb8Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb8NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb8NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb8StructureShaftText, rb8StructureShaftConcrete, rb8StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb8TimmemrBeamsYes, rb8TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb8FalseFloorYes, rb8FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb8Load.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb8Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb8CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb8CarDEpth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb8CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb8LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb8LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb8NumberOfCarEntrances.Text);//number of car entraces
                if (tb8LiftCarNotes.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb8LiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb8DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb8DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb8LandingDoorFinishText, rb8LandingDoorFinishOther, tb8LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb8DoorTracksText, rb8DoorTracksAluminium, rb8DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb8AdvncedOpeningNo, rb8AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb8DoorNudgingYes, tb8DoorNudgingNo));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb8CarDoorFinishText, rb8CarDoorFinishOther, rb8CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb8CeilingFinishText, rb8CeilingFinishStainlessSteel, rb8CeilingFinishWhite, rb8CeilingFinishMirrorStainlessSTeel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb8FalseCeilingNo, rb8FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb8BumpRailYes, rb8BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb8FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb8FrontWallText, rb8FrontWallStainlessSteel, rb8FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb8MirrorText, rb8MirrorFullSize, rb8MirrorHalfSize, rb8MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb8HandrailText, rb8HandRailStainlessSteel, rb8HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb8SideWallText, rb8SideWallStainlessSteel, rb8SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb8NumberofLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb8ProtectiveBlanketsYes, rb8ProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb8RearWallText, rb8RearWallOther, rb8RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb8NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb8MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb8AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb8Desiginations.Text); // designations 
                WordData(prefix + "AE191", tb8KeyswitchLocations.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb8COPFinishText, rb8COPFinishStainlessSteel, rb8COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb8LCDColourBlue, rb8LCDColourRed, rb8LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb8RearDoorKeySwitchNo, rb8RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb8GPOInCarNo, rb8GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb8VoiceAnnunicationYes, rb8VoiceAnnunicationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb8PositionIndicatorTypeSurfaceMount, rb8PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb8FacePlateMaterialText, rb8FacePlateMaterialStainlessSteel, rb8FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb8EmergencyLoweringSystemYes, rb8EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb8OutOfSErviceYes, rb8OutOFServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 9)
            {
                prefix = "P9";
                #region Page 9 Saving
                SaveRbToXML(rb9AdvancedOpeningNo, rb9AdvancedOpeningYes, rb9BumpRailNo, rb9BumpRailYes, rb9CarDoorFinishOther,
                    rb9CarDoorFinishStainlessSteel, rb9CeilingFinishMirrorStainlessSteel, rb9CeilingFinishOther, rb9CeilingFinishStainlessSteel,
                    rb9CeilingFinishWhite, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther, rb9ControlerLocationShaft, rb9ControlrLocationTopLanding,
                    rb9COPFinishOther, rb9COPFinishStainlessSteel, rb9DoorNudgingNo, rb9DoorNudgingYes, rb9DoorTracksAluminium, rb9DoorTracksOther,
                    rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening, rb9EmergencyLoweringSystemNo, rb9EmergencyLoweringSystemYes, rb9ExclusiveServiceNo,
                    rb9ExclusiveServiceYes, rb9FacePlateMaterialOther, rb9FacePlateMaterialStainlessSteel, rb9FalseCeilingNo, rb9FalseCeilingYes, rb9FalseFloorNo,
                    rb9FalseFloorYes, rb9FireServiceNo, rb9FireSErviceYes, rb9FrontWallOther, rb9FrontWallStainlessSteel, rb9GPOInCarNo, rb9GPOInCarYes,
                    rb9HandrailOther, rb9HandrailStainlessSteel, rb9IndependentServiceNo, rb9IndependentServiceYes, rb9LandingDoorFinishOther,
                    rb9LandingDoorFinishStainlessSteel, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite, rb9LoadWeighingNo, rb9LoadWeighingYes,
                    rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther, rb9OutOfServiceNo, rb9OutOfServiceYes, rb9PositionIndicatorTypeFlushMount,
                    rb9PositionIndicatorTypeSurfaceMount, rb9ProtectiveBlanketsNo, rb9ProtectiveBlanketsYes, rb9RearDoorKeySwitchNo, rb9RearDoorKeySwitchYes,
                    rb9RearWallOther, rb9RearWallStainlessSteel, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes, rb9SideWallOther, rb9SideWallStainlessSteel,
                    rb9StructureShaftConcrete, rb9StructureShaftOther, rb9TrimmerBeamsNo, rb9TrimmerBeamsYes,
                    rb9VoiceAnnunciationNo, rb9VoiceAnnunciationYes
                    );

                SaveTbToXML(tb9AuxCOPLocation, tb9CarDepth, tb9CarDoorFinishText, tb9CarHeight,
                    tb9CarNotes, tb9CarWidth, tb9CeilingFinishText, tb9ControlerLocationText, tb9COPFinishText, tb9Designations, tb9DoorHeight, tb9DoorTracksText,
                    tb9DoorWidth, tb9FacePlateMaterialText, tb9FloorFinish, tb9FrontWallText, tb9HandrailTexrt, tb9Headroom, tb9KeyswitchLocation,
                    tb9LandingDoorFinishText, tb9LiftNumbers, tb9LiftRating, tb9Load, tb9MainCOPLocation, tb9MirrorText, tb9NumberOFCarEntraces,
                    tb9NumberOfCOPs, tb9NumberOfLandingDoors, tb9NumberOfLandings, tb9NumberOfLEDLights, tb9PitDepth, tb9RearWallText,
                    tb9ShaftDepth, tb9ShaftWidth, tb9SideWallText, tb9Speed, tb9StructureShaftText, tb9Travel, tb9TypeOfLift
                    );
                #endregion
                #region Page 9 Word Export
                WordData(prefix + "AE114", tb9LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb9TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb9IndependentServiceYes, rb9IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb9LoadWeighingYes, rb9LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb9ControlerLocationText, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther, rb9ControlerLocationShaft, rb9ControlrLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb9FireServiceNo, rb9FireSErviceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb9ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb9ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb9PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb9Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb9Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb9NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb9NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb9StructureShaftText, rb9StructureShaftConcrete, rb9StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb9TrimmerBeamsYes, rb9TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb9FalseFloorYes, rb9FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb9Load.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb9Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb9CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb9CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb9CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb9LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb9LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb9NumberOFCarEntraces.Text);//number of car entraces
                if (tb9CarNotes.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb9CarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb9DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb9DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb9LandingDoorFinishText, rb9LandingDoorFinishOther, rb9LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb9DoorTracksText, rb9DoorTracksAluminium, rb9DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb9AdvancedOpeningNo, rb9AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb9DoorNudgingNo, rb9DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb9CarDoorFinishText, rb9CarDoorFinishOther, rb9CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb9CeilingFinishText, rb9CeilingFinishStainlessSteel, rb9CeilingFinishWhite, rb9CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb9FalseCeilingNo, rb9FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb9BumpRailYes, rb9BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb9FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb9FrontWallText, rb9FrontWallStainlessSteel, rb9FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb9MirrorText, rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb9HandrailTexrt, rb9HandrailStainlessSteel, rb9HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb9SideWallText, rb9SideWallStainlessSteel, rb9SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb9NumberOfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb9ProtectiveBlanketsYes, rb9ProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb9RearWallText, rb9RearWallOther, rb9RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb9NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb9MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb9AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb9Designations.Text); // designations 
                WordData(prefix + "AE191", tb9KeyswitchLocation.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb9COPFinishText, rb9COPFinishStainlessSteel, rb9COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb9ExclusiveServiceNo, rb9ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb9RearDoorKeySwitchNo, rb9RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb9GPOInCarNo, rb9GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb9VoiceAnnunciationYes, rb9VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb9PositionIndicatorTypeSurfaceMount, rb9PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb9FacePlateMaterialText, rb9FacePlateMaterialStainlessSteel, rb9FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb9EmergencyLoweringSystemYes, rb9EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb9OutOfServiceYes, rb9OutOfServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 10)
            {
                prefix = "P10";
                #region Page 10 Saving
                SaveTbToXML(tb10AuxCOPLocation, tb10CarDepth, tb10CarDoorFinishText,
                    tb10CarHeight, tb10CarWidth, tb10CEilingFinishText, tb10ControlerLocationText, tb10COPFinishText, tb10Desigination, tb10DoorHeight,
                    tb10DoorTracksText, tb10DoorWidth, tb10FacePlateMaterialText, tb10FloorFinish, tb10FrontWallText, tb10HandrailText,
                    tb10Headroom, tb10KeyswitchLocation, tb10LandingDoorFinishText, tb10LiftCarLoad, tb10LiftCarNotes, tb10LiftNumbers,
                    tb10LiftRating, tb10MainCOPLocation, tb10MirrorText, tb10NumberofCarEntrances, tb10NumberOfCOPs, tb10NumberofLandingDoors,
                    tb10NumberofLandings, tb10NumberOfLEDLIghts, tb10PitDepth, tb10RearWallText, tb10ShaftDepth, tb10ShaftWidth,
                    tb10SideWallText, tb10Speed, tb10StructureShaftText, tb10Travel, tb10TypeOfLift
                    );

                SaveRbToXML(rb10AdvancedOpeningNo, rb10AdvancedOpeningYes, rb10BumpRaidYes, rb10BumpRailNo, rb10CarDoorFinishOther,
                    rb10CarDoorFinishStainlessSteel, rb10CEilingFinishMirrorStainlessSteel, rb10CEilingFinishOther, rb10CeilingFinishStainlessSteel,
                    rb10CeilingFinishWhite, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther, rb10ControlerLocationShaft,
                    rb10ControlerLocationTopLanding, rb10COPFinishOther, rb10COPFinishStainlessSteel, rb10DoorNudgingNo, rb10DoorNudgingYes,
                    rb10DoorTracksAluminium, rb10DoorTracksOther, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening, rb10EmergencyLoweringSystemNo,
                    rb10EmergencyLoweringSystemYes, rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes, rb10FacePlateMaterialOther, rb10FacePlateMaterialStainlessSteel,
                    rb10FalseCeilingNo, rb10FalseCeilingYes, rb10FalseFloorNo, rb10FalseFloorYes, rb10FireSERviceNo, rb10FireSErviceYes, rb10FrontWallOther,
                    rb10FrontWallStainlessSteel, rb10GPOInCarNo, rb10GPOInCarYes, rb10HandrailOther, rb10HandrailStainlessSteel, rb10IndependentServiceNo,
                    rb10IndependentServiceYes, rb10LAndingDoorFinishOtherr, rb10LandingDoorFinishStainlessSteel, rb10LCDColourBlue, rb10LCDColourRed,
                    rb10LCDColourWhite, rb10LoadWEighingNo, rb10LoadWeighingYes, rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther, rb10OutOfServiceNo,
                    rb10OutOFServiceYes, rb10PositionIndicatorTypeFlushMount, rb10PositionIndicatorTypeSurfaceMount, rb10ProtectiveBlanketNo, rb10ProtectiveBlanketYes,
                    rb10RearDoorKeySwitchNo, rb10RearDoorKeySwitchYes, rb10RearWallOther, rb10RearWallStainlessSteel, rb10SecurityKeySwitchNo,
                    rb10SecurityKeySwitchYes, rb10SideWallOther, rb10SideWallStainlesSteel, rb10StructureShaftConcrete, rb10StructureShaftOther,
                      rb10TimmerbeamsYes, rb10TrimmerBeamsNo, rb10VoiceAnnunciationNo, rb10VoiceAnnunciationYes
                    );
                #endregion
                #region Page 10 Word Export
                WordData(prefix + "AE114", tb10LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb10TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb10IndependentServiceYes, rb10IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb10LoadWeighingYes, rb10LoadWEighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb10ControlerLocationText, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther, rb10ControlerLocationShaft, rb10ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb10FireSERviceNo, rb10FireSErviceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb10ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb10ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb10PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb10Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb10Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb10NumberofLandings.Text);// number of landings
                WordData(prefix + "AE129", tb10NumberofLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb10StructureShaftText, rb10StructureShaftConcrete, rb10StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb10TimmerbeamsYes, rb10TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb10FalseFloorYes, rb10FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb10LiftCarLoad.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb10Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb10CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb10CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb10CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb10LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb10LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb10NumberofCarEntrances.Text);//number of car entraces
                if (tb10LiftCarNotes.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb10LiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb10DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb10DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb10LandingDoorFinishText, rb10LAndingDoorFinishOtherr, rb10LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb10DoorTracksText, rb10DoorTracksAluminium, rb10DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb10AdvancedOpeningNo, rb10AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb10DoorNudgingNo, rb10DoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb10CarDoorFinishText, rb10CarDoorFinishOther, rb10CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb10CEilingFinishText, rb10CeilingFinishStainlessSteel, rb10CeilingFinishWhite, rb10CEilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb10FalseCeilingNo, rb10FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb10BumpRaidYes, rb10BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb10FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb10FrontWallText, rb10FrontWallStainlessSteel, rb10FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb10MirrorText, rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb10HandrailText, rb10HandrailStainlessSteel, rb10HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb10SideWallText, rb10SideWallStainlesSteel, rb10SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb10NumberOfLEDLIghts.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb10ProtectiveBlanketYes, rb10ProtectiveBlanketNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb10RearWallText, rb10RearWallOther, rb10RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb10NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb10MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb10AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb10Desigination.Text); // designations 
                WordData(prefix + "AE191", tb10KeyswitchLocation.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb10COPFinishText, rb10COPFinishStainlessSteel, rb10COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb10LCDColourBlue, rb10LCDColourRed, rb10LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb10RearDoorKeySwitchNo, rb10RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb10SecurityKeySwitchNo, rb10SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb10GPOInCarNo, rb10GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb10VoiceAnnunciationYes, rb10VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb10PositionIndicatorTypeSurfaceMount, rb10PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb10FacePlateMaterialText, rb10FacePlateMaterialStainlessSteel, rb10FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb10EmergencyLoweringSystemYes, rb10EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb10OutOFServiceYes, rb10OutOfServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 11)
            {
                prefix = "P11";
                #region Page 11 Saving
                SaveTbToXML(tb11AuxCOPLocation, tb11CarDepth, tb11CarDoorFinishText,
                    tb11CarHeight, tb11CarWidth, tb11CeilingFinishText, tb11ControlerLocationText, tb11COPFinishText, tb11Designations, tb11DoorHeight,
                    tb11DoorTracksText, tb11DoorWidth, tb11FaceplateMaterialText, tb11FloorFinish, tb11FrontWallText, tb11HandrailText, tb11Headroom,
                    tb11KeyswitchLocation, tb11LandingDoorFInishOther, tb11LiftCarLoad, tb11LiftCarNote, tb11LiftRating, tb11MainCOPLocation,
                    tb11MirrorText, tb11NumberofCarEntrances, tb11NumberOfCOPs, tb11NumberOfLandingDoors, tb11NumberOfLandings, tb11NumberOfLEDLights,
                     tb11PitDepth, tb11RearWallText, tb11ShaftDepth, tb11ShaftWidth, tb11SideWallText, tb11Speed, tb11Travel, tb11TypeOfLift,
                    rb11LiftNumbers, rb11StructureShaftText
                    );

                SaveRbToXML(rb11AdvancedOpeningNo, rb11AdvancedOpeningYes, rb11BumpRailNo, tb11ControlerLocationTopLanding, rb11BumpRailYes,
                    rb11CarDoorFinishOther, rb11CarDoorFinishStainlessSteel, rb11CeilingFinishMirrorStainlessSteel, rb11CeilingFinishOther, rb11CeilingFinishStainlessSteel,
                    rb11CeilingFinishWhite, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther, rb11ControlerLocationShaft, rb11COPFinishOther,
                    rb11COPFinishStainlessSteel, rb11DoorNudgingNo, rb11DoorNudgingYes, rb11DoorTracksAluminium, rb11DoorTracksOther, rb11DoorTypeCentreOpening,
                    rb11DoorTypeSideOpening, rb11EmergencyLoweringSystemNo, rb11EmergencyLoweringSystemYes, rb11ExclusiveServiceNo, rb11ExclusiveServiceYes,
                    rb11FacePlateMaterialOther, rb11FacePlateMaterialStainlessSTeel, rb11FalseCeilingNo, rb11FalseCEilingYes, rb11FalseFloorNo, rb11FalseFloorYes,
                    rb11FireServiceNo, rb11FireServiceYes, rb11FrontWallOther, rb11FrontWallStainlessSteel, rb11GPOInCarNo, rb11GPOInCarYes, rb11HandrailOther,
                    rb11HandrailStainlessSteel, rb11IndependentSErviceNO, rb11IndependentServiceYes, rb11LandingDoorFinishOther, rb11LandingDoorFinishStainlessSteel,
                    rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite, rb11LoadWeighingNo, rb11LoadWeighingYes, rb11MirrorFullSize,
                    rb11MirrorHalfSize, rb11MirrorOther, rb11OutOfServiceNo, rb11OutOfSErviceYes, rb11PositionIndicatorTypeFlushMount,
                    rb11PositionIndicatorTypeSurfaceMount, rb11ProtectiveBlanketNo, rb11ProtectiveBlanketsYes, rb11RearDoorKeySwitchNo, rb11RearDoorKeySwitchYes,
                    rb11RearWallOther, rb11RearWallStainlessSteel, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes, rb11SideWallOther, rb11SideWallStainlessSteel,
                    rb11StrructureShaftOther, rb11StructureShaftConcrete,
                    rb11TrimmerBeamNo, rb11TrimmerBeamsYes, rb11VoiceAnnunciationNo, rb11VoiceAnnunciationYes
                    );
                #endregion
                #region Page 11 Word Export
                WordData(prefix + "AE114", rb11LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb11TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb11IndependentServiceYes, rb11IndependentSErviceNO));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb11LoadWeighingYes, rb11LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb11ControlerLocationText, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther, rb11ControlerLocationShaft, tb11ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb11FireServiceNo, rb11FireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb11ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb11ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb11PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb11Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb11Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb11NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb11NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(rb11StructureShaftText, rb11StructureShaftConcrete, rb11StrructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb11TrimmerBeamsYes, rb11TrimmerBeamNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb11FalseFloorYes, rb11FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb11LiftCarLoad.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb11Speed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb11CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb11CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb11CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb11LiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb11LiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb11NumberofCarEntrances.Text);//number of car entraces
                if (tb11LiftCarNote.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb11LiftCarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb11DoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb11DoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb11LandingDoorFInishOther, rb11LandingDoorFinishOther, rb11LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb11DoorTypeCentreOpening, rb11DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb11DoorTracksText, rb11DoorTracksAluminium, rb11DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb11AdvancedOpeningNo, rb11AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb11DoorNudgingYes, rb11DoorNudgingNo));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb11CarDoorFinishText, rb11CarDoorFinishOther, rb11CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb11CeilingFinishText, rb11CeilingFinishStainlessSteel, rb11CeilingFinishWhite, rb11CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb11FalseCeilingNo, rb11FalseCEilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb11BumpRailYes, rb11BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb11FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb11FrontWallText, rb11FrontWallStainlessSteel, rb11FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb11MirrorText, rb11MirrorFullSize, rb11MirrorHalfSize, rb11MirrorOther));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb11HandrailText, rb11MirrorHalfSize, rb11HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb11SideWallText, rb11SideWallStainlessSteel, rb11SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb11NumberOfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb11ProtectiveBlanketsYes, rb11ProtectiveBlanketNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb11RearWallText, rb11RearWallOther, rb11RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb11NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb11MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb11AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb11Designations.Text); // designations 
                WordData(prefix + "AE191", tb11KeyswitchLocation.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb11COPFinishText, rb11COPFinishStainlessSteel, rb11COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb11ExclusiveServiceNo, rb11ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb11RearDoorKeySwitchNo, rb11RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb11GPOInCarNo, rb11GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb11VoiceAnnunciationYes, rb11VoiceAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb11PositionIndicatorTypeSurfaceMount, rb11PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb11FaceplateMaterialText, rb11FacePlateMaterialStainlessSTeel, rb11FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb11EmergencyLoweringSystemYes, rb11EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb11OutOfSErviceYes, rb11OutOfServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 12)
            {
                prefix = "P12";
                #region Page 12 Saving
                SaveTbToXML(tb12AuxCOPLocation, tb12CarDepth, tb12CarDoorFinishText,
                    tb12CarHeight, tb12CarLiftRating, tb12CarLoad, tb12CarNumberOfCarEntrances, tb12CarSpeed, tb12CarWidth, tb12CeilingFinishText,
                    tb12ControlerLocationText, tb12COPFinishText, tb12Designations, tb12DoorTracksText, tb12FacePlateMaterialText,
                    tb12FloorFinish, tb12FrontWallText, tb12HandrailText, tb12Headroom, tb12KeyswitchLocation, tb12LandingDoorFinishText,
                    tb12LandingDoorHeight, tb12LandingDoorWidth, tb12LiftCarNotes, tb12LiftNumbers, tb12MainCOPLocation, tb12MirrorText,
                    tb12NumberOfCOPs, tb12NumberOfLandingDoors, tb12NumberOfLandings, tb12NumberOfLEDLights, tb12PitDepth,
                     tb12RearWallText, tb12ShaftDepth, tb12ShaftWidth, tb12SideWallText, tb12StructureShaftText, tb12Travel, tb12TypeOfLift
                     );

                SaveRbToXML(rb12AdvancedOpeningNo, rb12AdvancedOpeningYes, rb12BumpRailNo, rb12BumpRailYes, rb12CarDoorFinishOther,
                    rb12CarDoorFinishStainlessSteel, rb12CeilingFinishMirrorStainlessSteel, rb12CeilingFinishOther, rb12CeilingFinishStainlessSteel,
                    rb12CeilingFinishWhite, rb12ControlerLocationBottomLanding, rb12ControlerLocationOther, rb12ControlerLocationShaft,
                    rb12ControlerLocationTopLanding, rb12COPFinishOther, rb12COPFinishStainlessSteel, rb12DoorTracksAluminium, rb12DoorTracksOther,
                    rb12DoorTypeCentreOpening, rb12DoorTypeSideOpening, rb12EmergencyLoweringSystemNo, rb12EmergencyLoweringSystemYes,
                    rb12ExclusiveServiceNo, rb12ExclusiveServiceYes, rb12FacePlateMaterialOther, rb12FacePlateMaterialStainlessSteel, rb12FalseCeilingNo,
                    rb12FalseCeilingYes, rb12FalseFloorNo, rb12FalseFloorYes, rb12FireServiceNo, rb12FireServiceYes, rb12FrontWallOther,
                    rb12FrontWallStainlessSteel, rb12GPOInCarNo, rb12GPOInCarYes, rb12HandrailOther, rb12HandrailStainlessSTeel, rb12IndependentServiceNo,
                    rb12IndependentServiceYes, rb12LandingDoorFinishOther, rb12LandingDoorFinishStainlessSteel, rb12LandingDoorNudgingNo,
                    rb12LandingDoorNudgingYes, rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite, rb12LoadWeighingNo, rb12LoadWeighingYes,
                    rb12MirrorFullSize, rb12MirrorHalfSize, rb12MirrorOTher, rb12OutOfServiceNo, rb12OutOfServiceYes, rb12PositionIndicatorTypeFlushMount,
                    rb12PositionIndicatorTypeSurfaceMount, rb12ProectiveBlanketsYes, rb12ProtectiveBlanketsNo, rb12RearDoorKeySwitchNo, rb12RearDoorKeySwitchYes,
                    rb12RearWallOther, rb12RearWallStainlessSteel, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes, rb12SideWallOther, rb12SideWallStainlessSteel,
                    rb12StructureShaftConcrete, rb12StructureShaftOther, rb12TrimmerBeamsNo,
                    rb12TrimmerBeamsYes, rb12VoicAnnunciationNo, rb12VoiceAnnuniationYes
                    );
                #endregion
                #region Page 12 Word Export
                WordData(prefix + "AE114", tb12LiftNumbers.Text);//lift number
                WordData(prefix + "AE115", tb12TypeOfLift.Text);//type of lift
                WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                WordData(prefix + "AE116", RadioButtonToAsteriskHandeler(rb12IndependentServiceYes, rb12IndependentServiceNo));//independent service
                WordData(prefix + "AE117", RadioButtonToAsteriskHandeler(rb12LoadWeighingYes, rb12LoadWeighingNo));//load weighing
                WordData(prefix + "AE118", RadioButtonHandeler(tb12ControlerLocationText, rb12ControlerLocationBottomLanding, rb12ControlerLocationOther, rb12ControlerLocationShaft, rb12ControlerLocationTopLanding));//controler location
                WordData(prefix + "AE120", RadioButtonHandeler(null, rb12FireServiceNo, rb12FireServiceYes));//fire service
                WordData(prefix + "AE123", MeasureStringChecker(tb12ShaftWidth.Text, "mm"));// shaft width
                WordData(prefix + "AE124", MeasureStringChecker(tb12ShaftDepth.Text, "mm"));//shaft depth
                WordData(prefix + "AE125", MeasureStringChecker(tb12PitDepth.Text, "mm"));//pit depth
                WordData(prefix + "AE126", MeasureStringChecker(tb12Headroom.Text, "mm"));//headroom
                WordData(prefix + "AE127", MeasureStringChecker(tb12Travel.Text, "mm"));//travel
                WordData(prefix + "AE128", tb12NumberOfLandings.Text);// number of landings
                WordData(prefix + "AE129", tb12NumberOfLandingDoors.Text);//number of landing doors 
                WordData(prefix + "AE130", RadioButtonHandeler(tb12StructureShaftText, rb12StructureShaftConcrete, rb12StructureShaftOther)); //structure shaft 
                WordData(prefix + "AE132", RadioButtonToAsteriskHandeler(rb12TrimmerBeamsYes, rb12TrimmerBeamsNo));//trimmer beams
                WordData(prefix + "AE133", RadioButtonToAsteriskHandeler(rb12FalseFloorYes, rb12FalseFloorNo));//false floor
                WordData(prefix + "AE134", MeasureStringChecker(tb12CarLoad.Text, "kg")); //load
                WordData(prefix + "AE135", MeasureStringChecker(tb12CarSpeed.Text, "mps"));//speed
                WordData(prefix + "AE136", MeasureStringChecker(tb12CarWidth.Text, "mm")); // width
                WordData(prefix + "AE137", MeasureStringChecker(tb12CarDepth.Text, "mm"));//depth
                WordData(prefix + "AE138", MeasureStringChecker(tb12CarHeight.Text, "mm"));//height
                WordData(prefix + "AE139", MeasureStringChecker(tb12CarLiftRating.Text, "passengers"));//classification rating
                WordData(prefix + "AE113", MeasureStringChecker(tb12CarLiftRating.Text, "passenger"));//classification rating
                WordData(prefix + "AE140", tb12CarNumberOfCarEntrances.Text);//number of car entraces
                if (tb12LiftCarNotes.Text != "")
                {
                    WordData(prefix + "AE142", "NOTE: " + tb12LiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    WordData(prefix + "AE142", "");//notes
                }
                WordData(prefix + "AE143", MeasureStringChecker(tb12LandingDoorWidth.Text, "mm"));//door width 
                WordData(prefix + "AE144", MeasureStringChecker(tb12LandingDoorHeight.Text, "mm")); //door height 
                WordData(prefix + "AE147", RadioButtonHandeler(tb12LandingDoorFinishText, rb12LandingDoorFinishOther, rb12LandingDoorFinishStainlessSteel));//landing door finish
                WordData(prefix + "AE148", RadioButtonHandeler(null, rb12DoorTypeCentreOpening, rb12DoorTypeSideOpening));//door type
                WordData(prefix + "AE150", RadioButtonHandeler(tb12DoorTracksText, rb12DoorTracksAluminium, rb12DoorTracksOther));// door tracks 
                WordData(prefix + "AE151", RadioButtonHandeler(null, rb12AdvancedOpeningNo, rb12AdvancedOpeningYes));//advanced opening
                WordData(prefix + "AE152", RadioButtonHandeler(null, rb12LandingDoorNudgingNo, rb12LandingDoorNudgingYes));//door nudging 
                WordData(prefix + "AE155", RadioButtonHandeler(tb12CarDoorFinishText, rb12CarDoorFinishOther, rb12CarDoorFinishStainlessSteel));//car door finish
                WordData(prefix + "AE156", RadioButtonHandeler(tb12CeilingFinishText, rb12CeilingFinishStainlessSteel, rb12CeilingFinishWhite, rb12CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                WordData(prefix + "AE157", RadioButtonHandeler(null, rb12FalseCeilingNo, rb12FalseCeilingYes));//false ceiling
                WordData(prefix + "AE158", RadioButtonHandeler(null, rb12BumpRailYes, rb12BumpRailNo));//bump rail
                WordData(prefix + "AE159", tb12FloorFinish.Text);//floor
                WordData(prefix + "AE160", RadioButtonHandeler(tb12FrontWallText, rb12FrontWallStainlessSteel, rb12FrontWallOther));//front wall
                WordData(prefix + "AE161", RadioButtonHandeler(tb12MirrorText, rb12MirrorFullSize, rb12MirrorHalfSize, rb12MirrorOTher));//mirror
                WordData(prefix + "AE162", RadioButtonHandeler(tb12HandrailText, rb12HandrailStainlessSTeel, rb12HandrailOther));//handrail
                WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                WordData(prefix + "AE164", RadioButtonHandeler(tb12SideWallText, rb12SideWallStainlessSteel, rb12SideWallOther)); //side wall 
                WordData(prefix + "AE165", tb12NumberOfLEDLights.Text + " LED Lights"); // lighting 
                WordData(prefix + "AE167", RadioButtonToAsteriskHandeler(rb12ProectiveBlanketsYes, rb12ProtectiveBlanketsNo)); // protective blankets 
                WordData(prefix + "AE216", RadioButtonHandeler(tb12RearWallText, rb12RearWallOther, rb12RearWallStainlessSteel)); //  rear wall
                WordData(prefix + "AE168", tb12NumberOfCOPs.Text); // number of COPS
                WordData(prefix + "AE169", tb12MainCOPLocation.Text);// main COP location
                WordData(prefix + "AE170", tb12AuxCOPLocation.Text);//aux cop location
                WordData(prefix + "AE171", tb12Designations.Text); // designations 
                WordData(prefix + "AE191", tb12KeyswitchLocation.Text); //keyt switch location
                WordData(prefix + "AE172", RadioButtonHandeler(tb12COPFinishText, rb12COPFinishStainlessSteel, rb12COPFinishOther));// COP finish
                WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                WordData(prefix + "AE174", RadioButtonHandeler(null, rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite));// LCD colour
                WordData(prefix + "AE183", RadioButtonHandeler(null, rb12ExclusiveServiceNo, rb12ExclusiveServiceYes));//exclusive service 
                WordData(prefix + "AE184", RadioButtonHandeler(null, rb12RearDoorKeySwitchNo, rb12RearDoorKeySwitchYes));// rear door kew switch 
                WordData(prefix + "AE186", RadioButtonHandeler(null, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes));//security key switch 
                WordData(prefix + "AE187", RadioButtonHandeler(null, rb12GPOInCarNo, rb12GPOInCarYes));//GPO in car
                WordData(prefix + "AE189", RadioButtonToAsteriskHandeler(rb12VoiceAnnuniationYes, rb12VoicAnnunciationNo));//voice annunciation 
                WordData(prefix + "AE190", RadioButtonHandeler(null, rb12PositionIndicatorTypeSurfaceMount, rb12PositionIndicatorTypeFlushMount));// position indicaor type 
                WordData(prefix + "AE192", RadioButtonHandeler(tb12FacePlateMaterialText, rb12FacePlateMaterialStainlessSteel, rb12FacePlateMaterialOther));//face plate material 
                WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                WordData(prefix + "AE209", RadioButtonHandeler(null, rb12EmergencyLoweringSystemYes, rb12EmergencyLoweringSystemNo));// emergency lowering system 
                WordData(prefix + "AE210", RadioButtonToAsteriskHandeler(rb12OutOfServiceYes, rb12OutOfServiceNo));//out of service 
                #endregion
            }

            this.Enabled = false;
            QuestionsComplete();
            this.Close();
        }
        #endregion

        #region Swapping Between Calculator and Exporter
        // click generate customer quote button
        private void button4_Click_2(object sender, EventArgs e)
        {
            SwapBetweenCalcAndExp(true);
        }

        // click edit quote prices button
        private void btnEditQuotePrices_Click(object sender, EventArgs e)
        {
            SwapBetweenCalcAndExp(false);
        }

        private void SwapBetweenCalcAndExp(bool swapToExporter)
        {
            for (int i = 0; i < numberOfPagesNeeded; i++)
            {
                NewPage();
                numberOfPagesUsed++;
            }

            PanelMenuChange(null);
            SetValuesForMainMenuLabels();
            //calc panels
            PanelDefaultSettings(panelAddress, panelAddress.Location, !swapToExporter, !swapToExporter);
            PanelDefaultSettings(panelCalcButtons, panelCalcButtons.Location, !swapToExporter, !swapToExporter);
            PanelDefaultSettings(panelCostBreakdown, panelCostBreakdown.Location, !swapToExporter, !swapToExporter);
            //exp panels
            if (numberOfPagesNeeded >= 1)
            {
                PanelDefaultSettings(panelPageNumberButtons, panelPageNumberButtons.Location, swapToExporter, swapToExporter);
            }
            PanelDefaultSettings(panelExportQuote, panelExportQuote.Location, swapToExporter, swapToExporter);
            PanelDefaultSettings(panelLift1, panelLift1.Location, swapToExporter, swapToExporter);
            PanelDefaultSettings(panelContactDetails, panelContactDetails.Location, swapToExporter, swapToExporter);
            //swap to exp button
            btnGenerateCustomerQuote.Enabled = !swapToExporter;
            btnGenerateCustomerQuote.Visible = !swapToExporter;
            //swap to calc button
            btnEditQuotePrices.Enabled = swapToExporter;
            btnEditQuotePrices.Visible = swapToExporter;
        }

        #endregion

        #region Price Rounding Methods

        private float Roundingbuffer(float unroundedPrice)
        {
            float priceWithDecimal = unroundedPrice;
            float roundingPriceBuffer = 0;

            if (!cbAutoRounding.Checked)
            {
                try
                {
                    roundingPriceBuffer = float.Parse(tbMinorPriceAdjustment.Text);
                }
                catch (Exception)
                {
                    tbMinorPriceAdjustment.Text = "0";
                    roundingPriceBuffer = 0;
                }
            }
            else if (cbAutoRounding.Checked)
            {
                float f = float.Parse(priceWithDecimal.ToString("0.00").Substring(priceWithDecimal.ToString("0.00").Length - Math.Min(6, priceWithDecimal.ToString("0.00").Length)));

                if (f < 100)
                {
                    roundingPriceBuffer = f * -1;
                }
                else if (f < 600)
                {
                    roundingPriceBuffer = (f - 500) * -1;
                }
                else if (f > 100)
                {
                    float x = 500 - f;
                    float y = 1000 - f;
                    if (x > y || x < 0)
                    {
                        roundingPriceBuffer = y;
                    }
                    else
                    {
                        roundingPriceBuffer = x;
                    }
                }
            }
            return roundingPriceBuffer;
        }

        private void cbAutoRounding_CheckedChanged(object sender, EventArgs e)
        {
            if (cbAutoRounding.Checked)
            {
                tbMinorPriceAdjustment.Enabled = false;
            }
            else
            {
                tbMinorPriceAdjustment.Enabled = true;
            }
        }
        #endregion

        #region Salesman Signature Placer
        /* Unneeded code as have persued another method
         * 
        private string[] DocumentSignaturePath()
        {
            string userName = Environment.UserName;
            string[] returnValues = { "", "", "" }; // 0 =signature address, 1 = salesman name, 2= salesman title

            switch (userName)
            {
                // need to add other salesman to this section to have their signatures added to documents
                case "myles":
                    returnValues[0] = "";
                    returnValues[1] = "Myles Okorn";
                    returnValues[2] = "IT";

                    break;

                default:
                    break;
            }

            return returnValues;
        }
        */
        #endregion

        private void tBMainQuoteNumber_TextChanged(object sender, EventArgs e)
        {
            this.Text = (tBMainQuoteNumber.Text + " Calculation Window");

        }
    }
}