using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace SQT
{
    public partial class Form1 : Form
    {
        // VARS
        public string quoteNumber = "";
        public float applicableExchangeRate = 1;
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();
        public string exchangeRateDate;
        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public int num20Ft;
        public int num40Ft;
        public float freightTotal = 0;
        public float liftPrice;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // currencySelectionGroup.Visible = false;
            FetchBasePrices();
            FetchCurrencyRates();
            FetchLabourPrices();
            GeneratePriceList();
            lblCostOfParts.Text = "$0";
            lblCostIncludingMargin.Text = "$0";
            lblGST.Text = "$0";
            lblPriceIncludingGST.Text = "$0";
            //tBAddress.Text = "$" + exchangeRates["USD"] + "  $" + exchangeRates["EUR"].ToString();
            quoteNumber = ("Qu" + DateTime.Now.ToString("yy") + "-000");
            tBQuoteNumber.Text = quoteNumber;
        }

        private void tBQuoteNumber_TextChanged(object sender, EventArgs e)
        {
            //when the quote number text box is changed, if it is different to the existing quote number update the quote number
            if (quoteNumber != tBQuoteNumber.Text)
            {
                quoteNumber = tBQuoteNumber.Text;
            }
            //MessageBox.Show(quoteNumber.ToString());
        }

        //find the live currency rates from floatrates.com and write them into the Exchange rate dictionary 
        private void FetchCurrencyRates()
        {
            //bool grabRate = false;
            //bool switchEUR = false;
            string dKey = "";
            float dName = 0;
            bool b = true;
            //XmlDocument currencyXML = new XmlDocument();
            //currencyXML.Load("http://www.floatrates.com/daily/aud.xml");

            XmlTextReader XMLR = new XmlTextReader("http://www.floatrates.com/daily/aud.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "targetCurrency")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }

                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "inverseRate")
                {
                    dName = float.Parse(XMLR.ReadElementContentAsString());
                }

                if (dKey != "" && dName != 0)
                {
                    exchangeRates.Add(dKey, dName * basePrices["16CurrencyMargin"]);
                    dKey = "";
                    dName = 0;
                }

                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "pubDate" && b)
                {
                    exchangeRateDate = XMLR.ReadElementContentAsString();
                    b = false;
                }
            }
            //MessageBox.Show("EXCHANGE");
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

                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "price")
                {
                    dName = float.Parse(XMLR.ReadElementContentAsString());
                }

                if (dKey != "" && dName != -1)
                {
                    //MessageBox.Show(dKey + dName);
                    basePrices.Add(dKey, dName);
                    //MessageBox.Show(basePrices[dKey].ToString());
                    dKey = "";
                    dName = -1;
                }
            }
            XMLR.Close();
            //MessageBox.Show("BASE");
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

                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Price")
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
            //MessageBox.Show("BASE");
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
            }
            else if (selector == "U")
            {
                applicableExchangeRate = exchangeRates["USD"];
                exchangeRateLbl.Visible = true;
                exchangeRateLbl.Enabled = true;
                lblExchangeDate.Enabled = true;
                lblExchangeDate.Visible = true;
                exchangeRateLbl.Text = "The current Exchange rate is $1 USD to $" + exchangeRates["USD"] + " AUD";
                lblExchangeDate.Text = "Correct as of " + exchangeRateDate;
            }
            else if (selector == "E")
            {
                applicableExchangeRate = exchangeRates["EUR"];
                exchangeRateLbl.Visible = true;
                exchangeRateLbl.Enabled = true;
                lblExchangeDate.Enabled = true;
                lblExchangeDate.Visible = true;
                exchangeRateLbl.Text = "The current Exchange rate is €1 EUR to $" + exchangeRates["EUR"] + " AUD";
                lblExchangeDate.Text = "Correct as of " + exchangeRateDate;
            }
            //MessageBox.Show(applicableExchangeRate.ToString());
        }

        private void label13_Click(object sender, EventArgs e) { }
        private void costOfLiftTB_TextChanged(object sender, EventArgs e) { }

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
                //1x 20ft Container - $7000
                num20Ft = 0;
                shippingLbl20.Text = num20Ft + "x 20ft Container(s) - $0";
                num40Ft = 0;
                shippingLbl40.Text = num40Ft + "x 40ft Container(s) - $0";
                shippingLblTotal.Text = "Total of $0 for shipping";
            }
            else if (selector == 2)
            {
                num20Ft++;
                shippingLbl20.Text = num20Ft + "x 20ft Container(s) - $" + basePrices["20ftFreight"] * num20Ft;
            }
            else if (selector == 3)
            {
                num40Ft++;
                shippingLbl40.Text = num40Ft + "x 40ft Container(s) - $" + basePrices["40ftFreight"] * num40Ft;
            }
            freightTotal = (num20Ft * basePrices["20ftFreight"]) + (num40Ft * basePrices["40ftFreight"]);
            shippingLblTotal.Text = "Total of $" + freightTotal + " for shipping";
        }

        public void GeneratePriceList()
        {
            liftPrice = 0;

            //prices pulled from base prices dictionary
            PriceListFormatting(lblCost, float.Parse(tbCost.Text) * applicableExchangeRate);
            PriceListFormatting(lblSundries, float.Parse(tbSundries.Text));
            PriceListFormatting(lblShaft, float.Parse(tbShaftLight.Text));
            PriceListFormatting(lblDuct, float.Parse(tbDuct.Text));
            PriceListFormatting(lblAccommodation, float.Parse(tbAccomodation.Text));
            PriceListFormatting(lblStorage, float.Parse(tbStorage.Text));
            PriceListFormatting(lblTravel, float.Parse(tbTravel.Text));
            //prices pulled from text boxes of program
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
            PriceListFormatting(lblScaffold, (float.Parse(tbScaffold.Text) * basePrices["14Scaffolds"]));
            PriceListFormatting(lblEntrance, (float.Parse(tbEntranceGuards.Text) * float.Parse(tbWeeksRequired.Text) * basePrices["15EntranceGuards"]));
            //add security from base prices dictionary if box is checked
            if (cbSecurity.Checked)
            {
                PriceListFormatting(lblSecurity, (basePrices["Security"] + (basePrices["SecurityPerFloor"] * int.Parse(tBFloors.Text))));
            }
            else
            {
                PriceListFormatting(lblSecurity, 0);
            }
            //add freight based on number of required containers 
            PriceListFormatting(lblFreight, freightTotal);
            //add labour from the labour costs dictionary based on number of floors in the building 
            PriceListFormatting(lblLabour, labourPrice[int.Parse(tBFloors.Text)]);

            float marginPercent = (1 + (float.Parse(tbMargin.Text) / 100));
            liftPrice = liftPrice * int.Parse(tbNumberLifts.Text);
            lblCostOfParts.Text = ("$" + Math.Round(liftPrice, 2).ToString());
            lblCostIncludingMargin.Text = ("$" + Math.Round((liftPrice * marginPercent), 2).ToString());
            lblGST.Text = ("$" + Math.Round(((liftPrice * marginPercent) * 0.1), 2).ToString());
            lblPriceIncludingGST.Text = ("$" + Math.Round((liftPrice * marginPercent) * 1.1, 2).ToString());
        }

        public void PriceListFormatting(Label label, float cost)
        {
            if (cost > 0)
            {
                label.Text = "$" + cost;
                liftPrice += cost;
            }
            else
            {
                label.Text = "N/A";
                liftPrice += 0;
            }
        }

        private void labelLiftCurrency_Click(object sender, EventArgs e) { }

        private void label35_Click(object sender, EventArgs e) { }

        private void button1_Click(object sender, EventArgs e)
        {
            GeneratePriceList();
        }
    }
}