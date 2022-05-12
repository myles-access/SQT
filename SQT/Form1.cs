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

namespace SQT
{
    public partial class Form1 : Form
    {
        // VARS
        public string quoteNumber = "";
        public float applicableExchangeRate = 1;
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();
        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public int num20Ft;
        public int num40Ft;
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
        }

        private void FetchCurrencyRates()
        {
            //bool grabRate = false;
            //bool switchEUR = false;
            string dKey = "";
            float dName = 0;
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
                    exchangeRates.Add(dKey, dName * basePrices["Currency Margin"]);
                    dKey = "";
                    dName = 0;
                }
            }
            //MessageBox.Show("EXCHANGE");
        }
        private void FetchBasePrices()
        {
            string dKey = "";
            float dName = 0;

            XmlTextReader XMLR = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\QuotePriceList.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "name")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }

                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "price")
                {
                    dName = float.Parse(XMLR.ReadElementContentAsString());
                }

                if (dKey != "" && dName != 0)
                {
                    basePrices.Add(dKey, dName);
                    dKey = "";
                    dName = 0;
                }
            }
            //MessageBox.Show("BASE");
        }
        private void FetchLabourPrices()
        {
            int dKey = 0;
            float dName = 0;

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

                if (dKey != 0 && dName != 0)
                {
                    labourPrice.Add(dKey, dName);
                    dKey = 0;
                    dName = 0;
                }
            }
            //MessageBox.Show("BASE");
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
                exchangeRateLbl.Text = "The current Exchange rate is $1 USD to $" + exchangeRates["USD"] + " AUD";
            }
            else if (selector == "E")
            {
                applicableExchangeRate = exchangeRates["EUR"];
                exchangeRateLbl.Visible = true;
                exchangeRateLbl.Enabled = true;
                exchangeRateLbl.Text = "The current Exchange rate is €1 EUR to $" + exchangeRates["EUR"] + " AUD";
                //MessageBox.Show(applicableExchangeRate.ToString());
            }
        }

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
                shippingLbl20.Text = num20Ft + "x 20ft Container(s) - $" + basePrices["20ft Freight"] * num20Ft;
            }
            else if (selector == 3)
            {
                num40Ft++;
                shippingLbl40.Text = num40Ft + "x 40ft Container(s) - $" + basePrices["40ft Freight"] * num40Ft;
            }
            shippingLblTotal.Text = "Total of $" + (basePrices["20ft Freight"] * num20Ft + basePrices["40ft Freight"] * num40Ft) + " for shipping";
        }
        private void GeneratePriceList()
        {
                PriceListFormatting(lblCost, float.Parse(tbCost.Text) * applicableExchangeRate);
                        PriceListFormatting(lblFinishes, basePrices["Car Finishes"]);
            PriceListFormatting(lblFire, basePrices["Fire Extinguisher"]);
            PriceListFormatting(lblGSM, basePrices["GSM Unit / Phone"]);
            PriceListFormatting(lblBlanket, basePrices["Protective Blanket"]);
            PriceListFormatting(lblSump, basePrices["Sump Cover"]);
            PriceListFormatting(lblSundries, float.Parse(tbSundries.Text));
            PriceListFormatting(lblWiring, basePrices["Wiring"]);

        }
        public void PriceListFormatting(Label label, float cost)
        {
            if (cost != 0)
            {
                label.Text = "$" + cost;
                liftPrice += cost;
            }
            else
            {
                label.Text = "N/A";
            }
        }

        private void labelLiftCurrency_Click(object sender, EventArgs e)
        {

        }

        private void label35_Click(object sender, EventArgs e)
        {

        }
    }
}