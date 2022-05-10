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

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
           // currencySelectionGroup.Visible = false;
            FetchCurrencyRates();
            FetchBasePrices();
            //tBAddress.Text = "$" + exchangeRates["USD"] + "  $" + exchangeRates["EUR"].ToString();
            quoteNumber = ("Qu" + DateTime.Now.ToString("yy") + "-000");
            tBQuoteNumber.Text = quoteNumber;
        }

#pragma warning disable IDE1006 // Naming Styles
        private void tBQuoteNumber_TextChanged(object sender, EventArgs e)
#pragma warning restore IDE1006 // Naming Styles
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

            XmlTextReader currencyXML = new XmlTextReader("http://www.floatrates.com/daily/aud.xml");
            while (currencyXML.Read())
            {
                if (currencyXML.NodeType == XmlNodeType.Element && currencyXML.Name == "targetCurrency")
                {
                    dKey = currencyXML.ReadElementContentAsString();
                }

                if (currencyXML.NodeType == XmlNodeType.Element && currencyXML.Name == "inverseRate")
                {
                    dName = float.Parse(currencyXML.ReadElementContentAsString());
                }

                if (dKey != "" && dName != 0)
                {
                    exchangeRates.Add(dKey, dName);
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

            XmlTextReader basePricesXML = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\QuotePriceList.xml");
            while (basePricesXML.Read())
            {
                if (basePricesXML.NodeType == XmlNodeType.Element && basePricesXML.Name == "name")
                {
                    dKey = basePricesXML.ReadElementContentAsString();
                }

                if (basePricesXML.NodeType == XmlNodeType.Element && basePricesXML.Name == "price")
                {
                    dName = float.Parse(basePricesXML.ReadElementContentAsString());
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
            }
            else if (selector == "E")
            {
                applicableExchangeRate = exchangeRates["EUR"];
                MessageBox.Show(applicableExchangeRate.ToString());
            }            
        }

        private void buttonAUD_Click(object sender, EventArgs e) { SelectCurrency("A"); }

        private void buttonUSD_Click(object sender, EventArgs e) { SelectCurrency("U"); }

        private void buttonEUR_Click(object sender, EventArgs e) { SelectCurrency("E"); }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void costOfLiftTB_TextChanged(object sender, EventArgs e)
        {
        }
    }
}
