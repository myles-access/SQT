using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.NetworkInformation;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace SQT
{
    public partial class MainMenu : Form
    {
        #region VARS
        private readonly string passWord = "LiftFix";
        AdminPanel adminPanel = new AdminPanel();
        public string exchangeRateDate;
        public int maxFloorNumber;
        public bool networkConnected;
        public bool internetConnected;

        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();
        public Dictionary<string, string> exchangeRateURL = new Dictionary<string, string>();

        #endregion

        #region Loading Methods
        public MainMenu()
        {
            InitializeComponent();
        }

        #endregion

        #region Admin Menu
        private void btSQAT_Click(object sender, EventArgs e)
        {
            if (SQATPasswordCheck())
            {
                AdminPanel fAdminPanel = new AdminPanel();
                fAdminPanel.Show();
            }
        }

        private bool SQATPasswordCheck()
        {
            bool bPasswordCheck;
            string input = Interaction.InputBox("Please enter Password", "Password", "", 0, 0);

            bPasswordCheck = input == passWord;
            return bPasswordCheck;
        }
        #endregion

        #region Opening Quote Calculator
        private void btPinDiff_Click(object sender, EventArgs e)
        {
            btPinDiff.Enabled = false;
            progressBar1.Value = 0;
            progressBar1.Maximum = 6;
            progressBar1.Visible = true;
            basePrices.Clear();
            labourPrice.Clear();
            exchangeRates.Clear();
            exchangeRateURL.Clear();

            //check if software can connect to the network
            lbTitleText.Text = "Checking Network Conectivity";
            NetworkAccess();
            if(!networkConnected || !internetConnected) { return; };
            ProgressBarStep();

            //fetch the base prices for the calculator from the XML file
            lbTitleText.Text = "Fetching Base Prices";
            FetchBasePrices();
            ProgressBarStep();

            //fetch the labour prices from the XML file
            lbTitleText.Text = "Fetching Labour Rates";
            FetchLabourPrices();
            ProgressBarStep();

            //fetch the URL for exchange rates from the XML file
            lbTitleText.Text = "Checking Exchange Rates Server";
            FetchExchangeRateURL();
            ProgressBarStep();

            //Ready the exchange rate from either the website or offline stored values
            lbTitleText.Text = "Fetching Exchange Rates";
            CurrencyRates();
            ProgressBarStep();

            //check the date of the used exchange rate and return out if rate is too old. 
            if (IsExchangeRateOld())
            {
                return;
            }

            //load up the calculator form
            lbTitleText.Text = "Loading Calculator";
            Pin_Dif_Calc fPinDif = new Pin_Dif_Calc();
            ProgressBarStep();

            // reset the main menu form and show the calculator
            lbTitleText.Text = "Quote Calculator";
            progressBar1.Visible = false;
            btPinDiff.Enabled = true;
            fPinDif.Show();
        }

        private void ProgressBarStep(int stepValue = 1)
        {
            progressBar1.Value = progressBar1.Value + stepValue;
        }

        #endregion

        #region Exchange Rate Methods
        private void CurrencyRates()
        {
            if (networkConnected)
            {
                //grab rates from website
                FetchCurrencyRates();
                LoadStoredCurrencyRates();
            }
            else
            {
                LoadStoredCurrencyRates();
            }
        }

        private bool IsExchangeRateOld()
        {
            DateTime dT = Convert.ToDateTime(exchangeRateDate);
            TimeSpan tS = DateTime.Now - dT;

            if (tS.Days >= 7)
            {
                MessageBox.Show("Warning, the exchange rates being used are over 7 days old. Please connect to the internet to refresh");
                return true;
            }
            else
            {
                MessageBox.Show("Exchange Rate OK");
                return false;
            }
        }

        private void FetchExchangeRateURL()
        {
            string dKey = "";
            string dName = "";

            XmlTextReader XMLR = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\ExchangeRateURL.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Name")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }

                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "URL")
                {
                    dName = XMLR.ReadElementContentAsString();
                }

                if (dKey != "" && dName != "")
                {
                    exchangeRateURL.Add(dKey, dName);
                    dKey = "";
                    dName = "";
                }
            }
            XMLR.Close();
        }

        //find the live currency rates from floatrates.com and write them into the Exchange rate dictionary 
        private void FetchCurrencyRates()
        {
            string dKey = "";
            string dName = "";
            bool foundPubDate = false;
            string path = "X:\\Program Dependancies\\Quote tool\\CurrecyExchangeRate.xml";

            XmlTextReader XMLR = new XmlTextReader(exchangeRateURL["1ExchangeRateURL"]);
            XmlTextWriter XMLW = new XmlTextWriter(path, Encoding.UTF8);
            XMLW.Formatting = Formatting.Indented;
            XMLW.WriteStartDocument();

            XMLW.WriteStartElement("Data");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "targetCurrency")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "inverseRate")
                {
                    dName = XMLR.ReadElementContentAsString();
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "pubDate" && !foundPubDate)
                {
                    XMLW.WriteStartElement("Object");
                    XMLW.WriteElementString("pubDate", XMLR.ReadElementContentAsString());
                    XMLW.WriteEndElement(); //Object end
                    foundPubDate = true;
                }

                if (dKey != "" && dName != "")
                {
                    //exchangeRates.Add(dKey, dName * basePrices["16CurrencyMargin"]);

                    XMLW.WriteStartElement("Object");

                    XMLW.WriteElementString("targetCurrency", dKey);
                    XMLW.WriteElementString("inverseRate", dName.ToString());

                    XMLW.WriteEndElement(); //Object end

                    dKey = "";
                    dName = "";
                }
            }

            XMLW.WriteEndElement();//Data end
            XMLW.Close();
            XMLR.Close();
        }

        private void LoadStoredCurrencyRates()
        {
            string dKey = "";
            string dName = "";
            bool foundPubDate = false;

            XmlTextReader XMLR = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\CurrecyExchangeRate.xml");

            while (XMLR.Read())
            {
                if (!foundPubDate && XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "pubDate")
                {
                    exchangeRateDate = XMLR.ReadElementContentAsString();
                    foundPubDate = true;
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "targetCurrency")
                {
                    dKey = XMLR.ReadElementContentAsString();
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "inverseRate")
                {
                    dName = XMLR.ReadElementContentAsString();
                }

                if (dKey != "" && dName != "")
                {
                    exchangeRates.Add(dKey, float.Parse(dName) * basePrices["16CurrencyMargin"]);

                    dKey = "";
                    dName = "";
                }
            }

            XMLR.Close();
        }

        #endregion

        #region Network Conectivity 
        private void NetworkAccess()
        {
            networkConnected = LocalNetworkCheck();
            internetConnected = InternetCheck();

            if (!networkConnected || !internetConnected)
            {
                lbTitleText.Text = "Networks Connected";
            }
            else
            {
                lbTitleText.Text = "Networks NOT Connected";
                MessageBox.Show("Local Network = " + networkConnected.ToString() + "\n Internet = "+internetConnected.ToString());
            }
        }

        private bool InternetCheck()
        {
            var hostUrl = "www.accesselevators.com.au";

            Ping ping = new Ping();

            PingReply result = ping.Send(hostUrl);
            return result.Status == IPStatus.Success;
        }

        private bool LocalNetworkCheck()
        {
            try
            {
                Directory.GetAccessControl(@"X:\");
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region XML Functions
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
        }
        #endregion

        #region Unused Methods

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            //
        }

        public void CloseFormMethod()
        {
            //
        }

        private void MainMenu_Load(object sender, EventArgs e)
        {
            //
        }
        private void progressBar1_Click(object sender, EventArgs e)
        {
            //
        }
        #endregion
    }
}
