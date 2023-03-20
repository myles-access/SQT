using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace SQT
{
    public partial class MainMenu : Form
    {
        private readonly string passWord = "LiftFix";
        AdminPanel adminPanel = new AdminPanel();
        public string exchangeRateDate;
        public int maxFloorNumber;
        public bool networkConnected;

        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();


        public MainMenu()
        {
            InitializeComponent();
        }

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

        private void btPinDiff_Click(object sender, EventArgs e)
        {
            btPinDiff.Enabled = false;
            progressBar1.Visible = true;
            progressBar1.Maximum = 5;

            lbTitleText.Text = "Checking Network Conectivity";
            networkConnected = NetworkCheck();
            if (networkConnected)
            {
                MessageBox.Show("NETWORK SUCCESS");
            }
            ProgressBarStep();

            lbTitleText.Text = "Loading Calculator";
            Pin_Dif_Calc fPinDif = new Pin_Dif_Calc();
            ProgressBarStep();

            lbTitleText.Text = "Fetching Base Prices";
            FetchBasePrices();
            ProgressBarStep();

            lbTitleText.Text = "Fetching Exchange Rates";
            FetchCurrencyRates();
            ProgressBarStep();

            lbTitleText.Text = "Fetching Labour Rates";
            FetchLabourPrices();
            ProgressBarStep();

            lbTitleText.Text = "Quote Calculator";
            progressBar1.Visible = false;
            btPinDiff.Enabled = true;
            fPinDif.Show();
        }

        private void ProgressBarStep (int stepValue = 1)
        {
            progressBar1.Value = progressBar1.Value + stepValue;
        }

        #region Network Conectivity 
        private bool NetworkCheck()
        {
            var hostUrl = "www.floatrates.com";

            Ping ping = new Ping();

            PingReply result = ping.Send(hostUrl);
            return result.Status == IPStatus.Success;
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
