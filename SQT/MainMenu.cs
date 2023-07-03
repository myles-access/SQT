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
        private string listboxSelected;
        public int maxFloorNumber;
        public bool networkConnected;
        public bool internetConnected;
        public bool loadMenuOpen = false;

        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public Dictionary<string, float> exchangeRates = new Dictionary<string, float>();
        public Dictionary<string, string> exchangeRateURL = new Dictionary<string, string>();


        private string[] xmlFiles = Directory.GetFiles(@"X:\Program Dependancies\Quote tool\Previous Prices", "*.*", SearchOption.AllDirectories);

        public string quNumber;
        #endregion

        #region Loading Methods
        public MainMenu()
        {
            InitializeComponent();
            this.Size = new System.Drawing.Size(449, 523);
            panelShipping.Visible = false;
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
            OpenQuoteCalculator(false);
        }

        private void OpenQuoteCalculator(bool loadingOldQuote = false, string fileLoadName = "")
        {
            bool startedSucessfully = true;
            btPinDiff.Enabled = false;
            btnLoadOldQuote.Enabled = false;
            progressBar1.Value = 0;
            progressBar1.Maximum = 7;
            progressBar1.Visible = true;
            basePrices.Clear();
            labourPrice.Clear();
            exchangeRates.Clear();
            exchangeRateURL.Clear();

            //check if software can connect to the network
            lbTitleText.Text = "Checking Network Conectivity";
            NetworkAccess();
            if (!networkConnected || !internetConnected) { return; };
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
                startedSucessfully = false;
            }

            //load up the calculator form
            lbTitleText.Text = "Loading Calculator";
            Pin_Dif_Calc fPinDif = new Pin_Dif_Calc();
            fPinDif.Show();
            ProgressBarStep();

            if (loadingOldQuote)
            {
                //load previous quote values into form, if errored close form and tell user
                lbTitleText.Text = "Importing Quote Values";
                if (fileLoadName == "" || !fPinDif.LoadProcessManager(fileLoadName))
                {
                    startedSucessfully = false;
                }
            }
            else
            {
                //claim a Qu number for the quote
                lbTitleText.Text = "Claiming Qu Number";
                ClaimQuNumber();
                fPinDif.tBMainQuoteNumber.Text = quNumber;
            }
            ProgressBarStep();

            // reset the main menu form and show the calculator
            lbTitleText.Text = "Quote Calculator";
            progressBar1.Visible = false;
            btPinDiff.Enabled = true;
            btnLoadOldQuote.Enabled = true;
            if (!startedSucessfully)
            {
                MessageBox.Show("ERROR: Could not load quoting tool.");
                fPinDif.Close();
            }
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
                lbTitleText.Text = "Networks NOT Connected";
            }
            else
            {
                lbTitleText.Text = "Networks Connected";
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

        #region Qu Number Claiming
        private void ClaimQuNumber()
        {
            string yearNum = "";
            string nextQu = "";

            //read XML to get next avaliable number
            XmlTextReader XMLR = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\QuNum.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Year")
                {
                    yearNum = XMLR.ReadElementContentAsString();
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Qu")
                {
                    nextQu = XMLR.ReadElementContentAsString();
                }
            }
            XMLR.Close();

            //correct year of Qu number if less than the current year 
            if (int.Parse(yearNum) < int.Parse(DateTime.Now.ToString("yy")))
            {
                yearNum = DateTime.Now.ToString("yy");
                nextQu = "001";
            }

            // set the Qu Number var to the concainated number string
            if (yearNum != "" && nextQu != "")
            {
                quNumber = ("Qu" + yearNum + "-" + nextQu);
            }

            //write to the XML file to claim the number
            XmlTextWriter XMLW = new XmlTextWriter("X:\\Program Dependancies\\Quote tool\\QuNum.xml", Encoding.UTF8);
            XMLW.Formatting = Formatting.Indented;
            XMLW.WriteStartDocument();
            XMLW.WriteStartElement("QuVar");
            XMLW.WriteElementString("Year", yearNum.ToString());
            int i = int.Parse(nextQu) + 1;
            XMLW.WriteElementString("Qu", i.ToString("000"));
            XMLW.WriteEndElement(); //QuVar end

            XMLW.Close();

            quNumber = "Qu" + yearNum + "-" + nextQu;
            AddQuToSalesmanRecord(quNumber);
        }

        private void AddQuToSalesmanRecord(string claimedQu)
        {
            //string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string userName = Environment.UserName;
            //Dictionary<String, List<string>> salesRecord = new Dictionary<string, List<string>>();

            List<string> salesmanList = new List<string>();
            List<List<string>> jobsList = new List<List<string>>();
            int salesmanTracker = -1;
            bool salesmanMatch = false;

            XmlTextReader XMLR = new XmlTextReader("X:\\Program Dependancies\\Quote tool\\SalesmanRecord.xml");
            while (XMLR.Read())
            {
                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Salesman")
                {
                    string s = XMLR.ReadElementContentAsString();
                    salesmanList.Add(s);
                    jobsList.Add(new List<string> { });
                    salesmanTracker++;
                    if (s == userName && !salesmanMatch)
                    {
                        salesmanMatch = true;
                        jobsList[salesmanTracker].Add(claimedQu);

                    }
                }
                else if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Qu")
                {
                    jobsList[salesmanTracker].Add(XMLR.ReadElementContentAsString());

                }
            }
            if (!salesmanMatch)
            {
                salesmanList.Add(userName);
                jobsList.Add(new List<string> { });
                salesmanTracker++;
                jobsList[salesmanTracker].Add(claimedQu);
            }

            XMLR.Close();

            //write to the XML file to record the number
            XmlTextWriter XMLW = new XmlTextWriter("X:\\Program Dependancies\\Quote tool\\SalesmanRecord.xml", Encoding.UTF8);
            XMLW.Formatting = Formatting.Indented;
            XMLW.WriteStartDocument();
            XMLW.WriteStartElement("QuoteRecords");
            int i = 0;
            foreach (List<string> sublist in jobsList)
            {
                XMLW.WriteStartElement("Record");
                XMLW.WriteElementString("Salesman", salesmanList[i]);
                i++;

                foreach (string s in sublist)
                {
                    XMLW.WriteElementString("Qu", s);

                }
                XMLW.WriteEndElement(); // Record Entry end

            }
            XMLW.WriteEndElement(); // Quote records end
            XMLW.Close();

        }
        #endregion

        #region Loading Old Quote
        private void btnLoadSelectedITem_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                OpenQuoteCalculator(true, listBox1.SelectedItem.ToString());
                OpenLoadMenu(false);
            }
        }

        private void btnLoadOldQuote_Click(object sender, EventArgs e)
        {
            loadMenuOpen = !loadMenuOpen;
            OpenLoadMenu(loadMenuOpen);
        }

        private void OpenLoadMenu(bool openingmenu)
        {
            if (openingmenu)
            {
                xmlFiles = Directory.GetFiles(@"X:\Program Dependancies\Quote tool\Previous Prices", "*.*", SearchOption.AllDirectories);
                ArrayToListBox(listBox1, xmlFiles);
                panelShipping.Visible = true;
                this.Size = new System.Drawing.Size(895, 523);
                tbLoadSearch.Text = "";
                loadMenuOpen = true;
            }
            else if (!openingmenu)
            {
                this.Size = new System.Drawing.Size(449, 523);
                panelShipping.Visible = false;
                loadMenuOpen = false;
            }
        }

        private void ArrayToListBox(ListBox lb, String[] s, string searchCriteria = "")
        {
            lb.Items.Clear();
            List<String> list = new List<String>();

            if (searchCriteria == "")
            {
                foreach (string arrayString in s)
                {
                    lb.Items.Add(Path.GetFileName(arrayString));
                }
            }
            else
            {
                foreach (string arrayString in s)
                {
                    bool contains = arrayString.IndexOf(searchCriteria, StringComparison.OrdinalIgnoreCase) >= 0;
                    if (contains)
                    {
                        lb.Items.Add(Path.GetFileName(arrayString));
                    }
                }
            }
        }

        private void tbLoadSearch_TextChanged(object sender, EventArgs e)
        {
            ArrayToListBox(listBox1, xmlFiles, tbLoadSearch.Text);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem.ToString() == listboxSelected)
            {
                OpenQuoteCalculator(true, listBox1.SelectedItem.ToString());
                OpenLoadMenu(false);
            }
            listboxSelected = listBox1.SelectedItem.ToString();
        }

        private void extraCostsClose_Click(object sender, EventArgs e)
        {
            OpenLoadMenu(false);
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