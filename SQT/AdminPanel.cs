using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace SQT
{
    public partial class AdminPanel : Form
    {
        #region VARS
        //VARS
        public Dictionary<int, float> labourPrice = new Dictionary<int, float>();
        public Dictionary<string, float> basePrices = new Dictionary<string, float>();
        //readonly passwordGate f = Application.OpenForms.OfType<passwordGate>().Single();
        #endregion

        #region Load Methods
        public AdminPanel()
        {
            InitializeComponent();
            this.FormClosing += new FormClosingEventHandler(myForm_FormClosing);
        }

        private void AdminPanel_Load(object sender, EventArgs e)
        {
            FetchBasePrices();
            FetchLabourPrices();
            SetTextToDict();

            //f.Hide();
        }
        #endregion

        #region XML Methods
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
            //MessageBox.Show("BASE");
            XMLR.Close();
        }

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

                if (XMLR.NodeType == XmlNodeType.Element && XMLR.Name == "Price")
                {
                    dName = float.Parse(XMLR.ReadElementContentAsString());

                }

                if (dKey != -1 && dName != -1)
                {
                    labourPrice.Add(dKey, dName);
                    dKey = -1;
                    dName = -1;

                }
            }
            XMLR.Close();
        }

        private void BasePricesXMLWriter()
        {
            XmlTextWriter xmlWriter = new XmlTextWriter("X:\\Program Dependancies\\Quote tool\\QuotePriceList.xml", Encoding.UTF8);
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("prices");

            foreach (KeyValuePair<string, float> bP in basePrices)
            {
                xmlWriter.WriteStartElement("item");

                xmlWriter.WriteElementString("costItem", bP.Key);
                xmlWriter.WriteElementString("price", bP.Value.ToString());

                xmlWriter.WriteEndElement(); // item element end
            }

            xmlWriter.WriteEndElement(); // prices element end
            xmlWriter.Close();
        }

        private void LabourPricesXMLWriter()
        {
            XmlTextWriter xmlWriter = new XmlTextWriter("X:\\Program Dependancies\\Quote tool\\LabourCosts.xml", Encoding.UTF8);
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("Costs");

            foreach (KeyValuePair<int, float> lC in labourPrice)
            {
                xmlWriter.WriteStartElement("Cost");

                xmlWriter.WriteElementString("Floors", lC.Key.ToString());
                xmlWriter.WriteElementString("Price", lC.Value.ToString());

                xmlWriter.WriteEndElement(); //Cost element end
            }

            xmlWriter.WriteEndElement(); //Costs element end 
            xmlWriter.Close();
        }
        #endregion

        #region Button Click Methods
        private void btnClose_Click(object sender, EventArgs e)
        {
            var mainForm = Application.OpenForms.OfType<MainMenu>().Single();
            mainForm.CloseFormMethod();
        }
        private void btnReset_Click(object sender, EventArgs e)
        {
            SetTextToDict();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.BackColor = Color.Red;
            DialogResult dR = MessageBox.Show("Are you sure you wish to save?", "Save?", MessageBoxButtons.YesNo);
            if (dR == DialogResult.Yes)
            {
                UpdateDict();
                BasePricesXMLWriter();
                LabourPricesXMLWriter();
                MessageBox.Show("Saving Successful", "Success");
            }
            else
            {
                MessageBox.Show("Saving Failed", "Failed");
            }
            btnSave.BackColor = Color.LightGreen;

        }
        #endregion

        #region Unused Methods
        private void label1_Click(object sender, EventArgs e) { }
        private void label33_Click(object sender, EventArgs e) { }
        #endregion

        void myForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            var mainForm = Application.OpenForms.OfType<MainMenu>().Single();
            mainForm.CloseFormMethod();
        }

        private void UpdateDict()
        {
            //update base price dictionary values
            basePrices["1CarFinishes"] = float.Parse(tbCarFinishes.Text);
            basePrices["2FireExtinguisher"] = float.Parse(tbFireExtinguishers.Text);
            basePrices["3GSMUnitPhone"] = float.Parse(tbGSM.Text);
            basePrices["4ProtectiveBlanket"] = float.Parse(tbProtectiveBlanket.Text);
            basePrices["5SumpCover"] = float.Parse(tbSumpCover.Text);
            basePrices["6Wiring"] = float.Parse(tbWiring.Text);
            basePrices["7Sinage"] = float.Parse(tbSinage.Text);
            basePrices["8ElectricalBox"] = float.Parse(tbElectricalBox.Text);
            basePrices["9Drawings"] = float.Parse(tbDrawings.Text);
            basePrices["10ForkLift"] = float.Parse(tbForkLift.Text);
            basePrices["11Maintenance"] = float.Parse(tbMaintenance.Text);
            basePrices["12Manuals"] = float.Parse(tbManuals.Text);
            basePrices["13WorkcoverFees"] = float.Parse(tbWorkcoverFees.Text);
            basePrices["20ftFreight"] = float.Parse(tb20ftFreight.Text);
            basePrices["40ftFreight"] = float.Parse(tb40ftFreight.Text);
            basePrices["14Scaffolds"] = float.Parse(tbScaffolds.Text);
            basePrices["15EntranceGuards"] = float.Parse(tbEntranceGuards.Text);
            basePrices["16CurrencyMargin"] = float.Parse(tbCurrencyMargin.Text);
            basePrices["17LowestMargin"] = float.Parse(tbMinMargin.Text);
            basePrices["18DefaultMargin"] = float.Parse(tbDefMargin.Text);

            //update labour costs dictionary values
            labourPrice[int.Parse("2")] = int.Parse(tbLabour2.Text);
            labourPrice[int.Parse("3")] = int.Parse(tbLabour3.Text);
            labourPrice[int.Parse("4")] = int.Parse(tbLabour4.Text);
            labourPrice[int.Parse("5")] = int.Parse(tbLabour5.Text);
            labourPrice[int.Parse("6")] = int.Parse(tbLabour6.Text);
            labourPrice[int.Parse("7")] = int.Parse(tbLabour7.Text);
            labourPrice[int.Parse("8")] = int.Parse(tbLabour8.Text);
            labourPrice[int.Parse("9")] = int.Parse(tbLabour9.Text);
            labourPrice[int.Parse("10")] = int.Parse(tbLabour10.Text);
            labourPrice[int.Parse("11")] = int.Parse(tbLabour11.Text);
            labourPrice[int.Parse("12")] = int.Parse(tbLabour12.Text);
            labourPrice[int.Parse("13")] = int.Parse(tbLabour13.Text);
            labourPrice[int.Parse("14")] = int.Parse(tbLabour14.Text);
            labourPrice[int.Parse("15")] = int.Parse(tbLabour15.Text);
            labourPrice[int.Parse("16")] = int.Parse(tbLabour16.Text);
        }

        private void SetTextToDict()
        {
            //update text box values based off base prices dictionary
            tbCarFinishes.Text = basePrices["1CarFinishes"].ToString();
            tbFireExtinguishers.Text = basePrices["2FireExtinguisher"].ToString();
            tbGSM.Text = basePrices["3GSMUnitPhone"].ToString();
            tbProtectiveBlanket.Text = basePrices["4ProtectiveBlanket"].ToString();
            tbSumpCover.Text = basePrices["5SumpCover"].ToString();
            tbWiring.Text = basePrices["6Wiring"].ToString();
            tbSinage.Text = basePrices["7Sinage"].ToString();
            tbElectricalBox.Text = basePrices["8ElectricalBox"].ToString();
            tbDrawings.Text = basePrices["9Drawings"].ToString();
            tbForkLift.Text = basePrices["10ForkLift"].ToString();
            tbMaintenance.Text = basePrices["11Maintenance"].ToString();
            tbManuals.Text = basePrices["12Manuals"].ToString();
            tbWorkcoverFees.Text = basePrices["13WorkcoverFees"].ToString();
            tb20ftFreight.Text = basePrices["20ftFreight"].ToString();
            tb40ftFreight.Text = basePrices["40ftFreight"].ToString();
            tbScaffolds.Text = basePrices["14Scaffolds"].ToString();
            tbEntranceGuards.Text = basePrices["15EntranceGuards"].ToString();
            tbCurrencyMargin.Text = basePrices["16CurrencyMargin"].ToString();
            tbMinMargin.Text = basePrices["17LowestMargin"].ToString();
            tbDefMargin.Text = basePrices["18DefaultMargin"].ToString();

            // update text boxes based on labour costs dictionary
            tbLabour2.Text = labourPrice[int.Parse("2")].ToString();
            tbLabour3.Text = labourPrice[int.Parse("3")].ToString();
            tbLabour4.Text = labourPrice[int.Parse("4")].ToString();
            tbLabour5.Text = labourPrice[int.Parse("6")].ToString();
            tbLabour6.Text = labourPrice[int.Parse("6")].ToString();
            tbLabour7.Text = labourPrice[int.Parse("7")].ToString();
            tbLabour8.Text = labourPrice[int.Parse("8")].ToString();
            tbLabour9.Text = labourPrice[int.Parse("9")].ToString();
            tbLabour10.Text = labourPrice[int.Parse("10")].ToString();
            tbLabour11.Text = labourPrice[int.Parse("11")].ToString();
            tbLabour12.Text = labourPrice[int.Parse("12")].ToString();
            tbLabour13.Text = labourPrice[int.Parse("13")].ToString();
            tbLabour14.Text = labourPrice[int.Parse("14")].ToString();
            tbLabour15.Text = labourPrice[int.Parse("15")].ToString();
            tbLabour16.Text = labourPrice[int.Parse("16")].ToString();
        }

        private void btnClose_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}

