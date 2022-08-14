using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo10 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo10()
        {
            InitializeComponent();
        }

        private void QuoteInfo10_Load(object sender, EventArgs e)
        {
            if (f.loadingPreviousData)
            {
                PullInfo();
            }
        }

        private void PullInfo()
        {
            try
            {
                f.LoadPreviousXmlTb(tbNumOfButtonRisers);
                f.LoadPreviousXmlRb(tbFacePlateMaterial, rbFacePlateMaterialOther, rbFacePlateMaterialSatinStainlessSteel);
                f.LoadPreviousXmlRb(null, rbHallLanternsNo, rbHallLanternsYes);
                f.LoadPreviousXmlRb(null, rbLandingFireServiceKeySwitchNo, rbLandingFireServiceKeySwitchYes);
                f.LoadPreviousXmlRb(null, rbBraileTactileSymbolsLandingNo, rbBraileTactileSymbolsLandingYes);
                f.LoadPreviousXmlRb(null, rbOutofServiceKeySwitchNo, rbOutOfServieKeySwitchYes);
                f.LoadPreviousXmlRb(null, rbDigitalIndicationIncorperatedNo, rbDigitalIndicationIncorperatedYes);

            }
            catch (Exception)
            {

                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

            QuoteInfo11 nF = new QuoteInfo11();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE192", f.RadioButtonHandeler(tbFacePlateMaterial, rbFacePlateMaterialSatinStainlessSteel,  rbFacePlateMaterialOther));//face plate material 
            f.WordData("AE193", "Dual illumination buttons with gong");//button type
            f.WordData("AE194", f.RadioButtonHandeler(null, rbHallLanternsNo, rbHallLanternsYes));//hall lanterns
            f.WordData("AE195", f.RadioButtonHandeler(null, rbBraileTactileSymbolsLandingNo, rbBraileTactileSymbolsLandingYes));// braille tactile symbols 
            f.WordData("AE196", f.RadioButtonHandeler(null, rbDigitalIndicationIncorperatedYes, rbDigitalIndicationIncorperatedNo));// digital indication incorperated 
            f.WordData("AE197", f.RadioButtonHandeler(null, rbOutofServiceKeySwitchNo, rbOutOfServieKeySwitchYes));//out of service key switch 
            f.WordData("AE198", f.RadioButtonHandeler(null, rbLandingFireServiceKeySwitchNo, rbLandingFireServiceKeySwitchYes));// fire service key switch 
            f.WordData("AE217", tbNumOfButtonRisers.Text);// number of button risers 

            f.SaveTbToXML(tbFacePlateMaterial, tbNumOfButtonRisers);
            f.SaveRbToXML(rbBraileTactileSymbolsLandingNo, rbBraileTactileSymbolsLandingYes, rbDigitalIndicationIncorperatedNo,
                rbDigitalIndicationIncorperatedYes, rbFacePlateMaterialOther, rbFacePlateMaterialSatinStainlessSteel, rbLandingFireServiceKeySwitchNo,
                rbLandingFireServiceKeySwitchYes, rbHallLanternsNo, rbHallLanternsYes, rbOutofServiceKeySwitchNo, rbOutOfServieKeySwitchYes);

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void label4_Click(object sender, EventArgs e)
        {
            //
        }
    }
}
