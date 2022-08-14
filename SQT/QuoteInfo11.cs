using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo11 : Form
    {
        readonly Form1 f = Application.OpenForms.OfType<Form1>().Single();

        public QuoteInfo11()
        {
            InitializeComponent();
        }

        private void QuoteInfo11_Load(object sender, EventArgs e)
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

                f.LoadPreviousXmlTb(tbGeneralNotes);
                f.LoadPreviousXmlRb(tbSupplyofTrueBolts, rbSupplyofTrueBoltsByAccess, rbSupplyofTrueBoltsOther);
                f.LoadPreviousXmlRb(tbLiftShaftLighting, rbLiftShaftLightingByAccess, rbLiftShaftLightingOther);
                f.LoadPreviousXmlRb(tbPitLadder, rbPitLAdderbyAccess, rbPitLadderOther);
                f.LoadPreviousXmlRb(tbSumpCover, rbSumpCoverByAccess, rbSumpCoverOther);
                f.LoadPreviousXmlRb(tbTempEntranceScreens, rbTempEntranceScreensByAccess, rbTempEntranceScreensOther);
                f.LoadPreviousXmlRb(tbLiftingEyeBeam, rbLiftingEyeBeamByAccess, rbLiftingEyeBeamOther);
                f.LoadPreviousXmlRb(tbButtonandIndicatorBoxes, rbButtonandIndicatorBoxesByAccess, rbButtonandIndicatorBoxesOther);
                f.LoadPreviousXmlRb(tbSupplyofScaffold, rbSupplyofScaffoldByAccess, rbSupplyofScaffoldOther);
                f.LoadPreviousXmlRb(null, rbManualsA4SizeNo, rbManualsA4SizeYes);
                f.LoadPreviousXmlRb(null, rbOutofServiceNo, rbOutofServiceYes);
                f.LoadPreviousXmlRb(null, rbEmergencyLoweringSystemNo, rbEmergencyLoweringSystemYes);
            }
            catch (Exception)
            {

                return;
            }
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {
            this.Enabled = false;
            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE199", f.RadioButtonHandeler(tbSupplyofTrueBolts, rbSupplyofTrueBoltsOther, rbSupplyofTrueBoltsByAccess));//supply of true bolts
            f.WordData("AE200", f.RadioButtonHandeler(tbLiftShaftLighting, rbLiftShaftLightingOther, rbLiftShaftLightingByAccess));//lift shaft lighting 
            f.WordData("AE201", f.RadioButtonHandeler(tbPitLadder, rbPitLAdderbyAccess, rbPitLadderOther));//pit ladder 
            f.WordData("AE202", f.RadioButtonHandeler(tbSumpCover, rbSumpCoverByAccess, rbSumpCoverOther));//sump cover
            f.WordData("AE203", f.RadioButtonHandeler(tbTempEntranceScreens, rbTempEntranceScreensByAccess, rbTempEntranceScreensOther));//temp entrance screens 
            f.WordData("AE204", f.RadioButtonHandeler(tbLiftingEyeBeam, rbLiftingEyeBeamByAccess, rbLiftingEyeBeamOther));// lifting eye beams 
            f.WordData("AE206", f.RadioButtonHandeler(tbButtonandIndicatorBoxes, rbButtonandIndicatorBoxesByAccess, rbButtonandIndicatorBoxesOther));//button and indicator boxes 
            f.WordData("AE207", f.RadioButtonHandeler(null, rbManualsA4SizeNo, rbManualsA4SizeYes));// manuals a4 size
            f.WordData("AE208", f.RadioButtonHandeler(tbSupplyofScaffold, rbSupplyofScaffoldByAccess, rbSupplyofScaffoldOther));//supply of scaffold 
            f.WordData("AE209", f.RadioButtonHandeler(null, rbEmergencyLoweringSystemYes, rbEmergencyLoweringSystemNo));// emergency lowering system 
            f.WordData("AE210", f.RadioButtonHandeler(null, rbOutofServiceNo, rbOutofServiceYes));//out of service 
            if (tbGeneralNotes.Text != "")
            {
                f.WordData("AE218", "NOTE: " + tbGeneralNotes.Text);//general notes
            }
            else
            {
                f.WordData("AE218", "");//general notes
            }

            f.SaveTbToXML(tbButtonandIndicatorBoxes, tbGeneralNotes, tbLiftingEyeBeam, tbLiftShaftLighting, tbPitLadder, tbSumpCover, 
                tbSupplyofScaffold, tbSupplyofTrueBolts, tbTempEntranceScreens);
            f.SaveRbToXML(rbButtonandIndicatorBoxesByAccess, rbButtonandIndicatorBoxesOther, rbEmergencyLoweringSystemNo, 
                rbEmergencyLoweringSystemYes, rbLiftingEyeBeamByAccess, rbLiftingEyeBeamOther, rbLiftShaftLightingByAccess, rbLiftShaftLightingOther, 
                rbManualsA4SizeNo, rbManualsA4SizeYes, rbOutofServiceNo, rbOutofServiceYes, rbPitLAdderbyAccess, rbPitLadderOther, 
                rbSumpCoverByAccess, rbSumpCoverOther, rbSupplyofScaffoldByAccess, rbSupplyofScaffoldOther, rbSupplyofTrueBoltsByAccess,
                rbSupplyofTrueBoltsOther, rbTempEntranceScreensByAccess, rbTempEntranceScreensOther);

            //this.Hide();
            //Load next form and close this one 
            f.QuestionsComplete();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }
    }
}
