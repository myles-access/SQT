using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo7 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo7()
        {
            InitializeComponent();
        }

        private void QuoteInfo7_Load(object sender, EventArgs e)
        {
            if (f.loadingPreviousData)
            {
                PullInfo();
            }
        }

        private void PullInfo()
        {
            f.LoadPreviousXmlTb(tbDoorWidth, tbDoorHeight, tbDoorVoltage, tbLandingDoorJambDepth);
            f.LoadPreviousXmlRb(null, rbDoorTypeCentreOpening, rbDoorTypeSideOpening);
            f.LoadPreviousXmlRb(null, rbDoorNudgingYes, rbDoorNudgingNo);
            f.LoadPreviousXmlRb(null, rbAdvancedOpeningNo, rbAdvancedOpeningYes);
            f.LoadPreviousXmlRb(null, rbFireRatedDoorsNo, rbFireRatedDoorsYes);
            f.LoadPreviousXmlRb(tbEntrancProtection, rbEntranceProtectionElectronicScanner, rbEntranceProtectionOTher);
            f.LoadPreviousXmlRb(tbLandingDoorFinish, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel);
            f.LoadPreviousXmlRb(tbDoorTracks, rbDoorTracksAnodisedAluminium, rbDoorTracksOther);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo8 nF = new QuoteInfo8();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE143", f.MeasureStringChecker(tbDoorWidth.Text, "mm"));//door width 
            f.WordData("AE144", f.MeasureStringChecker(tbDoorHeight.Text, "mm")); //door height 
            f.WordData("AE145", f.MeasureStringChecker(tbDoorVoltage.Text, "v"));//door operator voltage 
            f.WordData("AE146", f.MeasureStringChecker(tbLandingDoorJambDepth.Text, "mm"));//landing door jamb depth
            f.WordData("AE147", f.RadioButtonHandeler(tbLandingDoorFinish, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel));//landing door finish
            f.WordData("AE148", f.RadioButtonHandeler(null, rbDoorTypeCentreOpening, rbDoorTypeSideOpening));//door type
            f.WordData("AE150", f.RadioButtonHandeler(tbDoorTracks, rbDoorTracksAnodisedAluminium, rbDoorTracksOther));// door tracks 
            f.WordData("AE151", f.RadioButtonHandeler(null, rbAdvancedOpeningNo, rbAdvancedOpeningYes));//advanced opening
            f.WordData("AE152", f.RadioButtonHandeler(null, rbDoorNudgingNo, rbDoorNudgingYes));//door nudging 
            f.WordData("AE153", f.RadioButtonHandeler(null, rbFireRatedDoorsYes, rbFireRatedDoorsNo));//fire rated doors 
            f.WordData("AE154", f.RadioButtonHandeler(tbEntrancProtection, rbEntranceProtectionOTher, rbEntranceProtectionElectronicScanner)); //  entrance protection 

            //Load next form and close this one 
            nF.Show();
            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //
        }
    }
}
