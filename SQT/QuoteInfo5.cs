using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo5 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo5()
        {
            InitializeComponent();
        }

        private void QuoteInfo5_Load(object sender, EventArgs e)
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

                f.LoadPreviousXmlTb(tbShaftDepth, tbShaftWidth, tbPitDepth, tbHeadroom, tbTravel, tbNumofLandingDoors, tbNumofLandings);
                f.LoadPreviousXmlRb(tbStructureShaft, rbStructureShaftConcrete, rbStructureShaftOther);
                f.LoadPreviousXmlRb(tbFixings, rbFixingsOther, rbFixingsTrueBolts);
                f.LoadPreviousXmlRb(null, rbTrimmerBeamsNo, rbTrimmerBeamsYes);
                f.LoadPreviousXmlRb(null, rbFalseFloorNo, rbFalseFloorYes);
            }
            catch (Exception)
            {

                return;
            }
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {
            QuoteInfo6 nF = new QuoteInfo6();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE123", f.MeasureStringChecker(tbShaftWidth.Text, "mm"));// shaft width
            f.WordData("AE124", f.MeasureStringChecker(tbShaftDepth.Text, "mm"));//shaft depth
            f.WordData("AE125", f.MeasureStringChecker(tbPitDepth.Text, "mm"));//pit depth
            f.WordData("AE126", f.MeasureStringChecker(tbHeadroom.Text, "mm"));//headroom
            f.WordData("AE127", f.MeasureStringChecker(tbTravel.Text, "mm"));//travel
            f.WordData("AE128", tbNumofLandings.Text);// number of landings
            f.WordData("AE129", tbNumofLandingDoors.Text);//number of landing doors 
            f.WordData("AE130", f.RadioButtonHandeler(tbStructureShaft, rbStructureShaftConcrete, rbStructureShaftOther)); //structure shaft 
            f.WordData("AE131", f.RadioButtonHandeler(tbFixings, rbFixingsTrueBolts, rbFixingsOther));//fixings
            f.WordData("AE132", f.RadioButtonHandeler(null, rbTrimmerBeamsNo, rbTrimmerBeamsYes));//trimmer beams
            f.WordData("AE133", f.RadioButtonHandeler(null, rbFalseFloorYes, rbFalseFloorNo));//false floor

            f.SaveTbToXML(tbFixings, tbHeadroom, tbNumofLandingDoors, tbNumofLandings, tbPitDepth,
                tbShaftDepth, tbShaftWidth, tbStructureShaft, tbTravel);
            f.SaveRbToXML(rbFalseFloorNo, rbFalseFloorYes, rbFixingsOther, rbFixingsTrueBolts, rbStructureShaftConcrete,
                rbStructureShaftOther, rbTrimmerBeamsNo, rbTrimmerBeamsYes);

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {

            f.QuestionCloseCall(this);
        }

        private void groupBox5_Enter(object sender, EventArgs e)
        {
            //
        }
    }
}
