using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo8 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo8()
        {
            InitializeComponent();
        }

        private void QuoteInfo8_Load(object sender, EventArgs e)
        {
            if (f.loadingPreviousData)
            {
                PullInfo();
            }

            if (f.tbBlankets.Text != "0")
            {
                rbProtectriveBlanketsYes.Checked = true;
            }
            else
            {
                rbProtectiveBlanketsNo.Checked = true;
            }
        }

        private void PullInfo()
        {
            f.LoadPreviousXmlTb(tbNumOfLEDLights, tbFloorFinish);
            f.LoadPreviousXmlRb(tbCarDoorFinish, rbCarDoorFInishBrushedStainlessSteel, rbCarDoorFinishOther);
            f.LoadPreviousXmlRb(tbCeilingFinish, rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishOther, rbCeilingFinishWhite, rbCeilingFinishMirrorStainlessSteel);
            f.LoadPreviousXmlRb(null, rbBumpRailNo, rbBumpRailYes);
            f.LoadPreviousXmlRb(null, rbFalseCeilingNo, rbFalseCeilingYes);
            f.LoadPreviousXmlRb(tbFrontWall, rbFrontWallBrushedStainlessSteel, tbFrontWallOther);
            f.LoadPreviousXmlRb(tbMirror, rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther);
            f.LoadPreviousXmlRb(tbHandrail, rbHandrailBrushedStainlessSTeel, rbHandrailOther);
            f.LoadPreviousXmlRb(tbSideWall, rbSideWallBrushedStainlessSteel, rbSideWallOther);
            f.LoadPreviousXmlRb(tbRearWall, rbRearWallBrushedStainlessSteel, rbRearWallOther);
            f.LoadPreviousXmlRb(null, rbSkirtingNo, rbSkirtingYes);
            f.LoadPreviousXmlRb(null, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo9 nF = new QuoteInfo9();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE155", f.RadioButtonHandeler(tbCarDoorFinish, rbCarDoorFinishOther, rbCarDoorFInishBrushedStainlessSteel));//car door finish
            f.WordData("AE156", f.RadioButtonHandeler(tbCeilingFinish, rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishWhite, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
            f.WordData("AE157", f.RadioButtonHandeler(null, rbFalseCeilingNo, rbFalseCeilingYes));//false ceiling
            f.WordData("AE158", f.RadioButtonHandeler(null, rbBumpRailYes, rbBumpRailNo));//bump rail
            f.WordData("AE159", tbFloorFinish.Text);//floor
            f.WordData("AE160", f.RadioButtonHandeler(tbFrontWall, rbFrontWallBrushedStainlessSteel, tbFrontWallOther));//front wall
            f.WordData("AE161", f.RadioButtonHandeler(tbMirror, rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther));//mirror
            f.WordData("AE162", f.RadioButtonHandeler(tbHandrail, rbHandrailBrushedStainlessSTeel, rbHandrailOther));//handrail
            f.WordData("AE163", @"Natural & Mechanical");// ventelation fan
            f.WordData("AE164", f.RadioButtonHandeler(tbSideWall, rbSideWallBrushedStainlessSteel, rbSideWallOther)); //side wall 
            f.WordData("AE165", tbNumOfLEDLights.Text + " LED Lights"); // lighting 
            f.WordData("AE166", f.RadioButtonHandeler(null, rbSkirtingYes, rbSkirtingNo)); // skirting
            f.WordData("AE167", f.RadioButtonHandeler(null, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes)); // protective blankets 
            f.WordData("AE216", f.RadioButtonHandeler(tbRearWall, rbRearWallOther, rbRearWallBrushedStainlessSteel)); //  rear wall

            f.SaveTbToXML(tbCarDoorFinish, tbCeilingFinish, tbFloorFinish, tbFrontWall, tbHandrail, tbMirror, tbNumOfLEDLights, tbRearWall, tbSideWall);
            f.SaveRbToXML(tbFrontWallOther, rbBumpRailNo, rbBumpRailYes, rbCarDoorFInishBrushedStainlessSteel, rbCarDoorFinishOther,
                rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther, rbCeilingFinishWhite, rbFalseCeilingNo,
                rbFalseCeilingYes, rbFrontWallBrushedStainlessSteel, rbHandrailBrushedStainlessSTeel, rbHandrailOther, rbMirrorFullSize, rbMirrorHalfSize,
                rbMirrorOther, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes, rbRearWallBrushedStainlessSteel, rbRearWallOther,
                rbSideWallBrushedStainlessSteel, rbSideWallOther, rbSkirtingNo, rbSkirtingYes);

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {
            //
        }

    }
}
