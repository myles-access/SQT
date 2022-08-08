using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo4 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo4()
        {
            InitializeComponent();
        }

        private void QuoteInfo4_Load(object sender, EventArgs e)
        {
            if (f.loadingPreviousData)
            {
                PullInfo();
            }
        }

        private void PullInfo()
        {
            f.LoadPreviousXmlRb(null, rbIndependentServiceYes, rbIndependentServiceNo);
            f.LoadPreviousXmlRb(null, rbLoadWeighingNo, rbLoadWeighingYes);
            f.LoadPreviousXmlRb(tbmachinetype, rbMachineTypeGearless, rbMachinetypeOther);
            f.LoadPreviousXmlRb(null, rbFireServiceNo, rbFireServiceYes);
            f.LoadPreviousXmlRb(tbControlerLocation, rbControlerLoactionTopLanding, rbControlerLocationBottomLanding, rbControlerlocationShaft, rbControlerLocationOther);
            f.LoadPreviousXmlRb(tbDriveType, rbDriveTypeMRL, rbDriveTypeOther);
            f.LoadPreviousXmlRb(null, rbEmergencyPowerOperationNo, rbEmergencyPowerOperationYes);
            f.LoadPreviousXmlTb(tbEmergencyPowerOperationText);
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {

            QuoteInfo5 nF = new QuoteInfo5();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE116", f.RadioButtonHandeler(null, rbIndependentServiceYes, rbIndependentServiceNo));//independent service
            f.WordData("AE117", f.RadioButtonHandeler(null, rbLoadWeighingNo, rbLoadWeighingYes));//load weighing
            f.WordData("AE118", f.RadioButtonHandeler(tbControlerLocation, rbControlerLoactionTopLanding, rbControlerlocationShaft, rbControlerLocationBottomLanding, rbControlerLocationOther));//controler location
            f.WordData("AE119", f.RadioButtonHandeler(tbmachinetype, rbMachinetypeOther, rbMachineTypeGearless));//machine type
            f.WordData("AE120", f.RadioButtonHandeler(null, rbFireServiceNo, rbFireServiceYes));//fire service
            f.WordData("AE121", f.RadioButtonHandeler(null, rbEmergencyPowerOperationNo, rbEmergencyPowerOperationYes));//emergency power operation
            f.WordData("AE122", tbEmergencyPowerOperationText.Text);//emergency power operation text
            f.WordData("AE219", f.RadioButtonHandeler(tbDriveType, rbDriveTypeOther, rbDriveTypeMRL)); // drive type

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }
    }
}
