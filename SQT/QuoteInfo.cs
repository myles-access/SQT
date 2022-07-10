using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();

        public QuoteInfo()
        {
            InitializeComponent();
        }

        private void QuoteInfo_Load(object sender, EventArgs e)
        {
            PullInfo();
        }

        private void PullInfo()
        {
            tBAddress.Text = f.tBAddress.Text;
            tBQuoteNumber.Text = f.tBQuoteNumber.Text;
            tbNumberLifts.Text = f.tbNumberLifts.Text;
            tBFloors.Text = f.tBFloors.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void buttonEUR_Click(object sender, EventArgs e)
        {
            QuoteInfo2 nF = new QuoteInfo2();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictionary for writing 
            f.WordData("AE101", tBAddress.Text); //address
            f.WordData("AE102", tBQuoteNumber.Text);//quote number
            f.WordData("AE103", tbNumberLifts.Text);//number of lifts
            f.WordData("AE104", tBFloors.Text);//number of floors

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }
    }
}
