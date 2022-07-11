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
           // PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {

            QuoteInfo5 nF = new QuoteInfo5();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE116", f.RadioButtonHandeler(null, radioButton1, radioButton2));//independent service
            f.WordData("AE117", f.RadioButtonHandeler(null, radioButton5, radioButton6));//load weighing
            f.WordData("AE118", f.RadioButtonHandeler(textBox2, radioButton11, radioButton12, radioButton13, radioButton14));//controler location
            f.WordData("AE119", f.RadioButtonHandeler(textBox3, radioButton9, radioButton10));//machine type
            f.WordData("AE120", f.RadioButtonHandeler(null, radioButton3, radioButton4));//fire service
            f.WordData("AE121", f.RadioButtonHandeler(null, radioButton7, radioButton8));//emergency power operation
            f.WordData("AE122", textBox1.Text);//emergency power operation text
            f.WordData("AE219", f.RadioButtonHandeler(textBox4, radioButton15, radioButton16)); // drive type

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
