using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo2 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();

        public QuoteInfo2()
        {
            InitializeComponent();
        }

        private void QuoteInfo2_Load(object sender, EventArgs e)
        {
           // PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {
            QuoteInfo3 nF = new QuoteInfo3();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE105", tbfname.Text); //first name
            f.WordData("AE106", tblname.Text);//last name
            f.WordData("AE107", tbphone.Text);//phone number
            f.WordData("AE108", textBox3.Text);//address 1
            f.WordData("AE109", textBox1.Text);//address 2
            f.WordData("AE110", textBox2.Text);//address 3

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
