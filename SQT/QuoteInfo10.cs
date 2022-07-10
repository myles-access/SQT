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
            PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void button3_Click(object sender, EventArgs e)
        {

            QuoteInfo11 nF = new QuoteInfo11();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE192", f.RadioButtonHandeler(null, radioButton19, radioButton20));//face plate material 
            f.WordData("AE193", f.RadioButtonHandeler(null, radioButton2, radioButton1));//button type
            f.WordData("AE194", f.RadioButtonHandeler(null, radioButton7, radioButton8));//hall lanterns
            f.WordData("AE195", f.RadioButtonHandeler(null, radioButton13, radioButton14));// braille tactile symbols 
            f.WordData("AE196", f.RadioButtonHandeler(null, radioButton3, radioButton4));// digital indication incorperated 
            f.WordData("AE197", f.RadioButtonHandeler(null, radioButton17, radioButton18));//out of service key switch 
            f.WordData("AE198", f.RadioButtonHandeler(null, radioButton5, radioButton6));// fire service key switch 
            f.WordData("AE217", tbfname.Text);// number of button risers 

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            f.QuestionCloseCall(this);
        }
    }
}
