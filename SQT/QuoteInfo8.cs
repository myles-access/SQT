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
            PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo9 nF = new QuoteInfo9();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE155", f.RadioButtonHandeler(textBox3, radioButton1, radioButton8));//car door finish
            f.WordData("AE156", f.RadioButtonHandeler(textBox8, radioButton6, radioButton27, radioButton28, radioButton5));//ceiling finish
            f.WordData("AE157", f.RadioButtonHandeler(null, radioButton13, radioButton14));//false ceiling
            f.WordData("AE158", f.RadioButtonHandeler(null, radioButton18, radioButton17));//bump rail
            f.WordData("AE159", textBox9.Text);//floor
            f.WordData("AE160", f.RadioButtonHandeler(textBox4, radioButton2, radioButton7));//front wall
            f.WordData("AE161", f.RadioButtonHandeler(textBox10, radioButton10, radioButton9, radioButton21));//mirror
            f.WordData("AE162", f.RadioButtonHandeler(textBox2, radioButton15, radioButton16));//handrail
            f.WordData("AE163", @"Natural & Mechanical");// ventelation fan
            f.WordData("AE164", f.RadioButtonHandeler(textBox5, radioButton4, radioButton3)); //side wall 
            f.WordData("AE165", textBox1.Text + "LED Lights"); // lighting 
            f.WordData("AE166", f.RadioButtonHandeler(null, radioButton12, radioButton11)); // skirting
            f.WordData("AE167", f.RadioButtonHandeler(null, radioButton23, radioButton24)); // protective blankets 
            f.WordData("AE216", f.RadioButtonHandeler(textBox6, radioButton25, radioButton26)); //  rear wall

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

        }
    }
}
