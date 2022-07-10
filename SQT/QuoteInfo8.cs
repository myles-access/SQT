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
            f.WordData("AE155", f.RadioButtonHandeler(null, radioButton7, radioButton8));//car door finish
            f.WordData("AE156", f.RadioButtonHandeler(null, radioButton6, radioButton5));//ceiling finish
            f.WordData("AE157", f.RadioButtonHandeler(null, radioButton13, radioButton14));//false ceiling
            f.WordData("AE158", f.RadioButtonHandeler(null, radioButton18, radioButton17));//bump rail
            f.WordData("AE159", f.RadioButtonHandeler(null, radioButton21, radioButton22));//floor
            f.WordData("AE160", f.RadioButtonHandeler(null, radioButton2, radioButton1));//front wall
            f.WordData("AE161", f.RadioButtonHandeler(tbphone, radioButton10, radioButton9));//mirror
            f.WordData("AE162", f.RadioButtonHandeler(null, radioButton15, radioButton16));//handrail
            f.WordData("AE163", f.CheckBoxHandler(checkBox1, checkBox2));// ventelation fan
            f.WordData("AE164", f.RadioButtonHandeler(null, radioButton4, radioButton3)); //side wall 
            f.WordData("AE165", textBox1.Text + " " + f.RadioButtonHandeler(null, radioButton11, radioButton12)); // lighting 
            f.WordData("AE166", f.RadioButtonHandeler(null, radioButton20, radioButton19)); // skirting
            f.WordData("AE167", f.RadioButtonHandeler(null, radioButton23, radioButton24)); // protective blankets 
            f.WordData("AE216", f.RadioButtonHandeler(null, radioButton25, radioButton26)); //  rear wall

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
