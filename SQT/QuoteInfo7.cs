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
            PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo8 nF = new QuoteInfo8();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE143", f.MeasureStringChecker(tbfname.Text, "mm"));//door width 
            f.WordData("AE144", f.MeasureStringChecker(tbphone.Text, "mm")); //door height 
            f.WordData("AE145", f.MeasureStringChecker(textBox1.Text, "V"));//door operator voltage 
            f.WordData("AE146", f.MeasureStringChecker(textBox2.Text, "mm"));//landing door jamb depth
            f.WordData("AE147", f.RadioButtonHandeler(textBox3, radioButton7, radioButton8));//landing door finish
            f.WordData("AE148", f.RadioButtonHandeler(null, radioButton5, radioButton3));//door type
            f.WordData("AE150", f.RadioButtonHandeler(textBox5, radioButton1, radioButton10));// door tracks 
            f.WordData("AE151", f.RadioButtonHandeler(null, radioButton4, radioButton6));//advanced opening
            f.WordData("AE152", f.RadioButtonHandeler(null, radioButton11, radioButton12));//door nudging 
            f.WordData("AE153", f.RadioButtonHandeler(null, radioButton14, radioButton13));//fire rated doors 
            f.WordData("AE154", f.RadioButtonHandeler(textBox4, radioButton9, radioButton2)); //  entrance protection 

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
