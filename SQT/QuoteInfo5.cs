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
           // PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {
            QuoteInfo6 nF = new QuoteInfo6();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE123", f.MeasureStringChecker(tbfname.Text, "mm"));// shaft width
            f.WordData("AE124", f.MeasureStringChecker(tbphone.Text, "mm"));//shaft depth
            f.WordData("AE125", f.MeasureStringChecker(textBox2.Text, "mm"));//pit depth
            f.WordData("AE126", f.MeasureStringChecker(textBox1.Text, "mm"));//headroom
            f.WordData("AE127", f.MeasureStringChecker(textBox4.Text, "mm"));//travel
            f.WordData("AE128", textBox3.Text);// number of landings
            f.WordData("AE129", textBox5.Text);//number of landing doors 
            f.WordData("AE130", f.RadioButtonHandeler(textBox6, radioButton10, radioButton9)); //structure shaft 
            f.WordData("AE131", f.RadioButtonHandeler(textBox7, radioButton1, radioButton4));//fixings
            f.WordData("AE132", f.RadioButtonHandeler(null, radioButton5, radioButton6));//trimmer beams
            f.WordData("AE133", f.RadioButtonHandeler(null, radioButton3, radioButton2));//false floor

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
