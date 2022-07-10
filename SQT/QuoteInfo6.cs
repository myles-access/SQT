using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo6 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo6()
        {
            InitializeComponent();
        }

        private void QuoteInfo6_Load(object sender, EventArgs e)
        {
            PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo7 nF = new QuoteInfo7();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE134", f.MeasureStringChecker(textBox1.Text, "kg")); //load
            f.WordData("AE135", f.MeasureStringChecker(textBox2.Text, "mps"));//speed
            f.WordData("AE136", f.MeasureStringChecker(textBox7.Text, "mm")); // width
            f.WordData("AE137", f.MeasureStringChecker(textBox5.Text, "mm"));//depth
            f.WordData("AE138", f.MeasureStringChecker(textBox4.Text, "mm"));//height
            f.WordData("AE139", f.MeasureStringChecker(textBox3.Text, "passengers"));//classification rating
            f.WordData("AE140", textBox6.Text);//number of car entraces
            f.WordData("AE141", f.MeasureStringChecker(textBox8.Text, "mm"));//front wall return 
            f.WordData("AE142", "NOTE: " + textBox9.Text);//notes

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
