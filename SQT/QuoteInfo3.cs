using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo3 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo3()
        {
            InitializeComponent();
        }

        private void QuoteInfo3_Load(object sender, EventArgs e)
        {
            PullInfo();
        }

        private void PullInfo()
        {
            textBox2.Text = f.textBox2.Text;
            textBox4.Text = f.textBox4.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void buttonEUR_Click(object sender, EventArgs e)
        {
            //f.WordData("","");
            //call WordData method in form 1 to send all info into the dictiinary for writing 
            //open question form 3
            //close this form 
        }
    }
}
