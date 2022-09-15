﻿using System;
using System.Linq;
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
            if (f.loadingPreviousData)
            {
                PullInfo();
            }
        }

        private void PullInfo()
        {
            f.LoadPreviousXmlTb(tbModel, tbLiftNumbers, tbTypeofLift);
            f.LoadPreviousXmlRb(null, rbSL, rbSumasa, rbWittur);
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {
            QuoteInfo4 nF = new QuoteInfo4();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE111", f.RadioButtonHandeler(null, rbSL, rbWittur, rbSumasa)); //supplier
            f.WordData("AE112", tbModel.Text);//model
            //f.WordData("AE113",textBox1.Text);//capacity
            f.WordData("AE114", tbLiftNumbers.Text);//lift number
            f.WordData("AE115", tbTypeofLift.Text);//type of lift
            f.WordData("AE215", "Full Collective"); //control type, not changable 

            f.SaveRbToXML(rbSL, rbSumasa, rbWittur);
            f.SaveTbToXML(tbModel, tbTypeofLift, tbLiftNumbers);

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //
        }
    }
}