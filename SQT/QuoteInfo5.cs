﻿using System;
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
    public partial class QuoteInfo5 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo5()
        {
            InitializeComponent();
        }

        private void QuoteInfo5_Load(object sender, EventArgs e)
        {
            PullInfo();
        }

        private void PullInfo()
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void buttonEUR_Click(object sender, EventArgs e)
        {
            QuoteInfo6 nF = new QuoteInfo6();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 

            //Load next form and close this one 
            nF.Show();
            this.Close();
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {

            QuoteInfo6 nF = new QuoteInfo6();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 

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
