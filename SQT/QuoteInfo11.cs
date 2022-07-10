﻿using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo11 : Form
    {
        readonly Form1 f = Application.OpenForms.OfType<Form1>().Single();

        public QuoteInfo11()
        {
            InitializeComponent();
        }

        private void QuoteInfo11_Load(object sender, EventArgs e)
        {
            PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void buttonEUR_Click_1(object sender, EventArgs e)
        {
            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE199", f.RadioButtonHandeler(null, radioButton20, radioButton19));//supply of true bolts
            f.WordData("AE200", f.RadioButtonHandeler(null, radioButton2, radioButton1));//lift shaft lighting 
            f.WordData("AE201", f.RadioButtonHandeler(null, radioButton3, radioButton4));//pit ladder 
            f.WordData("AE202", f.RadioButtonHandeler(null, radioButton5, radioButton6));//sump cover
            f.WordData("AE203", f.RadioButtonHandeler(null, radioButton7, radioButton8));//temp entrance screens 
            f.WordData("AE204", f.RadioButtonHandeler(null, radioButton15, radioButton16));// lifting eye beams 
            f.WordData("AE205", f.RadioButtonHandeler(null, radioButton9, radioButton10));//grouting of door frames
            f.WordData("AE206", f.RadioButtonHandeler(null, radioButton11, radioButton12));//button and indicator boxes 
            f.WordData("AE207", f.RadioButtonHandeler(null, radioButton17, radioButton18));// manuals a4 size
            f.WordData("AE208", f.RadioButtonHandeler(null, radioButton13, radioButton14));//supply of scaffold 
            f.WordData("AE209", f.RadioButtonHandeler(null, radioButton22, radioButton21));// emergency lowering system 
            f.WordData("AE210", f.RadioButtonHandeler(null, radioButton23, radioButton24));//out of service 
            f.WordData("AE218", "NOTE: " + tbfname.Text);//general notes

            //Load next form and close this one 
            f.QuestionsComplete();
            this.Close();

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }
    }
}
