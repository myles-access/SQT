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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tBQuoteNumber.Text = ("Qu" + DateTime.Now.Year.ToString("yy") + "-000");

        }

        public void HideAll()
        {
            // a method to hide all the textboxes so that the relevent selected textboxes can be shown on top
        }
    }
}
