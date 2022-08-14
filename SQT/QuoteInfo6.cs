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
            if (f.loadingPreviousData)
            {
                PullInfo();

            }
        }

        private void PullInfo()
        {
            try
            {
                f.LoadPreviousXmlTb(tbLoad, tbSpeed, tbwidth, tbDepth, tbHeight, tbLiftRating, tbNumofCarEntrances, tbFrontWallReturn, tbLiftCarNotes);
            }
            catch (Exception)
            {

                return;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo7 nF = new QuoteInfo7();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE134", f.MeasureStringChecker(tbLoad.Text, "kg")); //load
            f.WordData("AE135", f.MeasureStringChecker(tbSpeed.Text, "mps"));//speed
            f.WordData("AE136", f.MeasureStringChecker(tbwidth.Text, "mm")); // width
            f.WordData("AE137", f.MeasureStringChecker(tbDepth.Text, "mm"));//depth
            f.WordData("AE138", f.MeasureStringChecker(tbHeight.Text, "mm"));//height
            f.WordData("AE139", f.MeasureStringChecker(tbLiftRating.Text, "passengers"));//classification rating
            f.WordData("AE113", f.MeasureStringChecker(tbLiftRating.Text, "passenger"));//classification rating
            f.WordData("AE140", tbNumofCarEntrances.Text);//number of car entraces
            f.WordData("AE141", f.MeasureStringChecker(tbFrontWallReturn.Text, "mm"));//front wall return 
            if (tbFrontWallReturn.Text != "")
            {
                f.WordData("AE142", "NOTE: " + tbLiftCarNotes.Text);//notes
            }
            else
            {
                f.WordData("AE142", "");//notes
            }

            f.SaveTbToXML(tbDepth, tbFrontWallReturn, tbHeight, tbLiftCarNotes, tbLiftRating, tbLoad, tbNumofCarEntrances, tbSpeed, tbwidth);

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
