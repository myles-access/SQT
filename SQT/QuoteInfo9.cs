using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class QuoteInfo9 : Form
    {
        Form1 f = Application.OpenForms.OfType<Form1>().Single();
        public QuoteInfo9()
        {
            InitializeComponent();
        }

        private void QuoteInfo9_Load(object sender, EventArgs e)
        {
            PullInfo();
        }

        private void PullInfo()
        {
            //
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo10 nF = new QuoteInfo10();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE168", tbfname.Text); // number of COPS
            f.WordData("AE169", textBox1.Text);// main COP location
            f.WordData("AE170", textBox2.Text);//aux cop location
            f.WordData("AE171", textBox3.Text); // designations 
            f.WordData("AE172", f.RadioButtonHandeler(textBox6, radioButton15, radioButton19));// COP finish
            f.WordData("AE173", "Dual illumination buttons with gong");// button type 
            f.WordData("AE174", f.RadioButtonHandeler(null, radioButton3, radioButton4, radioButton1));// LED colour
            f.WordData("AE175", f.RadioButtonHandeler(null, radioButton5, radioButton6));// door open button 
            f.WordData("AE176", f.RadioButtonHandeler(null, radioButton7, radioButton8));//door close button 
            f.WordData("AE177", f.RadioButtonHandeler(null, radioButton9, radioButton10));//telephone hands free 
            f.WordData("AE178", f.RadioButtonHandeler(null, radioButton11, radioButton12));//security cabiling only 
            f.WordData("AE179", f.RadioButtonHandeler(null, radioButton13, radioButton14));//braille tactile symbols 
            f.WordData("AE180", f.RadioButtonHandeler(null, radioButton15, radioButton16));// fan switch 
            f.WordData("AE181", f.RadioButtonHandeler(null, radioButton18, radioButton17));//car light switch
            f.WordData("AE182", f.RadioButtonHandeler(null, radioButton22, radioButton21));// fire service key switch 
            f.WordData("AE183", f.RadioButtonHandeler(null, radioButton23, radioButton24));//exclusive service 
            f.WordData("AE184", f.RadioButtonHandeler(null, radioButton25, radioButton26));// rear door kew switch 
            f.WordData("AE185", f.RadioButtonHandeler(null, radioButton27, radioButton28));// lift overload indicator 
            f.WordData("AE186", f.RadioButtonHandeler(null, radioButton29, radioButton30));//security key switch 
            f.WordData("AE187", f.RadioButtonHandeler(null, radioButton31, radioButton32));//GPO in car
            f.WordData("AE188", f.RadioButtonHandeler(null, radioButton33, radioButton34));// audible indication gong 
            f.WordData("AE189", f.RadioButtonHandeler(null, radioButton35, radioButton36));//voice annunciation 
            f.WordData("AE190", f.RadioButtonHandeler(null, radioButton37, radioButton38));// position indicaor type 

            //Load next form and close this one 
            nF.Show();
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }
    }
}
