using Microsoft.VisualBasic;
using System;
using System.Windows.Forms;

namespace SQT
{
    public partial class MainMenu : Form
    {
        private readonly string passWord = "LiftFix";
        AdminPanel adminPanel = new AdminPanel();

        public MainMenu()
        {
            InitializeComponent();
        }

        private void MainMenu_Load(object sender, EventArgs e)
        {

        }

        private void btSQAT_Click(object sender, EventArgs e)
        {
            if (SQATPasswordCheck())
            {
                AdminPanel fAdminPanel = new AdminPanel();
                fAdminPanel.Show();
            }
        }

        private bool SQATPasswordCheck()
        {
            bool bPasswordCheck;
            string input = Interaction.InputBox("Please enter Password", "Password", "", 0, 0);

            bPasswordCheck = input == passWord;
            return bPasswordCheck;
        }

        private void button1_Click(object sender, EventArgs e) // pintaric single 
        {
            Form1 fForm1 = new Form1();
            fForm1.Show();
        }

        public void CloseFormMethod()
        {

        }

        private void btPinMulti_Click(object sender, EventArgs e)
        {
            Pin_Mul_Exp fPinMul = new Pin_Mul_Exp();
            fPinMul.Show();
        }

        private void btLamSingle_Click(object sender, EventArgs e)
        {
            Lam_Sing_Calc fLamSing = new Lam_Sing_Calc();
            fLamSing.Show();
        }

        private void btLamMulti_Click(object sender, EventArgs e)
        {
            Lam_Mul_Calc fLamMul = new Lam_Mul_Calc();
            fLamMul.Show();
        }

        private void btPinDiff_Click(object sender, EventArgs e)
        {
            Pin_Dif_Calc fPinDif = new Pin_Dif_Calc();
            fPinDif.Show();
        }

        private void btLamDif_Click(object sender, EventArgs e)
        {
            Lam_Dif_Calc fLamDif = new Lam_Dif_Calc();
            fLamDif.Show();
        }
    }
}
