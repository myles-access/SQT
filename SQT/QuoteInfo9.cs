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
            if (f.loadingPreviousData)
            {
                PullInfo();
            }
        }

        private void PullInfo()
        {
            f.LoadPreviousXmlTb(tbNumberOfCOPS, tbMainCOPLocation, tbAuxCOPLocation, tbKeyswitchLocation, tbDesignations);
            f.LoadPreviousXmlRb(tbCOPFinish, rbCOPFinishOther, rbCOPFinishSatinStainlessSteel);
            f.LoadPreviousXmlRb(null, rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite);
            f.LoadPreviousXmlRb(null, rbPositionIndicatorTypeFlushMount, rbPositionIndicatorTypeSurfaceMount);
            f.LoadPreviousXmlRb(null, rbDoorOpenButtonYes, rbDoorOpenButtonNo);
            f.LoadPreviousXmlRb(null, rbDoorCLoseButtonYes, rbDoorCloseButtonNo);
            f.LoadPreviousXmlRb(null, rbTelephoneHandsFreeYes, rbTelephoneHandsFreeNo);
            f.LoadPreviousXmlRb(null, rbSecurityCabilingOnlyNo, rbSecurityCabilingOnlyYes);
            f.LoadPreviousXmlRb(null, rbBrailelTactileSymbolsYes, rbBraileTactileSymbolsNo);
            f.LoadPreviousXmlRb(null, rbFanSwitchNo, rbFanSwitchYes);
            f.LoadPreviousXmlRb(null, rbCarLightSwitchNo, rbCarLightSwitchYes);
            f.LoadPreviousXmlRb(null, rbFireSErviceKeySwitchNo, rbFireServiceKeySwitchYes);
            f.LoadPreviousXmlRb(null, rbExclusiveServiceNo, rbExclusiveServiceYes);
            f.LoadPreviousXmlRb(null, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes);
            f.LoadPreviousXmlRb(null, rbLiftOverloadIndicatorYes, rbLiftrOverloadIndicatorNo);
            f.LoadPreviousXmlRb(null, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes);
            f.LoadPreviousXmlRb(null, rbGPOInCarNo, rbGPOInCarYes);
            f.LoadPreviousXmlRb(null, rbAudibleIndicationGOngNo, rbAudibleIndicationGongYes);
            f.LoadPreviousXmlRb(null, rbVoiceAnnunciationNo, rbVoiceAnnunciationYes);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            QuoteInfo10 nF = new QuoteInfo10();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE168", tbNumberOfCOPS.Text); // number of COPS
            f.WordData("AE169", tbMainCOPLocation.Text);// main COP location
            f.WordData("AE170", tbAuxCOPLocation.Text);//aux cop location
            f.WordData("AE171", tbDesignations.Text); // designations 
            f.WordData("AE191", tbKeyswitchLocation.Text); //keyt switch location
            f.WordData("AE172", f.RadioButtonHandeler(tbCOPFinish, rbFanSwitchNo, rbCOPFinishSatinStainlessSteel));// COP finish
            f.WordData("AE173", "Dual illumination buttons with gong");// button type 
            f.WordData("AE174", f.RadioButtonHandeler(null, rbLEDColourRed, rbLEDColourBlue, rbLEDColourWhite));// LED colour
            f.WordData("AE175", f.RadioButtonHandeler(null, rbDoorOpenButtonNo, rbDoorOpenButtonYes));// door open button 
            f.WordData("AE176", f.RadioButtonHandeler(null, rbDoorCloseButtonNo, rbDoorCLoseButtonYes));//door close button 
            f.WordData("AE177", f.RadioButtonHandeler(null, rbTelephoneHandsFreeNo, rbTelephoneHandsFreeYes));//telephone hands free 
            f.WordData("AE178", f.RadioButtonHandeler(null, rbSecurityCabilingOnlyNo, rbSecurityCabilingOnlyYes));//security cabiling only 
            f.WordData("AE179", f.RadioButtonHandeler(null, rbBraileTactileSymbolsNo, rbBrailelTactileSymbolsYes));//braille tactile symbols 
            f.WordData("AE180", f.RadioButtonHandeler(null, rbFanSwitchNo, rbFanSwitchYes));// fan switch 
            f.WordData("AE181", f.RadioButtonHandeler(null, rbCarLightSwitchYes, rbCarLightSwitchNo));//car light switch
            f.WordData("AE182", f.RadioButtonHandeler(null, rbFireServiceKeySwitchYes, rbFireSErviceKeySwitchNo));// fire service key switch 
            f.WordData("AE183", f.RadioButtonHandeler(null, rbExclusiveServiceNo, rbExclusiveServiceYes));//exclusive service 
            f.WordData("AE184", f.RadioButtonHandeler(null, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes));// rear door kew switch 
            f.WordData("AE185", f.RadioButtonHandeler(null, rbLiftrOverloadIndicatorNo, rbLiftOverloadIndicatorYes));// lift overload indicator 
            f.WordData("AE186", f.RadioButtonHandeler(null, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes));//security key switch 
            f.WordData("AE187", f.RadioButtonHandeler(null, rbGPOInCarNo, rbGPOInCarYes));//GPO in car
            f.WordData("AE188", f.RadioButtonHandeler(null, rbAudibleIndicationGOngNo, rbAudibleIndicationGongYes));// audible indication gong 
            f.WordData("AE189", f.RadioButtonHandeler(null, rbVoiceAnnunciationNo, rbVoiceAnnunciationYes));//voice annunciation 
            f.WordData("AE190", f.RadioButtonHandeler(null, rbPositionIndicatorTypeSurfaceMount, rbPositionIndicatorTypeFlushMount));// position indicaor type 

            f.SaveTbToXML(tbAuxCOPLocation, tbCOPFinish, tbDesignations, tbKeyswitchLocation, tbMainCOPLocation, tbNumberOfCOPS);
            f.SaveRbToXML(rbAudibleIndicationGOngNo, rbAudibleIndicationGongYes, rbBrailelTactileSymbolsYes, rbBraileTactileSymbolsNo,
                rbCarLightSwitchNo, rbCarLightSwitchYes, rbCOPFinishOther, rbCOPFinishSatinStainlessSteel, rbDoorCloseButtonNo,
                rbDoorCLoseButtonYes, rbDoorOpenButtonNo, rbDoorOpenButtonYes, rbExclusiveServiceNo, rbExclusiveServiceYes, rbFanSwitchNo,
                rbFanSwitchYes, rbFireSErviceKeySwitchNo, rbFireServiceKeySwitchYes, rbGPOInCarNo, rbGPOInCarYes, rbLEDColourBlue,
                rbLEDColourRed, rbLEDColourWhite, rbLiftOverloadIndicatorYes, rbLiftrOverloadIndicatorNo, rbPositionIndicatorTypeFlushMount,
                rbPositionIndicatorTypeSurfaceMount, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes, rbSecurityCabilingOnlyNo,
                rbSecurityCabilingOnlyYes, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes, rbTelephoneHandsFreeNo, rbTelephoneHandsFreeYes,
                rbVoiceAnnunciationNo, rbVoiceAnnunciationYes);

            //Load next form and close this one 
            nF.Show();
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            //
        }
    }
}
