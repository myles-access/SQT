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
    public partial class Lam_Sing_Exp : Form
    {
        Lam_Sing_Calc f = Application.OpenForms.OfType<Lam_Sing_Calc>().Single();

        public Lam_Sing_Exp()
        {
            InitializeComponent();
        }

        private void Lam_Sing_Exp_Load(object sender, EventArgs e)
        {
            this.Enabled = true;
            if (f.loadingPreviousData)
            {
                PullInfo();
            }

            if (f.tbMainBlankets.Text != "0")
            {
                rbProtectriveBlanketsYes.Checked = true;
            }
            else
            {
                rbProtectiveBlanketsNo.Checked = true;
            }
        }

        private void PullInfo()
        {
            try
            {
                f.LoadPreviousXmlTb(tbfname, tblname, tbphone, tbAddress1, tbAddress2, tbAddress3, tbLiftNumbers, tbTypeofLift,
                    tbShaftDepth, tbShaftWidth, tbPitDepth, tbHeadroom, tbTravel, tbNumofLandingDoors, tbNumofLandings,
                    tbNumberOfCOPS, tbMainCOPLocation, tbAuxCOPLocation, tbKeyswitchLocation, tbDesignations, tbNumOfLEDLights, tbFloorFinish,
                    tbDoorWidth, tbDoorHeight, tbLoad, tbSpeed, tbwidth, tbDepth, tbHeight, tbLiftRating, tbNumofCarEntrances, tbLiftCarNotes);

                f.LoadPreviousXmlRb(null, rbSL, rbSumasa, rbWittur);
                f.LoadPreviousXmlRb(null, rbIndependentServiceYes, rbIndependentServiceNo);
                f.LoadPreviousXmlRb(null, rbLoadWeighingNo, rbLoadWeighingYes);
                f.LoadPreviousXmlRb(null, rbFireServiceNo, rbFireServiceYes);
                f.LoadPreviousXmlRb(tbControlerLocation, rbControlerLoactionTopLanding, rbControlerLocationBottomLanding, rbControlerlocationShaft, rbControlerLocationOther);
                f.LoadPreviousXmlRb(tbStructureShaft, rbStructureShaftConcrete, rbStructureShaftOther);
                f.LoadPreviousXmlRb(null, rbTrimmerBeamsNo, rbTrimmerBeamsYes);
                f.LoadPreviousXmlRb(null, rbFalseFloorNo, rbFalseFloorYes);
                f.LoadPreviousXmlRb(null, rbDoorTypeCentreOpening, rbDoorTypeSideOpening);
                f.LoadPreviousXmlRb(null, rbDoorNudgingYes, rbDoorNudgingNo);
                f.LoadPreviousXmlRb(null, rbAdvancedOpeningNo, rbAdvancedOpeningYes);
                f.LoadPreviousXmlRb(tbLandingDoorFinish, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel);
                f.LoadPreviousXmlRb(tbDoorTracks, rbDoorTracksAnodisedAluminium, rbDoorTracksOther);
                f.LoadPreviousXmlRb(tbCarDoorFinish, rbCarDoorFInishBrushedStainlessSteel, rbCarDoorFinishOther);
                f.LoadPreviousXmlRb(tbCeilingFinish, rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishOther, rbCeilingFinishWhite, rbCeilingFinishMirrorStainlessSteel);
                f.LoadPreviousXmlRb(null, rbBumpRailNo, rbBumpRailYes);
                f.LoadPreviousXmlRb(null, rbFalseCeilingNo, rbFalseCeilingYes);
                f.LoadPreviousXmlRb(tbFrontWall, rbFrontWallBrushedStainlessSteel, tbFrontWallOther);
                f.LoadPreviousXmlRb(tbMirror, rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther);
                f.LoadPreviousXmlRb(tbHandrail, rbHandrailBrushedStainlessSTeel, rbHandrailOther);
                f.LoadPreviousXmlRb(tbSideWall, rbSideWallBrushedStainlessSteel, rbSideWallOther);
                f.LoadPreviousXmlRb(tbRearWall, rbRearWallBrushedStainlessSteel, rbRearWallOther);
                f.LoadPreviousXmlRb(null, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes);
                f.LoadPreviousXmlRb(tbCOPFinish, rbCOPFinishOther, rbCOPFinishSatinStainlessSteel);
                f.LoadPreviousXmlRb(null, rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite);
                f.LoadPreviousXmlRb(null, rbPositionIndicatorTypeFlushMount, rbPositionIndicatorTypeSurfaceMount);
                f.LoadPreviousXmlRb(null, rbExclusiveServiceNo, rbExclusiveServiceYes);
                f.LoadPreviousXmlRb(null, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes);
                f.LoadPreviousXmlRb(null, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes);
                f.LoadPreviousXmlRb(null, rbGPOInCarNo, rbGPOInCarYes);
                f.LoadPreviousXmlRb(null, rbVoiceAnnunciationNo, rbVoiceAnnunciationYes);
                f.LoadPreviousXmlRb(tbFacePlateMaterial, rbFacePlateMaterialOther, rbFacePlateMaterialSatinStainlessSteel);
            }
            catch (Exception)
            {
                return;
            }
        }

        private void tbNumofCarEntrances_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (int.Parse(tbNumofCarEntrances.Text) > 2)
                {
                    rbRearDoorKeySwitchYes.Checked = true;
                }
                else
                {
                    rbRearDoorKeySwitchNo.Checked = true;
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        private void buttonEUR_Click(object sender, EventArgs e)
        {
            // QuoteInfo3 nF = new QuoteInfo3();

            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 
            f.WordData("AE105", tbfname.Text); //first name
            f.WordData("AE106", tblname.Text);//last name
            f.WordData("AE107", tbphone.Text);//phone number
            f.WordData("AE108", tbAddress1.Text);//address 1
            f.WordData("AE109", tbAddress2.Text);//address 2
            f.WordData("AE110", tbAddress3.Text);//address 3
            f.WordData("AE111", f.RadioButtonHandeler(null, rbSL, rbWittur, rbSumasa)); //supplier
            f.WordData("AE114", tbLiftNumbers.Text);//lift number
            f.WordData("AE115", tbTypeofLift.Text);//type of lift
            f.WordData("AE215", "Full Collective"); //control type, not changable 
            f.WordData("AE116", f.RadioButtonHandeler(null, rbIndependentServiceYes, rbIndependentServiceNo));//independent service
            f.WordData("AE117", f.RadioButtonHandeler(null, rbLoadWeighingNo, rbLoadWeighingYes));//load weighing
            f.WordData("AE118", f.RadioButtonHandeler(tbControlerLocation, rbControlerLoactionTopLanding, rbControlerlocationShaft, rbControlerLocationBottomLanding, rbControlerLocationOther));//controler location
            f.WordData("AE120", f.RadioButtonHandeler(null, rbFireServiceNo, rbFireServiceYes));//fire service
            f.WordData("AE123", f.MeasureStringChecker(tbShaftWidth.Text, "mm"));// shaft width
            f.WordData("AE124", f.MeasureStringChecker(tbShaftDepth.Text, "mm"));//shaft depth
            f.WordData("AE125", f.MeasureStringChecker(tbPitDepth.Text, "mm"));//pit depth
            f.WordData("AE126", f.MeasureStringChecker(tbHeadroom.Text, "mm"));//headroom
            f.WordData("AE127", f.MeasureStringChecker(tbTravel.Text, "mm"));//travel
            f.WordData("AE128", tbNumofLandings.Text);// number of landings
            f.WordData("AE129", tbNumofLandingDoors.Text);//number of landing doors 
            f.WordData("AE130", f.RadioButtonHandeler(tbStructureShaft, rbStructureShaftConcrete, rbStructureShaftOther)); //structure shaft 
            f.WordData("AE132", f.RadioButtonHandeler(null, rbTrimmerBeamsNo, rbTrimmerBeamsYes));//trimmer beams
            f.WordData("AE133", f.RadioButtonHandeler(null, rbFalseFloorYes, rbFalseFloorNo));//false floor
            f.WordData("AE134", f.MeasureStringChecker(tbLoad.Text, "kg")); //load
            f.WordData("AE135", f.MeasureStringChecker(tbSpeed.Text, "mps"));//speed
            f.WordData("AE136", f.MeasureStringChecker(tbwidth.Text, "mm")); // width
            f.WordData("AE137", f.MeasureStringChecker(tbDepth.Text, "mm"));//depth
            f.WordData("AE138", f.MeasureStringChecker(tbHeight.Text, "mm"));//height
            f.WordData("AE139", f.MeasureStringChecker(tbLiftRating.Text, "passengers"));//classification rating
            f.WordData("AE113", f.MeasureStringChecker(tbLiftRating.Text, "passenger"));//classification rating
            f.WordData("AE140", tbNumofCarEntrances.Text);//number of car entraces
            if (tbLiftCarNotes.Text != "")
            {
                f.WordData("AE142", "NOTE: " + tbLiftCarNotes.Text);//notes
            }
            else
            {
                f.WordData("AE142", "");//notes
            }
            f.WordData("AE143", f.MeasureStringChecker(tbDoorWidth.Text, "mm"));//door width 
            f.WordData("AE144", f.MeasureStringChecker(tbDoorHeight.Text, "mm")); //door height 
            f.WordData("AE147", f.RadioButtonHandeler(tbLandingDoorFinish, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel));//landing door finish
            f.WordData("AE148", f.RadioButtonHandeler(null, rbDoorTypeCentreOpening, rbDoorTypeSideOpening));//door type
            f.WordData("AE150", f.RadioButtonHandeler(tbDoorTracks, rbDoorTracksAnodisedAluminium, rbDoorTracksOther));// door tracks 
            f.WordData("AE151", f.RadioButtonHandeler(null, rbAdvancedOpeningNo, rbAdvancedOpeningYes));//advanced opening
            f.WordData("AE152", f.RadioButtonHandeler(null, rbDoorNudgingNo, rbDoorNudgingYes));//door nudging 
            f.WordData("AE155", f.RadioButtonHandeler(tbCarDoorFinish, rbCarDoorFinishOther, rbCarDoorFInishBrushedStainlessSteel));//car door finish
            f.WordData("AE156", f.RadioButtonHandeler(tbCeilingFinish, rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishWhite, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
            f.WordData("AE157", f.RadioButtonHandeler(null, rbFalseCeilingNo, rbFalseCeilingYes));//false ceiling
            f.WordData("AE158", f.RadioButtonHandeler(null, rbBumpRailYes, rbBumpRailNo));//bump rail
            f.WordData("AE159", tbFloorFinish.Text);//floor
            f.WordData("AE160", f.RadioButtonHandeler(tbFrontWall, rbFrontWallBrushedStainlessSteel, tbFrontWallOther));//front wall
            f.WordData("AE161", f.RadioButtonHandeler(tbMirror, rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther));//mirror
            f.WordData("AE162", f.RadioButtonHandeler(tbHandrail, rbHandrailBrushedStainlessSTeel, rbHandrailOther));//handrail
            f.WordData("AE163", @"Natural & Mechanical");// ventelation fan
            f.WordData("AE164", f.RadioButtonHandeler(tbSideWall, rbSideWallBrushedStainlessSteel, rbSideWallOther)); //side wall 
            f.WordData("AE165", tbNumOfLEDLights.Text + " LED Lights"); // lighting 
            f.WordData("AE167", f.RadioButtonHandeler(null, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes)); // protective blankets 
            f.WordData("AE216", f.RadioButtonHandeler(tbRearWall, rbRearWallOther, rbRearWallBrushedStainlessSteel)); //  rear wall
            f.WordData("AE168", tbNumberOfCOPS.Text); // number of COPS
            f.WordData("AE169", tbMainCOPLocation.Text);// main COP location
            f.WordData("AE170", tbAuxCOPLocation.Text);//aux cop location
            f.WordData("AE171", tbDesignations.Text); // designations 
            f.WordData("AE191", tbKeyswitchLocation.Text); //keyt switch location
            f.WordData("AE172", f.RadioButtonHandeler(tbCOPFinish, rbCOPFinishSatinStainlessSteel));// COP finish
            f.WordData("AE173", "Dual illumination buttons with gong");// button type 
            f.WordData("AE174", f.RadioButtonHandeler(null, rbLEDColourRed, rbLEDColourBlue, rbLEDColourWhite));// LCD colour
            f.WordData("AE183", f.RadioButtonHandeler(null, rbExclusiveServiceNo, rbExclusiveServiceYes));//exclusive service 
            f.WordData("AE184", f.RadioButtonHandeler(null, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes));// rear door kew switch 
            f.WordData("AE186", f.RadioButtonHandeler(null, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes));//security key switch 
            f.WordData("AE187", f.RadioButtonHandeler(null, rbGPOInCarNo, rbGPOInCarYes));//GPO in car
            f.WordData("AE189", f.RadioButtonHandeler(null, rbVoiceAnnunciationNo, rbVoiceAnnunciationYes));//voice annunciation 
            f.WordData("AE190", f.RadioButtonHandeler(null, rbPositionIndicatorTypeSurfaceMount, rbPositionIndicatorTypeFlushMount));// position indicaor type 
            f.WordData("AE192", f.RadioButtonHandeler(tbFacePlateMaterial, rbFacePlateMaterialSatinStainlessSteel, rbFacePlateMaterialOther));//face plate material 
            f.WordData("AE193", "Dual illumination buttons with gong");//button type
            f.WordData("AE209", f.RadioButtonHandeler(null, rbEmergencyLoweringSystemYes, rbEmergencyLoweringSystemNo));// emergency lowering system 
            f.WordData("AE210", f.RadioButtonHandeler(null, rbOutofServiceNo, rbOutofServiceYes));//out of service 
            f.WordData("AE178", f.CheckboxTrueToYes(f.cbMainSecurity));//security cabiling only 

            f.SaveTbToXML(tbAuxCOPLocation, tbCOPFinish, tbDesignations, tbKeyswitchLocation, tbMainCOPLocation, tbNumberOfCOPS,
                tbCarDoorFinish, tbCeilingFinish, tbFloorFinish, tbFrontWall, tbHandrail, tbMirror, tbNumOfLEDLights, tbRearWall, tbSideWall,
                tbDoorHeight, tbDoorTracks, tbDoorWidth, tbLandingDoorFinish, tbDepth, tbHeight, tbLiftCarNotes, tbLiftRating, tbLoad, tbNumofCarEntrances, tbSpeed,
                tbwidth, tbHeadroom, tbNumofLandingDoors, tbNumofLandings, tbPitDepth, tbTypeofLift, tbLiftNumbers,
                tbShaftDepth, tbShaftWidth, tbStructureShaft, tbTravel, tbControlerLocation, tbfname, tblname, tbphone, tbAddress1, tbAddress2, tbAddress3);

            f.SaveRbToXML(rbCOPFinishOther, rbCOPFinishSatinStainlessSteel, rbExclusiveServiceNo, rbExclusiveServiceYes,
                 rbGPOInCarNo, rbGPOInCarYes, rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite, rbPositionIndicatorTypeFlushMount,
                rbPositionIndicatorTypeSurfaceMount, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes,
                rbVoiceAnnunciationNo, rbVoiceAnnunciationYes, tbFrontWallOther, rbBumpRailNo, rbBumpRailYes, rbCarDoorFInishBrushedStainlessSteel, rbCarDoorFinishOther,
                rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther, rbCeilingFinishWhite, rbFalseCeilingNo,
                rbFalseCeilingYes, rbFrontWallBrushedStainlessSteel, rbHandrailBrushedStainlessSTeel, rbHandrailOther, rbMirrorFullSize, rbMirrorHalfSize,
                rbMirrorOther, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes, rbRearWallBrushedStainlessSteel, rbRearWallOther,
                rbSideWallBrushedStainlessSteel, rbSideWallOther, rbAdvancedOpeningNo, rbAdvancedOpeningYes, rbDoorNudgingNo, rbDoorNudgingYes,
                rbDoorTracksAnodisedAluminium, rbDoorTracksOther, rbDoorTypeCentreOpening, rbDoorTypeSideOpening, rbLandingDoorFinishOther,
                rbLandingDoorFinishStainlessSteel, rbFalseFloorNo, rbFalseFloorYes, rbStructureShaftConcrete, rbStructureShaftOther, rbTrimmerBeamsNo,
                rbTrimmerBeamsYes, rbControlerLoactionTopLanding, rbControlerLocationBottomLanding, rbControlerLocationOther, rbControlerlocationShaft,
                rbFireServiceNo, rbFireServiceYes, rbIndependentServiceNo, rbIndependentServiceYes, rbLoadWeighingNo, rbLoadWeighingYes,
                rbControlerLoactionTopLanding, rbControlerLocationBottomLanding, rbControlerLocationOther, rbControlerlocationShaft,
                rbFireServiceNo, rbFireServiceYes, rbIndependentServiceNo, rbIndependentServiceYes, rbLoadWeighingNo, rbLoadWeighingYes,
                rbSL, rbSumasa, rbWittur);

            this.Enabled = false;
            f.QuestionsComplete();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }

    }
}