using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    #region Form Setup Methods
    public partial class Pin_Dif_Exp : Form
    {
        Pin_Dif_Calc f = Application.OpenForms.OfType<Pin_Dif_Calc>().Single();
        Button[] pageButtons; // = new Button[12];
        Panel[] infoPages;
        int pageTracker = 1;
        int activePage = 0;
        public Pin_Dif_Exp()
        {
            InitializeComponent();
        }

        private void Pin_Dif_Exp_Load(object sender, EventArgs e)
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
            PageButtonSetup();
            PanelSetup();
        }

        private void PageButtonSetup()
        {
            pageButtons = new Button[] { btPanel1, btPanel2, btPanel3, btPanel4, btPanel5, btPanel6, btPanel7, btPanel8, btPanel9, btPanel10, btPanel11, btPanel12 };

            btNewPanel.Location = btPanel2.Location;
            foreach (Button bt in pageButtons)
            {
                bt.Enabled = false;
                bt.Visible = false;
            }
            btPanel1.Enabled = true;
            btPanel1.Visible = true;
            btPanel1.ForeColor = System.Drawing.Color.Blue;
        }

        private void PanelSetup()
        {
            infoPages = new Panel[] { panelLift1, panelLift2, panelLift3, panelLift4, panelLift5, panelLift6, panelLift7, panelLift8, panelLift9, panelLift10, panelLift11, panelLift12 };

            foreach (Panel p in infoPages)
            {
                p.Location = panelLift1.Location;
                p.Visible = false;
                p.Enabled = false;
            }
            panelLift1.Visible = true;
            panelLift1.Enabled = true;
        }
        #endregion

        #region Loading Old Data Methods
        private void PullInfo()
        {
            try
            {
                f.LoadPreviousXmlTb(tbfname, tblname, tbphone, tbAddress1, tbAddress2, tbAddress3, tbLiftNumbers, tbTypeofLift,
                    tbShaftDepth, tbShaftWidth, tbPitDepth, tbHeadroom, tbTravel, tbNumofLandingDoors, tbNumofLandings,
                    tbNumberOfCOPS, tbMainCOPLocation, tbAuxCOPLocation, tbKeyswitchLocation, tbDesignations, tbNumOfLEDLights, tbFloorFinish,
                    tbDoorWidth, tbDoorHeight, tbLoad, tbSpeed, tbwidth, tbDepth, tbHeight, tbLiftRating, tbNumofCarEntrances, tbLiftCarNotes
                    );

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

        #endregion

        #region Data Formatting Methods
        private void tbNumofCarEntrances_TextChanged(object sender, EventArgs e)
        {
            RearDoorChecker(tbNumofCarEntrances, rbRearDoorKeySwitchYes, rbRearDoorKeySwitchNo);
        }

        private void RearDoorChecker(TextBox carEntrance, RadioButton rbYes, RadioButton rbNo)
        {
            try
            {
                if (int.Parse(carEntrance.Text) > 2)
                {
                    rbYes.Checked = true;
                }
                else
                {
                    rbNo.Checked = true;
                }
            }
            catch (Exception)
            {
                return;
            }
        }
        #endregion

        #region Save Data Methods
        private void buttonEUR_Click(object sender, EventArgs e)
        {
            SaveData();
        }

        private void SaveData()
        {
            // QuoteInfo3 nF = new QuoteInfo3();
            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 

            #region Page 1 Word Export
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
            #endregion

            #region Page 1 Saving
            f.SaveTbToXML(tbAuxCOPLocation, tbCOPFinish, tbDesignations, tbKeyswitchLocation, tbMainCOPLocation, tbNumberOfCOPS,
                tbCarDoorFinish, tbCeilingFinish, tbFloorFinish, tbFrontWall, tbHandrail, tbMirror, tbNumOfLEDLights, tbRearWall, tbSideWall,
                tbDoorHeight, tbDoorTracks, tbDoorWidth, tbLandingDoorFinish, tbDepth, tbHeight, tbLiftCarNotes, tbLiftRating, tbLoad, tbNumofCarEntrances, tbSpeed,
                tbwidth, tbHeadroom, tbNumofLandingDoors, tbNumofLandings, tbPitDepth, tbTypeofLift, tbLiftNumbers,
                tbShaftDepth, tbShaftWidth, tbStructureShaft, tbTravel, tbControlerLocation, tbfname, tblname, tbphone, tbAddress1, tbAddress2, tbAddress3
                );

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
                rbSL, rbSumasa, rbWittur
                );
            #endregion
            #region Page 2 Saving
            f.SaveTbToXML( tb2AuxCOPLocation, tb2CarDepth, tb2CarDoorFinishText,
                tb2CarHeight, tb2CarLoad, tb2CarWidth, tb2CeilingFinishText, tb2ControlerLocationText, tb2COPFinishText, tb2Designations,
                tb2DoorHeight, tb2DoorTracksText, tb2DoorWidth, tb2FacePlateMaterialText,  tb2FloorFinish, tb2FrontWallText,
                tb2HandrailText, tb2Headroom, tb2KeyswitchLocation, tb2LandingDoorFinishText,  tb2LiftNumbers, tb2LiftRating,
                tb2MainCOPLocation, tb2MirrorText, tb2Note, tb2NumberOfCarEntrances, tb2NumberOfCOPs, tb2NumberOfLAndingDoors,
                tb2NumberOfLandings, tb2NumberofLEDLights,  tb2PitDepth, tb2RearWallText, tb2ShaftDepth, tb2ShaftWidth,
                tb2SideWallText, tb2Speed, tb2StructureShaftText, tb2Travel, tb2TypeOfLift
                );

            f.SaveRbToXML(rb2AdvancedOpeningNo, rb2AdvancedOpeningYes, rb2BumpRailNo, rb2BumpRailYes, rb2CarDoorFinishOther,
                rb2CarDoorFinishStainlessSteel, rb2CeilingFinishMirrorStainlessSteel, rb2CeilingFinishOther, rb2CeilingFinishStainlessSteel,
                rb2CeilingFinishWhite, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther, rb2ControlerLocationShaft,
                rb2ControlerLocationTopLanding, rb2COPFinishOther, rb2COPFinishStainlessSTeell, rb2DoorNudgingNo, rb2DoorNudgingYes,
                rb2DoorTracksAluminium, rb2DoorTracksOther, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening,
                rb2EmergemncyLoweringSystemNo, rb2EmergencyLoweringSystemYes, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes,
                rb2FacePlateMaterialOther, rb2FacePlateMaterialStainlessSteel, rb2FalseCeilingNo, rb2FalseCeilingYes, rb2FalseFloorNo,
                rb2FalseFloorYes, rb2FireSErviceNo, rb2FireSErviceYes, rb2FrontWallOther, rb2FrontWallStainlessSteel, rb2GPOInCarNo,
                rb2GPOInCarYes, rb2HandrailOther, rb2HandRailStainlessSteel, rb2IndependentServiceNo, rb2IndependentServiceYes,
                rb2LandingDoorFinishOther, rb2LandingDoorFinishStainlessSteel, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite,
                rb2LoadWeighingNo, rb2LoadWeighingYes, rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther, rb2OutOfServiceNo,
                rb2OutOfServiceYes, rb2PositionIndicatorTypeFlushMount, rb2PositionIndicatorTypeSurfaceMount, rb2ProtectiveBlanketsNo,
                rb2ProtectiveBlanketsYes, rb2RearDoorKeySwitchNo, rb2RearDoorKeySwitchYes, rb2RearWallOther, rb2RearWallStainlessSteel,
                rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes, rb2SideWallOther, rb2SideWallStainlessSteel, rb2StructureShaftConcrete,
                rb2StructureShaftOther,    rb2TrimmerBeamsNo, rb2TrimmerBeamsYes,
                rb2VoiceAnnunciationNo, rb2VoiceAnnunciationYes
                );
            #endregion
            #region Page 3 Saving
            f.SaveTbToXML(   tb3AuxCOPLocation, tb3CarDepth, tb3CarDoorFinishText,
                tb3CarHeight, tb3CarNote, tb3CarWidth, tb3CEilingFinishText, tb3ControlerLocationText, tb3COPFinishText, tb3Designations,
                tb3DoorHeight, tb3DoorTracksText, tb3DoorWidth, tb3FacePlaterMaterialText,  tb3FloorFinish, tb3FrontWallText,
                tb3HandrailText, tb3HeadRoom, tb3KeyswitchLocation, tb3LandingDoorFinishText,  tb3LiftNumbers, tb3LiftRating,
                tb3Load, tb3MainCOPLocation, tb3MirrorText, tb3NumberOfCarEntrances, tb3NumberOfCOPs, tb3NumberOfLandingDoors,
                tb3NumberOfLandings, tb3NumberOfLEDLights,  tb3PitDepth, tb3RearWallText, tb3ShaftDepth, tb3ShaftWidth,
                tb3SideWallText, tb3Speed, tb3StructureShaftText, tb3Travel, tb3TypeOfLift
                );

            f.SaveRbToXML(rb3AdvancedOpeningNo, rb3AdvancedOpeningYes, rb3BumpRailNo, rb3BumpRailYes, rb3CarDoorFinishOther,
                rb3CarDoorFinishStainlessSteel, rb3CeilingFinishOther, rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite,
                rb3ControleRLocationBottomLanding, rb3ControlerLocationOther, rb3ControlerLocationShaft, rb3ControlerLocationTopLanding,
                rb3COPFinishOther, rb3COPFinishStainlessSteel, rb3DoorNudgingNo, rb3DoorNudgingYes, rb3DoorTracksAluminium,
                rb3DoorTracksOther, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening, rb3EmergencyLoweringSystemNo,
                rb3EmergencyLoweringSystemYes, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes, rb3FacePlateMaterialOther,
                rb3FacePlateMaterialStainlessSteel, rb3FalseCeilingNo, rb3FalseCeilingYes, rb3FalseFloorNo, rb3FalseFloorYes,
                rb3FireServiceYes, rb3FireServieNo, rb3FrontWallOther, rb3FrontWallStainlessSteel, rb3GPOInCarNo, rb3GPOInCarYes,
                rb3HandrailOther, rb3HandrailStainlessSteel, rb3IndependentServiceNo, rb3IndependentServiceYes, rb3landingDoorFinishOther,
                rb3LandingDoorFinishStainlessSteel, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite, rb3LoadWeighingNo,
                rb3LoadWeighingYes, rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther, rb3MirrorStainlessSteel, rb3OutOfServiceNo,
                rb3OutOfSErviceYes, rb3PositionIndicatorTypeFlushMount, rb3PositionIndicatorTypeSurfaceMount, rb3ProtectiveBlanketsNo,
                rb3ProtectiveBlanketsYes, rb3RearDoorKeySwitchNo, rb3RearDoorKeySwitchYes, rb3RearWallOther, rb3RearWallStainlessSteel,
                rb3SecurityKeySwitchNo, rb3SecurityKeySwitchYes, rb3SideWallOther, rb3SideWallStainlessSteel, rb3StructureShaftConcrete,
                rb3StructureShaftOther,    rb3TrimmerBeamsNo, rb3TrimmerBeamsYes,
                rb3VoiceAnnunciationNo, rb3VoiceAnnunciationYes
                );
            #endregion
            #region Page 4 Saving
            f.SaveTbToXML(   tb4AuxCOPLocation, tb4CarDepth, tb4CarDoorFinish,
                tb4CarHeight, tb4CarNote, tb4CarWidth, tb4CeilingFinishText, tb4ControlerLocationText, tb4COPFinishText, tb4Designations,
                tb4DoorHeight, tb4DoorTracksText, tb4DoorWidth, tb4FacePlateMaterialText,  tb4FloorFinish, tb4FrontWallText,
                tb4HandrailText, tb4Headroom, tb4KeyswitchLocations, tb4LandingDoorFinishText,  tb4LiftNumbers, tb4LiftRating,
                tb4Load, tb4MainCOPLocation, tb4MirrorText, tb4NumberOfCarEntrances, tb4NumberOfCOPs, tb4NumberOfLandingDoors,
                tb4NumberOfLandings, tb4NumbeROfLEDLights, tb4PitDepth, tb4RearWallText, tb4ShaftDepth, tb4ShaftWidth, tb4SideWallText,
                tb4Speed, tb4StructureShaftText, tb4Travel, tb4TypeOfLift
                );

            f.SaveRbToXML(rb4AdvancedOpeningNo, rb4AdvancedOpeningYes, rb4BumpRailNo, rb4BumpRailYes, rb4CarDoorFinishOther,
                rb4CarDoorFinishStainlessSteel, rb4CeilingFinishMirrorStainlessSteel, rb4CeilingFinishOther, rb4CeilingFinishStainlessSteel,
                rb4CeilingFinishWhite, rb4ControelrLocationTopLanding, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther,
                rb4ControlerLocationShaft, rb4COPFinishOther, rb4COPFinishStainlessSteel, rb4DoorNudgingNo, rb4DoorNudgingYes,
                rb4DoorTracksAluminium, rb4DoorTracksOther, rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening, rb4EmergencyLoweringSystemNo,
                rb4EmergencyLoweringSystemYes, rb4ExclusiveServiceNo, rb4ExclusiveServiceYes, rb4FacePlateMaterialOther,
                rb4FacePlateMaterialStainlessSteel, rb4FalseCeilingNO, rb4FalseCeilingYes, rb4FalseFloorNo, rb4FalseFloorYes, rb4FireSErviceNo,
                rb4FireServiceYes, rb4FrontWallStainlessSteel, rb4FrotnWallOther, rb4GPOInCarNo, rb4GPOInCarYes, rb4HandrailOther,
                rb4HandRailStainlesSteel, rb4IndependentServiceYes, rb4LandingDoorFinishOther, rb4LandingDoorFinishStainlessSteel, rb4LCDColourBlue,
                rb4LCDColourRed, rb4LCDColourWhite, rb4LoadWeighingNo, rb4LoadWeighingYes, rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther,
                rb4OutOfServiceNo, rb4OutOfServiceYes, rb4PositionIndicatorTypeFlushMount, rb4PositionIndicatorTypeSurfaceMount,
                rb4ProtectiveBlanketsYes, rb4ProtetiveBlanketsNo, rb4RearDoorKeySwitchNo, rb4RearDoorKeySwitchYes, rb4RearWallOther,
                rb4RearWallStainlessSteel, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes, rb4SideWallOther, rb4SideWallStainlessSteel,
                rb4StructureShaftConcrete, rb4StructureShaftOther,    rb4TrimmerBeamsNo,
                rb4TrimmerBeamsYes, rb4VoiceAnnunciationNo, rb4VoiceAnnunciationYes
                );
            #endregion
            #region Page 5 Saving
            f.SaveTbToXML(   tb5AuxCOPLocation, tb5CarDepth, tb5CarDoorFinishText,
                tb5CaRHeight, tb5CarNote, tb5CarWidth, tb5CeilingFinishText, tb5ControlerLocationText, tb5COPFinishText, tb5Designations,
                tb5DoorHeight, tb5DoorTRacksText, tb5DoorWidth, tb5FacePlateMaterialText,  tb5FloorFinish, tb5FrontWallText,
                tb5HandrailText, tb5Headroom, tb5KetyswitchLocation, tb5LandingDoorFinishText,  tb5LiftNumbers, tb5LiftRating,
                tb5Load, tb5MainCOPLocation, tb5MirrorText, tb5NumberOfCarEntrances, tb5NumberOfCOPs, tb5NumberOfLandingDoors,
                tb5NumberOfLandings, tb5NumberOIfLEDLights,  tb5PitDepth, tb5RearWallText, tb5ShaftDEpth, tb5ShaftWidth,
                tb5SideWallText, tb5Speed, tb5StructureShaftText, tb5Travel, tb5TypeOfLift
                );

            f.SaveRbToXML(rb5AdvancedOpeningNo, rb5AdvancedOpeningYes, rb5BumpRailNo, rb5BumpRailYes, rb5CarDoorFinishOther,
                rb5CarDoorFinishStainlessSteel, rb5CeilingFinishMirrorStainlessSTeel, rb5CeilingFinishOther, rb5CeilingFinishStainlessSteel,
                rb5CeilingFinishWhite, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther, rb5ControlerLocationShaft,
                rb5ControlerLocationTopLanding, rb5COPFinishOther, rb5COPFinishStainlessSteel, rb5DoorNudgingNo, rb5DoorNudgingYes,
                rb5DoorTracksAluminium, rb5DoorTracksOther, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening, rb5EmergencyLoweringSystemNo,
                rb5EmergencyLoweringSystemYes, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes, rb5FacePlateMAterialOtjer,
                rb5FacePlateMaterialStainlessSteel, rb5FalseCeilingNo, rb5FalseCeilingYes, rb5FalseFloorNo, rb5FalseFloorYes, rb5FireSErviceNo,
                rb5FireServiceYes, rb5FrontWallOther, rb5FrontWallStainlessSteel, rb5GPOInCarNo, rb5GPOInCarYes, rb5HandRailOther,
                rb5HandRailStainlessSteel, rb5IndependentServiceNo, rb5IndependentServiceYes, rb5LAndingDoorFinishOther,
                rb5LandingDoorFinishStainlessSteel, rb5LCDColoiurWHite, rb5LCDColourBlue, rb5LCDColourRed, rb5LoadWeighingNo,
                rb5LoadWeighingYes, rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther, rb5OutOfServiceNo, rb5OutOfServiceYes,
                rb5PositionIndicatorTypeSurfaceMount, rb5PostionIndicatorTypeFlushMount, rb5ProtectiveBlanketsNo, rb5ProtectiveBlanketsYes,
                rb5RearDoorKeySwitchNo, rb5RearDoorKeySwitchYes, rb5RearWallOther, rb5RearWallStainlesSteel, rb5SecurityKeySwitchNo,
                rb5SecurityServiceYes, rb5SideWallOther, rb5SideWallStainlessSTeel, rb5StructureShaftConcrete, rb5StructureShaftOther,
                   rb5TrimmerBeamsNo, rb5TrimmerBeamsYes,
                rb5VoiceAnnunciationNo, rb5VoiceAnnunciationYes
                );
            #endregion
            #region Page 6 Saving
            f.SaveTbToXML(   tb6AuxCOPLocation, tb6CarDepth, tb6CarDoorFinishText,
                tb6CarHeight, tb6CarLoad, tb6CarNote, tb6CarSpeed, tb6CarWidth, tb6CeilingFinishText, tb6ControlerLocationText, tb6COPFinishText,
                tb6Designations, tb6DoorHeight, tb6DoorTracksOther, tb6DoorWidth, tb6FacePlateMaterialText,  tb6FloorFinish,
                tb6FrontWallText, tb6HAndrailText, tb6Headroom, tb6KeySwitchLocation, tb6LandingDoorFinishText,  tb6LiftNumbers,
                tb6LiftRating, tb6MainCOPLocation, tb6MirrorText, tb6NumberOfCarEntrances, tb6NumberOFCOPs, tb6NumberOfLandingDoors,
                tb6NumberOfLandings, tb6NumberOfLEDLights,  tb6PitDepth, tb6RearWallText, tb6ShaftDepth, tb6ShaftWidth,
                tb6SideWallText, tb6StructureShaftText, tb6Travel, tb6TypeOfLift
                );

            f.SaveRbToXML(rb6AdvancedOpeningNo, rb6AdvancedOpeningYes, rb6BumpRailNo, rb6BumpRailYes, rb6CarDoorFinishOther,
                rb6CarDoorFinishStainlessSteel, rb6CeilingFinishMirrorStainlessSteel, rb6CeilingFinishOther, rb6CeilingFinishStainlessSteel,
                rb6CeilingFinishWhite, rb6ControlerLocationBottomLanding, rb6ControlerLocationOther, rb6ControlerLocationShaft,
                rb6ControlerLocationTopLanding, rb6COPFinishOther, rb6COPFinishStainlessSteel, rb6DoorNudgingNo, rb6DoorNudgingYes,
                rb6DoorTracksAluminium, rb6DoorTracksOther, rb6DoorTypeCentreOpening, rb6DoorTypeSideOpening, rb6EmergencyLoweringSystemNo,
                rb6EmergencyLoweringSystemYes, rb6ExclusiveServiceNo, rb6ExclusiveServiceYes, rb6FacePlateMaterialOther,
                rb6FacvePlateMaterialStainlessSteel, rb6FalseCeilingNo, rb6FalseCeilingYes, rb6FalseFloorNo, rb6FalseFloorYes, rb6FireServiceNo,
                rb6FireServiceYes, rb6FrontWallOther, rb6FrontWallStainlessSteel, rb6GPOInCarNo, rb6GPOInCarYes, rb6HandrailOther,
                rb6HandrailStainlessSteel, rb6IndependentNo, rb6IndependentServiceYes, rb6LandingDoorFinishOther, rb6LandingDoorFinishStainlessSteel,
                rb6LCDColourBlue, rb6LCDColourRed, rb6LCDColourWhite, rb6LoadWeighingNo, rb6LoadWeighingYes, rb6MirrorFullSize, rb6MirrorHalfSize,
                rb6MirrorOther, rb6OutOfServiceNo, rb6OutOfServiceYes, rb6PositionIndicatorTypeFlushMount, rb6PositionIndicatorTypeSurfaceMount,
                rb6ProtectiveBlanketsNo, rb6ProtectiveBlanketsYes, rb6RearDoorKeySwitchNo, rb6RearDoorKeySwitchYes, rb6RearWallOther,
                rb6RearWallStainlessSteel, rb6SecurityKeySwitchNo, rb6SecurityKeySwitchYes, rb6SideWallOther, rb6SideWallStainlessSteel,
                rb6StructureShaftCOncrete, rb6StructureShaftOther,    rb6TrimmerBeamsNo,
                rb6TrimmerBeamsYes, rb6VoiceAnnunciationNo, rb6VoiceAnnunciationYes
                );
            #endregion
            #region Page 7 Saving
            f.SaveTbToXML(   tb7AuzCOPLocation, tb7CarDepth, tb7CarDoorFinishText,
                tb7CarHeight, tb7CarLoad, tb7CarNotes, tb7CarSpeed, tb7CarWidth, tb7CEilingFinishText, tb7ControlerLocationText, tb7COPFinishText,
                tb7Designations, tb7DoorHeight, tb7DoorTracksText, tb7DoorWidth, tb7FacePlateMaterialText,  tb7FloorFinish, tb7FrontWallText,
                tb7HandrailText, tb7HeadRoom, tb7KeyswitchLocation, tb7LandingDoorFinishText,  tb7LiftNumbers, tb7LiftRating,
                tb7MainCOPLocation, tb7MirrorText, tb7NumberOfCarEntrances, tb7NumberOfCOPs, tb7NumberOfLandingDoors, tb7NumberOfLandings,
                tb7NumberOfLEDLights,  tb7PitDepth, tb7RearWallText, tb7ShaftDepth, tb7ShaftWidth, tb7SideWallText,
                tb7StructureShaftText, tb7Travel, tb7TypeOfLift
                );

            f.SaveRbToXML(rb7AdvancedOpeningNo, rb7AdvancedOpeningYes, rb7BumpRailNo, rb7BumpRailYes, rb7CarDoorFinishOther,
                rb7CarDoorFinishStainlessSteel, rb7CeilingFinishMirrorStainlessSteel, rb7CeilingFinishOther, rb7CeilingFinishStainlessSteel,
                rb7CEilingFinishWhite, rb7ControlerLocationBottomLAnding, rb7ControlerLocationOther, rb7ControlerLocationShaft,
                rb7ControlerLocationTopLanding, rb7COPFinishOther, rb7COPFinishStainlessSteel, rb7DoorNudgingNo, rb7DoorNudgingYes,
                rb7DoorTracksAluminium, rb7DoorTracksOther, rb7DoorTypeCentreOpening, rb7DoorTypeSideOpening, rb7EmergencyLoweringSystemNo,
                rb7EmergencyLoweringSystemYes, rb7ExclusiveServiceNo, rb7ExclusiveServiceYes, rb7FacePlateMaterialOther,
                rb7FacePlateMaterialStainlessSteel, rb7FalseCeilingNo, rb7FalseCeilingYes, rb7FalseFloorNo, rb7FalseFloorYes, rb7FireServiceNo,
                rb7FireSErviceYes, rb7FrontWallOther, rb7FrontWallStainlessSteel, rb7GPOInCarNo, rb7GPOInCarYes, rb7HandrailOther,
                rb7HandrailStainlessSteel, rb7IndependentServiceNo, rb7IndpendentServiceYes, rb7LandingDoorFinishOther,
                rb7LandingDoorFinishStainlessSteel, rb7LCDColourBlue, rb7LCDColourRed, rb7LCDColourWhite, rb7LoadWeighingNo, rb7LoadWeighingYes,
                rb7MirrorFullSize, rb7MirrorHalfSize, rb7MirrorOther, rb7OutOfSErviceNo, rb7OutOfServiceYes, rb7PositionIndicatorTypeFlushMount,
                rb7PositionIndicatorTypeSurfaceMount, rb7ProtctiveBlanketsNo, rb7ProtectiveBlanketsYes, rb7RearDoorKeySwitchNo,
                rb7RearDoorKeySwitchYes, rb7RearWallOther, rb7RearWallStainlessSteel, rb7SecurityKeySwitchNo, rb7SecurityKeySwitchYes,
                rb7SideWallOther, rb7SideWallStainlessSteel, rb7StructureShaftConcrete, rb7StructureShaftOther,  
                 rb7TrimmerBeamsNo, rb7TrimmerBeamsYes, rb7VoiceAnnunciationYes, rb7VoiceAnunciationNo
                );
            #endregion
            #region Page 8 Saving
            f.SaveRbToXML(rb8AdvancedOpeningYes, rb8AdvncedOpeningNo, rb8BumpRailNo, rb8BumpRailYes, rb8CarDoorFinishOther,
                rb8CarDoorFinishStainlessSteel, rb8CeilingFinishMirrorStainlessSTeel, rb8CeilingFinishOther, rb8CeilingFinishStainlessSteel,
                rb8CeilingFinishWhite, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther, rb8ControlerLocationShaft,
                rb8ControlerLocationTopLanding, rb8COPFinishOther, rb8COPFinishStainlessSteel, rb8DoorNudgingYes, rb8DoorTracksAluminium,
                rb8DoorTracksOther, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening, rb8EmergencyLoweringSystemNo,
                rb8EmergencyLoweringSystemYes, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes, rb8FacePlateMaterialOther,
                rb8FacePlateMaterialStainlessSteel, rb8FalseCeilingNo, rb8FalseCeilingYes, rb8FalseFloorNo, rb8FalseFloorYes, rb8FireSErviceNo,
                rb8FireServiceYes, rb8FrontWallOther, rb8FrontWallStainlessSteel, rb8GPOInCarNo, rb8GPOInCarYes, rb8HandrailOther,
                rb8HandRailStainlessSteel, rb8IndependentServiceNo, rb8IndependentServiceYes, rb8LandingDoorFinishOther, rb8LCDColourBlue,
                rb8LCDColourRed, rb8LCDColourWhite, rb8LoadWeighingNo, rb8LoadWeighingYes, rb8MirrorFullSize, rb8MirrorHalfSize,
                rb8MirrorOther, rb8OutOFServiceNo, rb8OutOfSErviceYes, rb8PositionIndicatorTypeFlushMount, rb8PositionIndicatorTypeSurfaceMount,
                rb8ProtectiveBlanketsNo, rb8ProtectiveBlanketsYes, rb8RearDoorKeySwitchNo, rb8RearDoorKeySwitchYes, rb8RearWallOther,
                rb8RearWallStainlessSteel, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes, rb8SideWallOther, rb8SideWallStainlessSteel,
                rb8StructureShaftConcrete, rb8StructureShaftOther,    rb8TimmemrBeamsYes,
                rb8TrimmerBeamsNo, rb8VoiceAnnunicationNo, rb8VoiceAnnunicationYes, tb8DoorNudgingNo, tb8LandingDoorFinishStainlessSteel
                );

            f.SaveTbToXML(   tb8AuxCOPLocation, tb8CarDEpth, tb8CarDoorFinishText,
                tb8CarHeight, tb8CarWidth, tb8CeilingFinishText, tb8ControlerLocationText, tb8COPFinishText, tb8Desiginations, tb8DoorHeight,
                tb8DoorTracksText, tb8DoorWidth, tb8FacePlateMaterialText,  tb8FloorFinish, tb8FrontWallText, tb8HandrailText,
                tb8Headroom, tb8KeyswitchLocations, tb8LandingDoorFinishText,  tb8LiftCarNotes,
                tb8LiftNumbers, tb8LiftRating, tb8Load, tb8MainCOPLocation, tb8MirrorText, tb8NumberOfCarEntrances, tb8NumberOfCOPs,
                tb8NumberOfLandingDoors, tb8NumberOfLandings, tb8NumberofLEDLights,  tb8PitDepth, tb8RearWallText,
                tb8ShaftDepth, tb8ShaftWidth, tb8SideWallText, tb8Speed, tb8StructureShaftText, tb8Travel, tb8TypeOfLift
                );
            #endregion
            #region Page 9 Saving
            f.SaveRbToXML(rb9AdvancedOpeningNo, rb9AdvancedOpeningYes, rb9BumpRailNo, rb9BumpRailYes, rb9CarDoorFinishOther,
                rb9CarDoorFinishStainlessSteel, rb9CeilingFinishMirrorStainlessSteel, rb9CeilingFinishOther, rb9CeilingFinishStainlessSteel,
                rb9CeilingFinishWhite, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther, rb9ControlerLocationShaft, rb9ControlrLocationTopLanding,
                rb9COPFinishOther, rb9COPFinishStainlessSteel, rb9DoorNudgingNo, rb9DoorNudgingYes, rb9DoorTracksAluminium, rb9DoorTracksOther,
                rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening, rb9EmergencyLoweringSystemNo, rb9EmergencyLoweringSystemYes, rb9ExclusiveServiceNo,
                rb9ExclusiveServiceYes, rb9FacePlateMaterialOther, rb9FacePlateMaterialStainlessSteel, rb9FalseCeilingNo, rb9FalseCeilingYes, rb9FalseFloorNo,
                rb9FalseFloorYes, rb9FireServiceNo, rb9FireSErviceYes, rb9FrontWallOther, rb9FrontWallStainlessSteel, rb9GPOInCarNo, rb9GPOInCarYes,
                rb9HandrailOther, rb9HandrailStainlessSteel, rb9IndependentServiceNo, rb9IndependentServiceYes, rb9LandingDoorFinishOther,
                rb9LandingDoorFinishStainlessSteel, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite, rb9LoadWeighingNo, rb9LoadWeighingYes,
                rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther, rb9OutOfServiceNo, rb9OutOfServiceYes, rb9PositionIndicatorTypeFlushMount,
                rb9PositionIndicatorTypeSurfaceMount, rb9ProtectiveBlanketsNo, rb9ProtectiveBlanketsYes, rb9RearDoorKeySwitchNo, rb9RearDoorKeySwitchYes,
                rb9RearWallOther, rb9RearWallStainlessSteel, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes, rb9SideWallOther, rb9SideWallStainlessSteel,
                rb9StructureShaftConcrete, rb9StructureShaftOther,    rb9TrimmerBeamsNo, rb9TrimmerBeamsYes,
                rb9VoiceAnnunciationNo, rb9VoiceAnnunciationYes
                );

            f.SaveTbToXML(   tb9AuxCOPLocation, tb9CarDepth, tb9CarDoorFinishText, tb9CarHeight,
                tb9CarNotes, tb9CarWidth, tb9CeilingFinishText, tb9ControlerLocationText, tb9COPFinishText, tb9Designations, tb9DoorHeight, tb9DoorTracksText,
                tb9DoorWidth, tb9FacePlateMaterialText,  tb9FloorFinish, tb9FrontWallText, tb9HandrailTexrt, tb9Headroom, tb9KeyswitchLocation,
                tb9LandingDoorFinishText,  tb9LiftNumbers, tb9LiftRating, tb9Load, tb9MainCOPLocation, tb9MirrorText, tb9NumberOFCarEntraces,
                tb9NumberOfCOPs, tb9NumberOfLandingDoors, tb9NumberOfLandings, tb9NumberOfLEDLights,  tb9PitDepth, tb9RearWallText,
                tb9ShaftDepth, tb9ShaftWidth, tb9SideWallText, tb9Speed, tb9StructureShaftText, tb9Travel, tb9TypeOfLift
                );
            #endregion
            #region Page 10 Saving
            f.SaveTbToXML(   tb10AuxCOPLocation, tb10CarDepth, tb10CarDoorFinishText,
                tb10CarHeight, tb10CarWidth, tb10CEilingFinishText, tb10ControlerLocationText, tb10COPFinishText, tb10Desigination, tb10DoorHeight,
                tb10DoorTracksText, tb10DoorWidth, tb10FacePlateMaterialText,  tb10FloorFinish, tb10FrontWallText, tb10HandrailText,
                tb10Headroom, tb10KeyswitchLocation, tb10LandingDoorFinishText,  tb10LiftCarLoad, tb10LiftCarNotes, tb10LiftNumbers,
                tb10LiftRating, tb10MainCOPLocation, tb10MirrorText, tb10NumberofCarEntrances, tb10NumberOfCOPs, tb10NumberofLandingDoors,
                tb10NumberofLandings, tb10NumberOfLEDLIghts,  tb10PitDepth, tb10RearWallText, tb10ShaftDepth, tb10ShaftWidth,
                tb10SideWallText, tb10Speed, tb10StructureShaftText, tb10Travel, tb10TypeOfLift
                );

            f.SaveRbToXML(rb10AdvancedOpeningNo, rb10AdvancedOpeningYes, rb10BumpRaidYes, rb10BumpRailNo, rb10CarDoorFinishOther,
                rb10CarDoorFinishStainlessSteel, rb10CEilingFinishMirrorStainlessSteel, rb10CEilingFinishOther, rb10CeilingFinishStainlessSteel,
                rb10CeilingFinishWhite, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther, rb10ControlerLocationShaft,
                rb10ControlerLocationTopLanding, rb10COPFinishOther, rb10COPFinishStainlessSteel, rb10DoorNudgingNo, rb10DoorNudgingYes,
                rb10DoorTracksAluminium, rb10DoorTracksOther, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening, rb10EmergencyLoweringSystemNo,
                rb10EmergencyLoweringSystemYes, rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes, rb10FacePlateMaterialOther, rb10FacePlateMaterialStainlessSteel,
                rb10FalseCeilingNo, rb10FalseCeilingYes, rb10FalseFloorNo, rb10FalseFloorYes, rb10FireSERviceNo, rb10FireSErviceYes, rb10FrontWallOther,
                rb10FrontWallStainlessSteel, rb10GPOInCarNo, rb10GPOInCarYes, rb10HandrailOther, rb10HandrailStainlessSteel, rb10IndependentServiceNo,
                rb10IndependentServiceYes, rb10LAndingDoorFinishOtherr, rb10LandingDoorFinishStainlessSteel, rb10LCDColourBlue, rb10LCDColourRed,
                rb10LCDColourWhite, rb10LoadWEighingNo, rb10LoadWeighingYes, rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther, rb10OutOfServiceNo,
                rb10OutOFServiceYes, rb10PositionIndicatorTypeFlushMount, rb10PositionIndicatorTypeSurfaceMount, rb10ProtectiveBlanketNo, rb10ProtectiveBlanketYes,
                rb10RearDoorKeySwitchNo, rb10RearDoorKeySwitchYes, rb10RearWallOther, rb10RearWallStainlessSteel, rb10SecurityKeySwitchNo,
                rb10SecurityKeySwitchYes, rb10SideWallOther, rb10SideWallStainlesSteel, rb10StructureShaftConcrete, rb10StructureShaftOther, 
                  rb10TimmerbeamsYes, rb10TrimmerBeamsNo, rb10VoiceAnnunciationNo, rb10VoiceAnnunciationYes
                );
            #endregion
            #region Page 11 Saving
            f.SaveTbToXML(   tb11AuxCOPLocation, tb11CarDepth, tb11CarDoorFinishText,
                tb11CarHeight, tb11CarWidth, tb11CeilingFinishText, tb11ControlerLocationText, tb11COPFinishText, tb11Designations, tb11DoorHeight,
                tb11DoorTracksText, tb11DoorWidth, tb11FaceplateMaterialText,  tb11FloorFinish, tb11FrontWallText, tb11HandrailText, tb11Headroom,
                tb11KeyswitchLocation, tb11LandingDoorFInishOther,  tb11LiftCarLoad, tb11LiftCarNote, tb11LiftRating, tb11MainCOPLocation,
                tb11MirrorText, tb11NumberofCarEntrances, tb11NumberOfCOPs, tb11NumberOfLandingDoors, tb11NumberOfLandings, tb11NumberOfLEDLights,
                 tb11PitDepth, tb11RearWallText, tb11ShaftDepth, tb11ShaftWidth, tb11SideWallText, tb11Speed, tb11Travel, tb11TypeOfLift,
                rb11LiftNumbers, rb11StructureShaftText
                );

            f.SaveRbToXML(rb11AdvancedOpeningNo, rb11AdvancedOpeningYes, rb11BumpRailNo, tb11ControlerLocationTopLanding, rb11BumpRailYes,
                rb11CarDoorFinishOther, rb11CarDoorFinishStainlessSteel, rb11CeilingFinishMirrorStainlessSteel, rb11CeilingFinishOther, rb11CeilingFinishStainlessSteel,
                rb11CeilingFinishWhite, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther, rb11ControlerLocationShaft, rb11COPFinishOther,
                rb11COPFinishStainlessSteel, rb11DoorNudgingNo, rb11DoorNudgingYes, rb11DoorTracksAluminium, rb11DoorTracksOther, rb11DoorTypeCentreOpening,
                rb11DoorTypeSideOpening, rb11EmergencyLoweringSystemNo, rb11EmergencyLoweringSystemYes, rb11ExclusiveServiceNo, rb11ExclusiveServiceYes,
                rb11FacePlateMaterialOther, rb11FacePlateMaterialStainlessSTeel, rb11FalseCeilingNo, rb11FalseCEilingYes, rb11FalseFloorNo, rb11FalseFloorYes,
                rb11FireServiceNo, rb11FireServiceYes, rb11FrontWallOther, rb11FrontWallStainlessSteel, rb11GPOInCarNo, rb11GPOInCarYes, rb11HandrailOther,
                rb11HandrailStainlessSteel, rb11IndependentSErviceNO, rb11IndependentServiceYes, rb11LandingDoorFinishOther, rb11LandingDoorFinishStainlessSteel,
                rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite, rb11LoadWeighingNo, rb11LoadWeighingYes, rb11MirrorFullSize,
                rb11MirrorHalfSize, rb11MirrorOther, rb11OutOfServiceNo, rb11OutOfSErviceYes, rb11PositionIndicatorTypeFlushMount,
                rb11PositionIndicatorTypeSurfaceMount, rb11ProtectiveBlanketNo, rb11ProtectiveBlanketsYes, rb11RearDoorKeySwitchNo, rb11RearDoorKeySwitchYes,
                rb11RearWallOther, rb11RearWallStainlessSteel, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes, rb11SideWallOther, rb11SideWallStainlessSteel,
                rb11StrructureShaftOther, rb11StructureShaftConcrete,   
                rb11TrimmerBeamNo, rb11TrimmerBeamsYes, rb11VoiceAnnunciationNo, rb11VoiceAnnunciationYes
                );
            #endregion
            #region Page 12 Saving
            f.SaveTbToXML(   tb12AuxCOPLocation, tb12CarDepth, tb12CarDoorFinishText,
                tb12CarHeight, tb12CarLiftRating, tb12CarLoad, tb12CarNumberOfCarEntrances, tb12CarSpeed, tb12CarWidth, tb12CeilingFinishText,
                tb12ControlerLocationText, tb12COPFinishText, tb12Designations, tb12DoorTracksText, tb12FacePlateMaterialText, 
                tb12FloorFinish, tb12FrontWallText, tb12HandrailText, tb12Headroom, tb12KeyswitchLocation, tb12LandingDoorFinishText,
                tb12LandingDoorHeight, tb12LandingDoorWidth,  tb12LiftCarNotes, tb12LiftNumbers, tb12MainCOPLocation, tb12MirrorText,
                tb12NumberOfCOPs, tb12NumberOfLandingDoors, tb12NumberOfLandings, tb12NumberOfLEDLights,  tb12PitDepth,
                 tb12RearWallText, tb12ShaftDepth, tb12ShaftWidth, tb12SideWallText, tb12StructureShaftText, tb12Travel, tb12TypeOfLift
                 );

            f.SaveRbToXML(rb12AdvancedOpeningNo, rb12AdvancedOpeningYes, rb12BumpRailNo, rb12BumpRailYes, rb12CarDoorFinishOther,
                rb12CarDoorFinishStainlessSteel, rb12CeilingFinishMirrorStainlessSteel, rb12CeilingFinishOther, rb12CeilingFinishStainlessSteel,
                rb12CeilingFinishWhite, rb12ControlerLocationBottomLanding, rb12ControlerLocationOther, rb12ControlerLocationShaft,
                rb12ControlerLocationTopLanding, rb12COPFinishOther, rb12COPFinishStainlessSteel, rb12DoorTracksAluminium, rb12DoorTracksOther,
                rb12DoorTypeCentreOpening, rb12DoorTypeSideOpening, rb12EmergencyLoweringSystemNo, rb12EmergencyLoweringSystemYes,
                rb12ExclusiveServiceNo, rb12ExclusiveServiceYes, rb12FacePlateMaterialOther, rb12FacePlateMaterialStainlessSteel, rb12FalseCeilingNo,
                rb12FalseCeilingYes, rb12FalseFloorNo, rb12FalseFloorYes, rb12FireServiceNo, rb12FireServiceYes, rb12FrontWallOther,
                rb12FrontWallStainlessSteel, rb12GPOInCarNo, rb12GPOInCarYes, rb12HandrailOther, rb12HandrailStainlessSTeel, rb12IndependentServiceNo,
                rb12IndependentServiceYes, rb12LandingDoorFinishOther, rb12LandingDoorFinishStainlessSteel, rb12LandingDoorNudgingNo,
                rb12LandingDoorNudgingYes, rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite, rb12LoadWeighingNo, rb12LoadWeighingYes,
                rb12MirrorFullSize, rb12MirrorHalfSize, rb12MirrorOTher, rb12OutOfServiceNo, rb12OutOfServiceYes, rb12PositionIndicatorTypeFlushMount,
                rb12PositionIndicatorTypeSurfaceMount, rb12ProectiveBlanketsYes, rb12ProtectiveBlanketsNo, rb12RearDoorKeySwitchNo, rb12RearDoorKeySwitchYes,
                rb12RearWallOther, rb12RearWallStainlessSteel, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes, rb12SideWallOther, rb12SideWallStainlessSteel,
                rb12StructureShaftConcrete, rb12StructureShaftOther,    rb12TrimmerBeamsNo,
                rb12TrimmerBeamsYes, rb12VoicAnnunciationNo, rb12VoiceAnnuniationYes
                );
            #endregion

            this.Enabled = false;
            f.QuestionsComplete();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            f.QuestionCloseCall(this);
        }
        #endregion

        #region Page Controls
        private void btNewPanel_Click(object sender, EventArgs e)
        {
            NewPage();
        }

        private void NewPage()
        {
            pageButtons[pageTracker].Visible = true;
            pageButtons[pageTracker].Enabled = true;
            pageTracker++;

            if (pageTracker <= 11)
            {
                btNewPanel.Location = pageButtons[pageTracker].Location;
            }
            else
            {
                btNewPanel.Visible = false;
                btNewPanel.Enabled = false;
            }
        }
        private void btPanel1_Click(object sender, EventArgs e)
        {

            OpenInfoPage(0);
        }

        private void btPanel2_Click(object sender, EventArgs e)
        {

            OpenInfoPage(1);
        }

        private void btPanel3_Click(object sender, EventArgs e)
        {

            OpenInfoPage(2);
        }

        private void btPanel4_Click(object sender, EventArgs e)
        {

            OpenInfoPage(3);
        }

        private void btPanel5_Click(object sender, EventArgs e)
        {

            OpenInfoPage(4);
        }

        private void btPanel6_Click(object sender, EventArgs e)
        {

            OpenInfoPage(5);
        }

        private void btPanel7_Click(object sender, EventArgs e)
        {

            OpenInfoPage(6);
        }

        private void btPanel8_Click(object sender, EventArgs e)
        {

            OpenInfoPage(7);
        }

        private void btPanel9_Click(object sender, EventArgs e)
        {

            OpenInfoPage(8);
        }

        private void btPanel10_Click(object sender, EventArgs e)
        {

            OpenInfoPage(9);
        }

        private void btPanel11_Click(object sender, EventArgs e)
        {

            OpenInfoPage(10);
        }

        private void btPanel12_Click(object sender, EventArgs e)
        {

            OpenInfoPage(11);
        }

        private void OpenInfoPage(int pageToOpen)
        {
            if (pageToOpen == activePage)
            {
                return;
            }
            else
            {
                infoPages[pageToOpen].Visible = true;
                infoPages[pageToOpen].Enabled = true;
                pageButtons[pageToOpen].ForeColor = System.Drawing.Color.Blue;
                infoPages[activePage].Visible = false;
                infoPages[activePage].Enabled = false;
                pageButtons[activePage].ForeColor = System.Drawing.Color.Black;
                activePage = pageToOpen;
            }
        }
        #endregion

        #region Unused Methods
        private void textBox64_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton120_CheckedChanged(object sender, EventArgs e)
        {
            //
        }
        private void radioButton590_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton589_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox132_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton181_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton178_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton151_CheckedChanged(object sender, EventArgs e)
        {
            //
        }
        private void textBox276_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton515_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton480_CheckedChanged(object sender, EventArgs e)
        {
            //
        }
        private void radioButton735_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox213_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton443_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox230_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton382_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox229_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox248_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox232_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox247_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton424_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton377_CheckedChanged(object sender, EventArgs e)
        {
            //
        }

        private void radioButton400_CheckedChanged(object sender, EventArgs e)
        {
            //
        }
        #endregion

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            //
        }
    }
}