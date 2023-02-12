using System;
using System.Linq;
using System.Windows.Forms;

namespace SQT
{
    public partial class Pin_Dif_Exp : Form
    {
        #region Vars

        Pin_Dif_Calc f = Application.OpenForms.OfType<Pin_Dif_Calc>().Single();
        Button[] pageButtons; // = new Button[12];
        Panel[] infoPages;
        int pageTracker = 1;
        int activePage = 0;

        bool rearDoorChecker = false;
        bool page2Opened = false;
        bool page3Opened = false;
        bool page4Opened = false;
        bool page5Opened = false;
        bool page6Opened = false;
        bool page7Opened = false;
        bool page8Opened = false;
        bool page9Opened = false;
        bool page10Opened = false;
        bool page11Opened = false;
        bool page12Opened = false;

        #endregion

        #region Form Setup Methods
        public Pin_Dif_Exp()
        {
            InitializeComponent();
        }

        private void Pin_Dif_Exp_Load(object sender, EventArgs e)
        {
            this.Enabled = true;

            PageButtonSetup();
            PanelSetup();

            if (f.tbMainBlankets.Text != "0")
            {
                rbProtectriveBlanketsYes.Checked = true;
            }
            else
            {
                rbProtectiveBlanketsNo.Checked = true;
            }

            if (f.loadingPreviousData)
            {
                PullInfo();
            }

            for (int i = 1; i < f.numberOfPagesNeeded; i++)
            {
                NewPage();
            }
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
                #region TB Load
                f.LoadPreviousXmlTb(tbfname, tblname, tbphone, tbAddress1, tbAddress2, tbAddress3, tbLiftNumbers, tbTypeofLift,
                    tbShaftDepth, tbShaftWidth, tbPitDepth, tbHeadroom, tbTravel, tbNumofLandingDoors, tbNumofLandings,
                    tbNumberOfCOPS, tbMainCOPLocation, tbAuxCOPLocation, tbKeyswitchLocation, tbDesignations, tbNumOfLEDLights, tbFloorFinish,
                    tbDoorWidth, tbDoorHeight, tbLoad, tbSpeed, tbwidth, tbDepth, tbHeight, tbLiftRating, tbNumofCarEntrances, tbLiftCarNotes,
                    tb2AuxCOPLocation, tb2CarDepth, tb2CarDoorFinishText,
                    tb2CarHeight, tb2CarLoad, tb2CarWidth, tb2CeilingFinishText, tb2ControlerLocationText, tb2COPFinishText, tb2Designations,
                    tb2DoorHeight, tb2DoorTracksText, tb2DoorWidth, tb2FacePlateMaterialText, tb2FloorFinish, tb2FrontWallText,
                    tb2HandrailText, tb2Headroom, tb2KeyswitchLocation, tb2LandingDoorFinishText, tb2LiftNumbers, tb2LiftRating,
                    tb2MainCOPLocation, tb2MirrorText, tb2Note, tb2NumberOfCarEntrances, tb2NumberOfCOPs, tb2NumberOfLAndingDoors,
                    tb2NumberOfLandings, tb2NumberofLEDLights, tb2PitDepth, tb2RearWallText, tb2ShaftDepth, tb2ShaftWidth,
                    tb2SideWallText, tb2Speed, tb2StructureShaftText, tb2Travel, tb2TypeOfLift, tb3AuxCOPLocation, tb3CarDepth, tb3CarDoorFinishText,
                    tb3CarHeight, tb3CarNote, tb3CarWidth, tb3CEilingFinishText, tb3ControlerLocationText, tb3COPFinishText, tb3Designations,
                    tb3DoorHeight, tb3DoorTracksText, tb3DoorWidth, tb3FacePlaterMaterialText, tb3FloorFinish, tb3FrontWallText,
                    tb3HandrailText, tb3HeadRoom, tb3KeyswitchLocation, tb3LandingDoorFinishText, tb3LiftNumbers, tb3LiftRating,
                    tb3Load, tb3MainCOPLocation, tb3MirrorText, tb3NumberOfCarEntrances, tb3NumberOfCOPs, tb3NumberOfLandingDoors,
                    tb3NumberOfLandings, tb3NumberOfLEDLights, tb3PitDepth, tb3RearWallText, tb3ShaftDepth, tb3ShaftWidth,
                    tb3SideWallText, tb3Speed, tb3StructureShaftText, tb3Travel, tb3TypeOfLift, tb4AuxCOPLocation, tb4CarDepth, tb4CarDoorFinish,
                    tb4CarHeight, tb4CarNote, tb4CarWidth, tb4CeilingFinishText, tb4ControlerLocationText, tb4COPFinishText, tb4Designations,
                    tb4DoorHeight, tb4DoorTracksText, tb4DoorWidth, tb4FacePlateMaterialText, tb4FloorFinish, tb4FrontWallText,
                    tb4HandrailText, tb4Headroom, tb4KeyswitchLocations, tb4LandingDoorFinishText, tb4LiftNumbers, tb4LiftRating,
                    tb4Load, tb4MainCOPLocation, tb4MirrorText, tb4NumberOfCarEntrances, tb4NumberOfCOPs, tb4NumberOfLandingDoors,
                    tb4NumberOfLandings, tb4NumbeROfLEDLights, tb4PitDepth, tb4RearWallText, tb4ShaftDepth, tb4ShaftWidth, tb4SideWallText,
                    tb4Speed, tb4StructureShaftText, tb4Travel, tb4TypeOfLift, tb5AuxCOPLocation, tb5CarDepth, tb5CarDoorFinishText,
                    tb5CaRHeight, tb5CarNote, tb5CarWidth, tb5CeilingFinishText, tb5ControlerLocationText, tb5COPFinishText, tb5Designations,
                    tb5DoorHeight, tb5DoorTRacksText, tb5DoorWidth, tb5FacePlateMaterialText, tb5FloorFinish, tb5FrontWallText,
                    tb5HandrailText, tb5Headroom, tb5KetyswitchLocation, tb5LandingDoorFinishText, tb5LiftNumbers, tb5LiftRating,
                    tb5Load, tb5MainCOPLocation, tb5MirrorText, tb5NumberOfCarEntrances, tb5NumberOfCOPs, tb5NumberOfLandingDoors,
                    tb5NumberOfLandings, tb5NumberOIfLEDLights, tb5PitDepth, tb5RearWallText, tb5ShaftDEpth, tb5ShaftWidth,
                    tb5SideWallText, tb5Speed, tb5StructureShaftText, tb5Travel, tb5TypeOfLift, tb6AuxCOPLocation, tb6CarDepth, tb6CarDoorFinishText,
                    tb6CarHeight, tb6CarLoad, tb6CarNote, tb6CarSpeed, tb6CarWidth, tb6CeilingFinishText, tb6ControlerLocationText, tb6COPFinishText,
                    tb6Designations, tb6DoorHeight, tb6DoorTracksOther, tb6DoorWidth, tb6FacePlateMaterialText, tb6FloorFinish,
                    tb6FrontWallText, tb6HAndrailText, tb6Headroom, tb6KeySwitchLocation, tb6LandingDoorFinishText, tb6LiftNumbers,
                    tb6LiftRating, tb6MainCOPLocation, tb6MirrorText, tb6NumberOfCarEntrances, tb6NumberOFCOPs, tb6NumberOfLandingDoors,
                    tb6NumberOfLandings, tb6NumberOfLEDLights, tb6PitDepth, tb6RearWallText, tb6ShaftDepth, tb6ShaftWidth,
                    tb6SideWallText, tb6StructureShaftText, tb6Travel, tb6TypeOfLift, tb7AuzCOPLocation, tb7CarDepth, tb7CarDoorFinishText,
                    tb7CarHeight, tb7CarLoad, tb7CarNotes, tb7CarSpeed, tb7CarWidth, tb7CEilingFinishText, tb7ControlerLocationText, tb7COPFinishText,
                    tb7Designations, tb7DoorHeight, tb7DoorTracksText, tb7DoorWidth, tb7FacePlateMaterialText, tb7FloorFinish, tb7FrontWallText,
                    tb7HandrailText, tb7HeadRoom, tb7KeyswitchLocation, tb7LandingDoorFinishText, tb7LiftNumbers, tb7LiftRating,
                    tb7MainCOPLocation, tb7MirrorText, tb7NumberOfCarEntrances, tb7NumberOfCOPs, tb7NumberOfLandingDoors, tb7NumberOfLandings,
                    tb7NumberOfLEDLights, tb7PitDepth, tb7RearWallText, tb7ShaftDepth, tb7ShaftWidth, tb7SideWallText,
                    tb7StructureShaftText, tb7Travel, tb7TypeOfLift, tb8AuxCOPLocation, tb8CarDEpth, tb8CarDoorFinishText,
                    tb8CarHeight, tb8CarWidth, tb8CeilingFinishText, tb8ControlerLocationText, tb8COPFinishText, tb8Desiginations, tb8DoorHeight,
                    tb8DoorTracksText, tb8DoorWidth, tb8FacePlateMaterialText, tb8FloorFinish, tb8FrontWallText, tb8HandrailText,
                    tb8Headroom, tb8KeyswitchLocations, tb8LandingDoorFinishText, tb8LiftCarNotes,
                    tb8LiftNumbers, tb8LiftRating, tb8Load, tb8MainCOPLocation, tb8MirrorText, tb8NumberOfCarEntrances, tb8NumberOfCOPs,
                    tb8NumberOfLandingDoors, tb8NumberOfLandings, tb8NumberofLEDLights, tb8PitDepth, tb8RearWallText,
                    tb8ShaftDepth, tb8ShaftWidth, tb8SideWallText, tb8Speed, tb8StructureShaftText, tb8Travel, tb8TypeOfLift, tb9AuxCOPLocation, tb9CarDepth, tb9CarDoorFinishText, tb9CarHeight,
                    tb9CarNotes, tb9CarWidth, tb9CeilingFinishText, tb9ControlerLocationText, tb9COPFinishText, tb9Designations, tb9DoorHeight, tb9DoorTracksText,
                    tb9DoorWidth, tb9FacePlateMaterialText, tb9FloorFinish, tb9FrontWallText, tb9HandrailTexrt, tb9Headroom, tb9KeyswitchLocation,
                    tb9LandingDoorFinishText, tb9LiftNumbers, tb9LiftRating, tb9Load, tb9MainCOPLocation, tb9MirrorText, tb9NumberOFCarEntraces,
                    tb9NumberOfCOPs, tb9NumberOfLandingDoors, tb9NumberOfLandings, tb9NumberOfLEDLights, tb9PitDepth, tb9RearWallText,
                    tb9ShaftDepth, tb9ShaftWidth, tb9SideWallText, tb9Speed, tb9StructureShaftText, tb9Travel, tb9TypeOfLift, tb10AuxCOPLocation, tb10CarDepth, tb10CarDoorFinishText,
                    tb10CarHeight, tb10CarWidth, tb10CEilingFinishText, tb10ControlerLocationText, tb10COPFinishText, tb10Desigination, tb10DoorHeight,
                    tb10DoorTracksText, tb10DoorWidth, tb10FacePlateMaterialText, tb10FloorFinish, tb10FrontWallText, tb10HandrailText,
                    tb10Headroom, tb10KeyswitchLocation, tb10LandingDoorFinishText, tb10LiftCarLoad, tb10LiftCarNotes, tb10LiftNumbers,
                    tb10LiftRating, tb10MainCOPLocation, tb10MirrorText, tb10NumberofCarEntrances, tb10NumberOfCOPs, tb10NumberofLandingDoors,
                    tb10NumberofLandings, tb10NumberOfLEDLIghts, tb10PitDepth, tb10RearWallText, tb10ShaftDepth, tb10ShaftWidth,
                    tb10SideWallText, tb10Speed, tb10StructureShaftText, tb10Travel, tb10TypeOfLift, tb11AuxCOPLocation, tb11CarDepth, tb11CarDoorFinishText,
                    tb11CarHeight, tb11CarWidth, tb11CeilingFinishText, tb11ControlerLocationText, tb11COPFinishText, tb11Designations, tb11DoorHeight,
                    tb11DoorTracksText, tb11DoorWidth, tb11FaceplateMaterialText, tb11FloorFinish, tb11FrontWallText, tb11HandrailText, tb11Headroom,
                    tb11KeyswitchLocation, tb11LandingDoorFInishOther, tb11LiftCarLoad, tb11LiftCarNote, tb11LiftRating, tb11MainCOPLocation,
                    tb11MirrorText, tb11NumberofCarEntrances, tb11NumberOfCOPs, tb11NumberOfLandingDoors, tb11NumberOfLandings, tb11NumberOfLEDLights,
                     tb11PitDepth, tb11RearWallText, tb11ShaftDepth, tb11ShaftWidth, tb11SideWallText, tb11Speed, tb11Travel, tb11TypeOfLift,
                    rb11LiftNumbers, rb11StructureShaftText, tb12AuxCOPLocation, tb12CarDepth, tb12CarDoorFinishText,
                    tb12CarHeight, tb12CarLiftRating, tb12CarLoad, tb12CarNumberOfCarEntrances, tb12CarSpeed, tb12CarWidth, tb12CeilingFinishText,
                    tb12ControlerLocationText, tb12COPFinishText, tb12Designations, tb12DoorTracksText, tb12FacePlateMaterialText,
                    tb12FloorFinish, tb12FrontWallText, tb12HandrailText, tb12Headroom, tb12KeyswitchLocation, tb12LandingDoorFinishText,
                    tb12LandingDoorHeight, tb12LandingDoorWidth, tb12LiftCarNotes, tb12LiftNumbers, tb12MainCOPLocation, tb12MirrorText,
                    tb12NumberOfCOPs, tb12NumberOfLandingDoors, tb12NumberOfLandings, tb12NumberOfLEDLights, tb12PitDepth,
                     tb12RearWallText, tb12ShaftDepth, tb12ShaftWidth, tb12SideWallText, tb12StructureShaftText, tb12Travel, tb12TypeOfLift,
                     tbControlerLocation, tbStructureShaft, tbLandingDoorFinish, tbDoorTracks, tbCarDoorFinish, tbCeilingFinish, tbFrontWall,
                     tbMirror, tbHandrail, tbSideWall, tbRearWall, tbCOPFinish, tbFacePlateMaterial
                    );
                #endregion

                #region RB Load
                f.LoadPreviousXmlRb(rb10AdvancedOpeningNo, rb10AdvancedOpeningYes, rb10BumpRaidYes, rb10BumpRailNo, rb10CarDoorFinishOther,
                    rb10CarDoorFinishStainlessSteel, rb10CEilingFinishMirrorStainlessSteel, rb10CEilingFinishOther, rb10CeilingFinishStainlessSteel,
                    rb10CeilingFinishWhite, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther, rb10ControlerLocationShaft,
                    rb10ControlerLocationTopLanding, rb10COPFinishOther, rb10COPFinishStainlessSteel, rb10DoorNudgingNo, rb10DoorNudgingYes,
                    rb10DoorTracksAluminium, rb10DoorTracksOther, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening, rb10EmergencyLoweringSystemYes,
                    rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes, rb10FacePlateMaterialOther, rb10FacePlateMaterialStainlessSteel, rb10FalseCeilingNo,
                    rb10FalseCeilingYes, rb10FalseFloorNo, rb10FalseFloorYes, rb10FireSERviceNo, rb10FireSErviceYes, rb10FrontWallOther,
                    rb10FrontWallStainlessSteel, rb10GPOInCarNo, rb10GPOInCarYes, rb10HandrailOther, rb10HandrailStainlessSteel, rb10IndependentServiceNo,
                    rb10IndependentServiceYes, rb10LAndingDoorFinishOtherr, rb10LandingDoorFinishStainlessSteel, rb10LCDColourBlue, rb10LCDColourRed,
                    rb10LCDColourWhite, rb10LoadWEighingNo, rb10LoadWeighingYes, rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther,
                    rb10OutOfServiceNo, rb10OutOFServiceYes, rb10PositionIndicatorTypeFlushMount, rb10PositionIndicatorTypeSurfaceMount,
                    rb10ProtectiveBlanketNo, rb10ProtectiveBlanketYes, rb10RearDoorKeySwitchNo, rb10RearDoorKeySwitchYes, rb10RearWallOther,
                    rb10RearWallStainlessSteel, rb10SecurityKeySwitchNo, rb10SecurityKeySwitchYes, rb10SideWallOther, rb10SideWallStainlesSteel,
                    rb10StructureShaftConcrete, rb10StructureShaftOther, rb10TimmerbeamsYes, rb10TrimmerBeamsNo, rb10VoiceAnnunciationNo,
                    rb10VoiceAnnunciationYes, rb11AdvancedOpeningNo, rb11AdvancedOpeningYes, rb11BumpRailNo, rb11BumpRailYes, rb11CarDoorFinishOther,
                    rb11CarDoorFinishStainlessSteel, rb11CeilingFinishMirrorStainlessSteel, rb11CeilingFinishOther, rb11CeilingFinishStainlessSteel,
                    rb11CeilingFinishWhite, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther, rb11ControlerLocationShaft, rb11COPFinishOther,
                    rb11COPFinishStainlessSteel, rb11DoorNudgingNo, rb11DoorNudgingYes, rb11DoorTracksAluminium, rb11DoorTracksOther,
                    rb11DoorTypeCentreOpening, rb11DoorTypeSideOpening, rb11EmergencyLoweringSystemNo, rb11EmergencyLoweringSystemYes,
                    rb11ExclusiveServiceNo, rb11ExclusiveServiceYes, rb11FacePlateMaterialOther, rb11FacePlateMaterialStainlessSTeel, rb11FalseCeilingNo,
                    rb11FalseCEilingYes, rb11FalseFloorNo, rb11FalseFloorYes, rb11FireServiceNo, rb11FireServiceYes, rb11FrontWallOther,
                    rb11FrontWallStainlessSteel, rb11GPOInCarNo, rb11GPOInCarYes, rb11HandrailOther, rb11HandrailStainlessSteel,
                    rb11IndependentSErviceNO, rb11IndependentServiceYes, rb11LandingDoorFinishOther, rb11LandingDoorFinishStainlessSteel,
                    rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite, rb11LoadWeighingNo, rb11LoadWeighingYes,
                    rb11MirrorFullSize, rb11MirrorHalfSize, rb11MirrorOther, rb11OutOfServiceNo, rb11OutOfSErviceYes, rb11PositionIndicatorTypeFlushMount,
                    rb11PositionIndicatorTypeSurfaceMount, rb11ProtectiveBlanketNo, rb11ProtectiveBlanketsYes, rb11RearDoorKeySwitchNo,
                    rb11RearDoorKeySwitchYes, rb11RearWallOther, rb11RearWallStainlessSteel, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes,
                    rb11SideWallOther, rb11SideWallStainlessSteel, rb11StrructureShaftOther, rb11StructureShaftConcrete,
                    rb11TrimmerBeamNo, rb11TrimmerBeamsYes, rb11VoiceAnnunciationNo, rb11VoiceAnnunciationYes, rb12AdvancedOpeningNo,
                    rb12AdvancedOpeningYes, rb12BumpRailNo, rb12BumpRailYes, rb12CarDoorFinishOther, rb12CarDoorFinishStainlessSteel,
                    rb12CeilingFinishMirrorStainlessSteel, rb12CeilingFinishOther, rb12CeilingFinishStainlessSteel, rb12CeilingFinishWhite,
                    rb12ControlerLocationBottomLanding, rb12ControlerLocationOther, rb12ControlerLocationShaft, rb12ControlerLocationTopLanding,
                    rb12COPFinishOther, rb12COPFinishStainlessSteel, rb12DoorTracksAluminium, rb12DoorTracksOther, rb12DoorTypeCentreOpening,
                    rb12DoorTypeSideOpening, rb12EmergencyLoweringSystemNo, rb12EmergencyLoweringSystemYes, rb12ExclusiveServiceNo,
                    rb12ExclusiveServiceYes, rb12FacePlateMaterialOther, rb12FacePlateMaterialStainlessSteel, rb12FalseCeilingNo, rb12FalseCeilingYes,
                    rb12FalseFloorNo, rb12FalseFloorYes, rb12FireServiceNo, rb12FireServiceYes, rb12FrontWallOther, rb12FrontWallStainlessSteel,
                    rb12GPOInCarNo, rb12GPOInCarYes, rb12HandrailOther, rb12HandrailStainlessSTeel, rb12IndependentServiceNo, rb12IndependentServiceYes,
                    rb12LandingDoorFinishOther, rb12LandingDoorFinishStainlessSteel, rb12LandingDoorNudgingNo, rb12LandingDoorNudgingYes,
                    rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite, rb12LoadWeighingNo, rb12LoadWeighingYes, rb12MirrorFullSize,
                    rb12MirrorHalfSize, rb12MirrorOTher, rb12OutOfServiceNo, rb12OutOfServiceYes, rb12PositionIndicatorTypeFlushMount,
                    rb12PositionIndicatorTypeSurfaceMount, rb12ProectiveBlanketsYes, rb12ProtectiveBlanketsNo, rb12RearDoorKeySwitchNo,
                    rb12RearDoorKeySwitchYes, rb12RearWallOther, rb12RearWallStainlessSteel, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes,
                    rb12SideWallOther, rb12SideWallStainlessSteel, rb12StructureShaftConcrete, rb12StructureShaftOther, rb12TrimmerBeamsNo,
                    rb12TrimmerBeamsYes, rb12VoicAnnunciationNo, rb12VoiceAnnuniationYes, rb2AdvancedOpeningNo, rb2AdvancedOpeningYes,
                    rb2BumpRailNo, rb2BumpRailYes, rb2CarDoorFinishOther, rb2CarDoorFinishStainlessSteel, rb2CeilingFinishMirrorStainlessSteel,
                    rb2CeilingFinishOther, rb2CeilingFinishStainlessSteel, rb2CeilingFinishWhite, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther,
                    rb2ControlerLocationShaft, rb2ControlerLocationTopLanding, rb2COPFinishOther, rb2COPFinishStainlessSTeell, rb2DoorNudgingNo,
                    rb2DoorNudgingYes, rb2DoorTracksAluminium, rb2DoorTracksOther, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening,
                    rb2EmergemncyLoweringSystemNo, rb2EmergencyLoweringSystemYes, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes,
                    rb2FacePlateMaterialOther, rb2FacePlateMaterialStainlessSteel, rb2FalseCeilingNo, rb2FalseCeilingYes, rb2FalseFloorNo,
                    rb2FalseFloorYes, rb2FireSErviceNo, rb2FireSErviceYes, rb2FrontWallOther, rb2FrontWallStainlessSteel, rb2GPOInCarNo,
                    rb2GPOInCarYes, rb2HandrailOther, rb2HandRailStainlessSteel, rb2IndependentServiceNo, rb2IndependentServiceYes,
                    rb2LandingDoorFinishOther, rb2LandingDoorFinishStainlessSteel, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite,
                    rb2LoadWeighingNo, rb2LoadWeighingYes, rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther, rb2OutOfServiceNo,
                    rb2OutOfServiceYes, rb2PositionIndicatorTypeFlushMount, rb2PositionIndicatorTypeSurfaceMount, rb2ProtectiveBlanketsNo,
                    rb2ProtectiveBlanketsYes, rb2RearDoorKeySwitchNo, rb2RearDoorKeySwitchYes, rb2RearWallOther, rb2RearWallStainlessSteel,
                    rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes, rb2SideWallOther, rb2SideWallStainlessSteel, rb2StructureShaftConcrete,
                    rb2StructureShaftOther, rb2TrimmerBeamsNo, rb2TrimmerBeamsYes, rb2VoiceAnnunciationNo, rb2VoiceAnnunciationYes,
                    rb3AdvancedOpeningNo, rb3AdvancedOpeningYes, rb3BumpRailNo, rb3BumpRailYes, rb3CarDoorFinishOther, rb3CarDoorFinishStainlessSteel,
                    rb3CeilingFinishOther, rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite, rb3ControleRLocationBottomLanding, rb3ControlerLocationOther,
                    rb3ControlerLocationShaft, rb3ControlerLocationTopLanding, rb3COPFinishOther, rb3COPFinishStainlessSteel, rb3DoorNudgingNo,
                    rb3DoorNudgingYes, rb3DoorTracksAluminium, rb3DoorTracksOther, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening,
                    rb3EmergencyLoweringSystemNo, rb3EmergencyLoweringSystemYes, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes,
                    rb3FacePlateMaterialOther, rb3FacePlateMaterialStainlessSteel, rb3FalseCeilingNo, rb3FalseCeilingYes, rb3FalseFloorNo, rb3FalseFloorYes,
                    rb3FireServiceYes, rb3FireServieNo, rb3FrontWallOther, rb3FrontWallStainlessSteel, rb3GPOInCarNo, rb3GPOInCarYes,
                    rb3HandrailOther, rb3HandrailStainlessSteel, rb3IndependentServiceNo, rb3IndependentServiceYes, rb3landingDoorFinishOther,
                    rb3LandingDoorFinishStainlessSteel, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite, rb3LoadWeighingNo, rb3LoadWeighingYes,
                    rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther, rb3MirrorStainlessSteel, rb3OutOfServiceNo, rb3OutOfSErviceYes,
                    rb3PositionIndicatorTypeFlushMount, rb3PositionIndicatorTypeSurfaceMount, rb3ProtectiveBlanketsNo, rb3ProtectiveBlanketsYes,
                    rb3RearDoorKeySwitchNo, rb3RearDoorKeySwitchYes, rb3RearWallOther, rb3RearWallStainlessSteel, rb3SecurityKeySwitchNo,
                    rb3SecurityKeySwitchYes, rb3SideWallOther, rb3SideWallStainlessSteel, rb3StructureShaftConcrete, rb3StructureShaftOther,
                    rb3TrimmerBeamsNo, rb3TrimmerBeamsYes, rb3VoiceAnnunciationNo, rb3VoiceAnnunciationYes, rb4AdvancedOpeningNo,
                    rb4AdvancedOpeningYes, rb4BumpRailNo, rb4BumpRailYes, rb4CarDoorFinishOther, rb4CarDoorFinishStainlessSteel,
                    rb4CeilingFinishMirrorStainlessSteel, rb4CeilingFinishOther, rb4CeilingFinishStainlessSteel, rb4CeilingFinishWhite,
                    rb4ControelrLocationTopLanding, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther, rb4ControlerLocationShaft,
                    rb4COPFinishOther, rb4COPFinishStainlessSteel, rb4DoorNudgingNo, rb4DoorNudgingYes, rb4DoorTracksAluminium, rb4DoorTracksOther,
                    rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening, rb4EmergencyLoweringSystemNo, rb4EmergencyLoweringSystemYes,
                    rb4ExclusiveServiceNo, rb4ExclusiveServiceYes, rb4FacePlateMaterialOther, rb4FacePlateMaterialStainlessSteel, rb4FalseCeilingNO,
                    rb4FalseCeilingYes, rb4FalseFloorNo, rb4FalseFloorYes, rb4FireSErviceNo, rb4FireServiceYes, rb4FrontWallStainlessSteel, rb4FrotnWallOther,
                    rb4GPOInCarNo, rb4GPOInCarYes, rb4HandrailOther, rb4HandRailStainlesSteel, rb4IndependentServiceYes, rb4LandingDoorFinishOther,
                    rb4LandingDoorFinishStainlessSteel, rb4LCDColourBlue, rb4LCDColourRed, rb4LCDColourWhite, rb4LoadWeighingNo, rb4LoadWeighingYes,
                    rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther, rb4OutOfServiceNo, rb4OutOfServiceYes, rb4PositionIndicatorTypeFlushMount,
                    rb4PositionIndicatorTypeSurfaceMount, rb4ProtectiveBlanketsYes, rb4ProtetiveBlanketsNo, rb4RearDoorKeySwitchNo,
                    rb4RearDoorKeySwitchYes, rb4RearWallOther, rb4RearWallStainlessSteel, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes,
                    rb4SideWallOther, rb4SideWallStainlessSteel, rb4StructureShaftConcrete, rb4StructureShaftOther, rb4TrimmerBeamsNo,
                    rb4TrimmerBeamsYes, rb4VoiceAnnunciationNo, rb4VoiceAnnunciationYes, rb5AdvancedOpeningNo, rb5AdvancedOpeningYes, rb5BumpRailNo,
                    rb5BumpRailYes, rb5CarDoorFinishOther, rb5CarDoorFinishStainlessSteel, rb5CeilingFinishMirrorStainlessSTeel, rb5CeilingFinishOther,
                    rb5CeilingFinishStainlessSteel, rb5CeilingFinishWhite, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther,
                    rb5ControlerLocationShaft, rb5ControlerLocationTopLanding, rb5COPFinishOther, rb5COPFinishStainlessSteel, rb5DoorNudgingNo,
                    rb5DoorNudgingYes, rb5DoorTracksAluminium, rb5DoorTracksOther, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening,
                    rb5EmergencyLoweringSystemNo, rb5EmergencyLoweringSystemYes, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes,
                    rb5FacePlateMAterialOtjer, rb5FacePlateMaterialStainlessSteel, rb5FalseCeilingNo, rb5FalseCeilingYes, rb5FalseFloorNo, rb5FalseFloorYes,
                    rb5FireSErviceNo, rb5FireServiceYes, rb5FrontWallOther, rb5FrontWallStainlessSteel, rb5GPOInCarNo, rb5GPOInCarYes,
                    rb5HandRailOther, rb5HandRailStainlessSteel, rb5IndependentServiceNo, rb5IndependentServiceYes, rb5LAndingDoorFinishOther,
                    rb5LandingDoorFinishStainlessSteel, rb5LCDColoiurWHite, rb5LCDColourBlue, rb5LCDColourRed, rb5LoadWeighingNo, rb5LoadWeighingYes,
                    rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther, rb5OutOfServiceNo, rb5OutOfServiceYes, rb5PositionIndicatorTypeSurfaceMount,
                    rb5PostionIndicatorTypeFlushMount, rb5ProtectiveBlanketsNo, rb5ProtectiveBlanketsYes, rb5RearDoorKeySwitchNo, rb5RearDoorKeySwitchYes,
                    rb5RearWallOther, rb5RearWallStainlesSteel, rb5SecurityKeySwitchNo, rb5SecurityServiceYes, rb5SideWallOther, rb5SideWallStainlessSTeel,
                    rb5StructureShaftConcrete, rb5StructureShaftOther, rb5TrimmerBeamsNo, rb5TrimmerBeamsYes, rb5VoiceAnnunciationNo,
                    rb5VoiceAnnunciationYes, rb6AdvancedOpeningNo, rb6AdvancedOpeningYes, rb6BumpRailNo, rb6BumpRailYes, rb6CarDoorFinishOther,
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
                    rb6StructureShaftCOncrete, rb6StructureShaftOther, rb6TrimmerBeamsNo, rb6TrimmerBeamsYes, rb6VoiceAnnunciationNo,
                    rb6VoiceAnnunciationYes, rb7AdvancedOpeningNo, rb7AdvancedOpeningYes, rb7BumpRailNo, rb7BumpRailYes, rb7CarDoorFinishOther,
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
                    rb7SideWallOther, rb7SideWallStainlessSteel, rb7StructureShaftConcrete, rb7StructureShaftOther, rb7TrimmerBeamsNo,
                    rb7TrimmerBeamsYes, rb7VoiceAnnunciationYes, rb7VoiceAnunciationNo, rb8AdvancedOpeningYes, rb8AdvncedOpeningNo, rb8BumpRailNo,
                    rb8BumpRailYes, rb8CarDoorFinishOther, rb8CarDoorFinishStainlessSteel, rb8CeilingFinishMirrorStainlessSTeel, rb8CeilingFinishOther,
                    rb8CeilingFinishStainlessSteel, rb8CeilingFinishWhite, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther,
                    rb8ControlerLocationShaft, rb8ControlerLocationTopLanding, rb8COPFinishOther, rb8COPFinishStainlessSteel, rb8DoorNudgingYes,
                    rb8DoorTracksAluminium, rb8DoorTracksOther, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening, rb8EmergencyLoweringSystemNo,
                    rb8EmergencyLoweringSystemYes, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes, rb8FacePlateMaterialOther,
                    rb8FacePlateMaterialStainlessSteel, rb8FalseCeilingNo, rb8FalseCeilingYes, rb8FalseFloorNo, rb8FalseFloorYes, rb8FireSErviceNo,
                    rb8FireServiceYes, rb8FrontWallOther, rb8FrontWallStainlessSteel, rb8GPOInCarNo, rb8GPOInCarYes, rb8HandrailOther,
                    rb8HandRailStainlessSteel, rb8IndependentServiceNo, rb8IndependentServiceYes, rb8LandingDoorFinishOther, rb8LCDColourBlue,
                    rb8LCDColourRed, rb8LCDColourWhite, rb8LoadWeighingNo, rb8LoadWeighingYes, rb8MirrorFullSize, rb8MirrorHalfSize,
                    rb8MirrorOther, rb8OutOFServiceNo, rb8OutOfSErviceYes, rb8PositionIndicatorTypeFlushMount, rb8PositionIndicatorTypeSurfaceMount,
                    rb8ProtectiveBlanketsNo, rb8ProtectiveBlanketsYes, rb8RearDoorKeySwitchNo, rb8RearDoorKeySwitchYes, rb8RearWallOther,
                    rb8RearWallStainlessSteel, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes, rb8SideWallOther, rb8SideWallStainlessSteel,
                    rb8StructureShaftConcrete, rb8StructureShaftOther, rb8TimmemrBeamsYes, rb8TrimmerBeamsNo, rb8VoiceAnnunicationNo,
                    rb8VoiceAnnunicationYes, rb9AdvancedOpeningNo, rb9AdvancedOpeningYes, rb9BumpRailNo, rb9BumpRailYes, rb9CarDoorFinishOther,
                    rb9CarDoorFinishStainlessSteel, rb9CeilingFinishMirrorStainlessSteel, rb9CeilingFinishOther, rb9CeilingFinishStainlessSteel,
                    rb9CeilingFinishWhite, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther, rb9ControlerLocationShaft,
                    rb9ControlrLocationTopLanding, rb9COPFinishOther, rb9COPFinishStainlessSteel, rb9DoorNudgingNo, rb9DoorNudgingYes,
                    rb9DoorTracksAluminium, rb9DoorTracksOther, rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening, rb9EmergencyLoweringSystemNo,
                    rb9EmergencyLoweringSystemYes, rb9ExclusiveServiceNo, rb9ExclusiveServiceYes, rb9FacePlateMaterialOther,
                    rb9FacePlateMaterialStainlessSteel, rb9FalseCeilingNo, rb9FalseCeilingYes, rb9FalseFloorNo, rb9FalseFloorYes, rb9FireServiceNo,
                    rb9FireSErviceYes, rb9FrontWallOther, rb9FrontWallStainlessSteel, rb9GPOInCarNo, rb9GPOInCarYes, rb9HandrailOther,
                    rb9HandrailStainlessSteel, rb9IndependentServiceNo, rb9IndependentServiceYes, rb9LandingDoorFinishOther,
                    rb9LandingDoorFinishStainlessSteel, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite, rb9LoadWeighingNo, rb9LoadWeighingYes,
                    rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther, rb9OutOfServiceNo, rb9OutOfServiceYes, rb9PositionIndicatorTypeFlushMount,
                    rb9PositionIndicatorTypeSurfaceMount, rb9ProtectiveBlanketsNo, rb9ProtectiveBlanketsYes, rb9RearDoorKeySwitchNo,
                    rb9RearDoorKeySwitchYes, rb9RearWallOther, rb9RearWallStainlessSteel, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes,
                    rb9SideWallOther, rb9SideWallStainlessSteel, rb9StructureShaftConcrete, rb9StructureShaftOther, rb9TrimmerBeamsNo,
                    rb9TrimmerBeamsYes, rb9VoiceAnnunciationNo, rb9VoiceAnnunciationYes, rbAdvancedOpeningNo, rbAdvancedOpeningYes, rbBumpRailNo,
                    rbBumpRailYes, rbCarDoorFInishBrushedStainlessSteel, rbCarDoorFinishOther, rbCeilingFinishBrushedStasinlessSteel,
                    rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther, rbCeilingFinishWhite, rbControlerLoactionTopLanding,
                    rbControlerLocationBottomLanding, rbControlerLocationOther, rbControlerlocationShaft, rbCOPFinishOther, rbCOPFinishSatinStainlessSteel,
                    rbDoorNudgingNo, rbDoorNudgingYes, rbDoorTracksAnodisedAluminium, rbDoorTracksOther, rbDoorTypeCentreOpening,
                    rbDoorTypeSideOpening, rbEmergencyLoweringSystemNo, rbEmergencyLoweringSystemYes, rbExclusiveServiceNo, rbExclusiveServiceYes,
                    rbFacePlateMaterialOther, rbFacePlateMaterialSatinStainlessSteel, rbFalseCeilingNo, rbFalseCeilingYes, rbFalseFloorNo, rbFalseFloorYes,
                    rbFireServiceNo, rbFireServiceYes, rbFrontWallBrushedStainlessSteel, rbGPOInCarNo, rbGPOInCarYes, rbHandrailBrushedStainlessSTeel,
                    rbHandrailOther, rbIndependentServiceNo, rbIndependentServiceYes, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel,
                    rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite, rbLoadWeighingNo, rbLoadWeighingYes, rbMirrorFullSize, rbMirrorHalfSize,
                    rbMirrorOther, rbOutofServiceNo, rbOutofServiceYes, rbPositionIndicatorTypeFlushMount, rbPositionIndicatorTypeSurfaceMount,
                    rbProtectiveBlanketsNo, rbProtectriveBlanketsYes, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes, rbRearWallBrushedStainlessSteel,
                    rbRearWallOther, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes, rbSideWallBrushedStainlessSteel, rbSideWallOther, rbSL,
                    rbStructureShaftConcrete, rbStructureShaftOther, rbSumasa, rbTrimmerBeamsNo, rbTrimmerBeamsYes, rbVoiceAnnunciationNo,
                    rbVoiceAnnunciationYes, rbWittur);
                #endregion
            }
            catch (Exception)
            {
                return;
            }
        }

        #endregion

        #region Data Formatting Methods
        private void tbNumofCarEntrances_TextChanged_1(object sender, EventArgs e)
        {
            if (!rearDoorChecker)
            {
                RearDoorChecker(tbNumofCarEntrances, rbRearDoorKeySwitchYes, rbRearDoorKeySwitchNo);
            }
        }

        private void RearDoorChecker(TextBox carEntrance, RadioButton rbYes, RadioButton rbNo)
        {
            try
            {
                if (int.Parse(carEntrance.Text) >= 2)
                {
                    rbYes.Checked = true;
                }
                else
                {
                    rbNo.Checked = true;
                }
                //if try fails the bool will remain false and thus be able to try again
                // if try is sucessful it will change the bool to true and prevent additional edits
                rearDoorChecker = true;
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
            string prefix = "P0";
            f.saveData["NumberOfPagesOpen"] = pageTracker.ToString();
            // QuoteInfo3 nF = new QuoteInfo3();
            //f.WordData("","");            //call WordData method in form 1 to send all info into the dictiinary for writing 

            if (pageTracker >= 1)
            {
                prefix = "P1";
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
                #region Page 1 Word Export
                f.WordData("AE105", tbfname.Text); //first name
                f.WordData("AE106", tblname.Text);//last name
                f.WordData("AE107", tbphone.Text);//phone number
                f.WordData("AE108", tbAddress1.Text);//address 1
                f.WordData("AE109", tbAddress2.Text);//address 2
                f.WordData("AE110", tbAddress3.Text);//address 3
                f.WordData(prefix + "AE111", f.RadioButtonHandeler(null, rbSL, rbWittur, rbSumasa)); //supplier
                f.WordData("AE114", f.TotalLifts().ToString());//lift number
                f.WordData("AE115", tbTypeofLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tbControlerLocation, rbControlerLoactionTopLanding, rbControlerlocationShaft, rbControlerLocationBottomLanding, rbControlerLocationOther));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rbFireServiceNo, rbFireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tbShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tbShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tbPitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tbHeadroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tbTravel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tbNumofLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tbNumofLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tbStructureShaft, rbStructureShaftConcrete, rbStructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tbLoad.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tbSpeed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tbwidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tbDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tbHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tbLiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tbLiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tbNumofCarEntrances.Text);//number of car entraces
                if (tbLiftCarNotes.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tbLiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tbDoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tbDoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tbLandingDoorFinish, rbLandingDoorFinishOther, rbLandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rbDoorTypeCentreOpening, rbDoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tbDoorTracks, rbDoorTracksAnodisedAluminium, rbDoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rbAdvancedOpeningNo, rbAdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rbDoorNudgingNo, rbDoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tbCarDoorFinish, rbCarDoorFinishOther, rbCarDoorFInishBrushedStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tbCeilingFinish, rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishWhite, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rbFalseCeilingNo, rbFalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rbBumpRailYes, rbBumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tbFloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tbFrontWall, rbFrontWallBrushedStainlessSteel, tbFrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tbMirror, rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tbHandrail, rbHandrailBrushedStainlessSTeel, rbHandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tbSideWall, rbSideWallBrushedStainlessSteel, rbSideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tbNumOfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tbRearWall, rbRearWallOther, rbRearWallBrushedStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tbNumberOfCOPS.Text); // number of COPS
                f.WordData(prefix + "AE169", tbMainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tbAuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tbDesignations.Text); // designations 
                f.WordData(prefix + "AE191", tbKeyswitchLocation.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tbCOPFinish, rbCOPFinishSatinStainlessSteel));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rbLEDColourRed, rbLEDColourBlue, rbLEDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rbExclusiveServiceNo, rbExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rbRearDoorKeySwitchNo, rbRearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rbGPOInCarNo, rbGPOInCarYes));//GPO in car
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rbPositionIndicatorTypeSurfaceMount, rbPositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tbFacePlateMaterial, rbFacePlateMaterialSatinStainlessSteel, rbFacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rbEmergencyLoweringSystemYes, rbEmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE178", f.CheckboxTrueToYes(f.cbMainSecurity));//security cabiling only 

                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rbIndependentServiceYes, rbIndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rbLoadWeighingYes, rbLoadWeighingNo));//load weighing
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rbFalseFloorYes, rbFalseFloorNo));//false floor
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rbTrimmerBeamsYes, rbTrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rbProtectriveBlanketsYes, rbProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rbVoiceAnnunciationYes, rbVoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rbOutofServiceYes, rbOutofServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 2)
            {
                prefix = "P2";
                #region Page 2 Saving
                f.SaveTbToXML(tb2AuxCOPLocation, tb2CarDepth, tb2CarDoorFinishText,
                    tb2CarHeight, tb2CarLoad, tb2CarWidth, tb2CeilingFinishText, tb2ControlerLocationText, tb2COPFinishText, tb2Designations,
                    tb2DoorHeight, tb2DoorTracksText, tb2DoorWidth, tb2FacePlateMaterialText, tb2FloorFinish, tb2FrontWallText,
                    tb2HandrailText, tb2Headroom, tb2KeyswitchLocation, tb2LandingDoorFinishText, tb2LiftNumbers, tb2LiftRating,
                    tb2MainCOPLocation, tb2MirrorText, tb2Note, tb2NumberOfCarEntrances, tb2NumberOfCOPs, tb2NumberOfLAndingDoors,
                    tb2NumberOfLandings, tb2NumberofLEDLights, tb2PitDepth, tb2RearWallText, tb2ShaftDepth, tb2ShaftWidth,
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
                    rb2StructureShaftOther, rb2TrimmerBeamsNo, rb2TrimmerBeamsYes,
                    rb2VoiceAnnunciationNo, rb2VoiceAnnunciationYes
                    );
                #endregion
                #region Page 2 Word Export
                f.WordData(prefix + "AE114", tb2LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb2TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb2IndependentServiceYes, rb2IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb2LoadWeighingYes, rb2LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb2ControlerLocationText, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther, rb2ControlerLocationShaft, rb2ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb2FireSErviceNo, rb2FireSErviceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb2ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb2ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb2PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb2Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb2Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb2NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb2NumberOfLAndingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb2StructureShaftText, rb2StructureShaftConcrete, rb2StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb2TrimmerBeamsYes, rb2TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb2FalseFloorYes, rb2FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb2CarLoad.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb2Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb2CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb2CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb2CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb2LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb2LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb2NumberOfCarEntrances.Text);//number of car entraces
                if (tb2Note.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb2Note.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb2DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb2DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb2LandingDoorFinishText, rb2LandingDoorFinishOther, rb2LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb2DoorTracksText, rb2DoorTracksAluminium, rb2DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb2AdvancedOpeningNo, rb2AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb2DoorNudgingNo, rb2DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb2CarDoorFinishText, rb2CarDoorFinishOther, rb2CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb2CeilingFinishText, rb2CeilingFinishStainlessSteel, rb2CeilingFinishWhite, rb2CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb2FalseCeilingNo, rb2FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb2BumpRailYes, rb2BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb2FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb2FrontWallText, rb2FrontWallStainlessSteel, rb2FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb2MirrorText, rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb2HandrailText, rb2HandRailStainlessSteel, rb2HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb2SideWallText, rb2SideWallStainlessSteel, rb2SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb2NumberofLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb2ProtectiveBlanketsYes, rb2ProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb2RearWallText, rb2RearWallOther, rb2RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb2NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb2MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb2AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb2Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb2KeyswitchLocation.Text); //key switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb2COPFinishText, rb2COPFinishStainlessSTeell, rb2COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb2RearDoorKeySwitchNo, rb2RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb2GPOInCarNo, rb2GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb2VoiceAnnunciationYes, rb2VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb2PositionIndicatorTypeSurfaceMount, rb2PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb2FacePlateMaterialText, rb2FacePlateMaterialStainlessSteel, rb2FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb2EmergencyLoweringSystemYes, rb2EmergemncyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb2OutOfServiceYes, rb2OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 3)
            {
                prefix = "P3";
                #region Page 3 Saving
                f.SaveTbToXML(tb3AuxCOPLocation, tb3CarDepth, tb3CarDoorFinishText,
                    tb3CarHeight, tb3CarNote, tb3CarWidth, tb3CEilingFinishText, tb3ControlerLocationText, tb3COPFinishText, tb3Designations,
                    tb3DoorHeight, tb3DoorTracksText, tb3DoorWidth, tb3FacePlaterMaterialText, tb3FloorFinish, tb3FrontWallText,
                    tb3HandrailText, tb3HeadRoom, tb3KeyswitchLocation, tb3LandingDoorFinishText, tb3LiftNumbers, tb3LiftRating,
                    tb3Load, tb3MainCOPLocation, tb3MirrorText, tb3NumberOfCarEntrances, tb3NumberOfCOPs, tb3NumberOfLandingDoors,
                    tb3NumberOfLandings, tb3NumberOfLEDLights, tb3PitDepth, tb3RearWallText, tb3ShaftDepth, tb3ShaftWidth,
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
                    rb3StructureShaftOther, rb3TrimmerBeamsNo, rb3TrimmerBeamsYes,
                    rb3VoiceAnnunciationNo, rb3VoiceAnnunciationYes
                    );
                #endregion
                #region Page 3 Word Export
                f.WordData(prefix + "AE114", tb3LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb3TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb3IndependentServiceYes, rb3IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb3LoadWeighingYes, rb3LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb3ControlerLocationText, rb3ControleRLocationBottomLanding, rb3ControlerLocationOther, rb3ControlerLocationShaft, rb3ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb3FireServieNo, rb3FireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb3ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb3ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb3PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb3HeadRoom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb3Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb3NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb3NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb3StructureShaftText, rb3StructureShaftConcrete, rb3StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb3TrimmerBeamsYes, rb3TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb3FalseFloorYes, rb3FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb3Load.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb3Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb3CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb3CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb3CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb3LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb3LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb3NumberOfCarEntrances.Text);//number of car entraces
                if (tb3CarNote.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb3CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb3DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb3DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb3LandingDoorFinishText, rb3landingDoorFinishOther, rb3LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb3DoorTracksText, rb3DoorTracksAluminium, rb3DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb3AdvancedOpeningNo, rb3AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb3DoorNudgingNo, rb3DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb3CarDoorFinishText, rb3CarDoorFinishOther, rb3CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb3CEilingFinishText, rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite, rb3MirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb3FalseCeilingNo, rb3FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb3BumpRailYes, rb3BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb3FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb3FrontWallText, rb3FrontWallStainlessSteel, rb3FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb3MirrorText, rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb3HandrailText, rb3HandrailStainlessSteel, rb3HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb3SideWallText, rb3SideWallStainlessSteel, rb3SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb3NumberOfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb3ProtectiveBlanketsYes, rb3ProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb3RearWallText, rb3RearWallOther, rb3RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb3NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb3MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb3AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb3Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb3KeyswitchLocation.Text); //key switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb3COPFinishText, rb3COPFinishStainlessSteel, rb3COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb3RearDoorKeySwitchNo, rb3RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb3SecurityKeySwitchNo, rb3SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb3GPOInCarNo, rb3GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb3VoiceAnnunciationYes, rb3VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb3PositionIndicatorTypeSurfaceMount, rb3PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb3FacePlaterMaterialText, rb3FacePlateMaterialStainlessSteel, rb3FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb3EmergencyLoweringSystemYes, rb3EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb3OutOfSErviceYes, rb3OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 4)
            {
                prefix = "P4";
                #region Page 4 Saving
                f.SaveTbToXML(tb4AuxCOPLocation, tb4CarDepth, tb4CarDoorFinish,
                    tb4CarHeight, tb4CarNote, tb4CarWidth, tb4CeilingFinishText, tb4ControlerLocationText, tb4COPFinishText, tb4Designations,
                    tb4DoorHeight, tb4DoorTracksText, tb4DoorWidth, tb4FacePlateMaterialText, tb4FloorFinish, tb4FrontWallText,
                    tb4HandrailText, tb4Headroom, tb4KeyswitchLocations, tb4LandingDoorFinishText, tb4LiftNumbers, tb4LiftRating,
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
                    rb4StructureShaftConcrete, rb4StructureShaftOther, rb4TrimmerBeamsNo,
                    rb4TrimmerBeamsYes, rb4VoiceAnnunciationNo, rb4VoiceAnnunciationYes
                    );
                #endregion
                #region Page 4 Word Export
                f.WordData(prefix + "AE114", tb4LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb4TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb4IndependentServiceYes, IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb4LoadWeighingYes, rb4LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb4ControlerLocationText, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther, rb4ControlerLocationShaft, rb4ControelrLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb4FireSErviceNo, rb4FireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb4ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb4ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb4PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb4Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb4Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb4NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb4NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb4StructureShaftText, rb4StructureShaftConcrete, rb4StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb4TrimmerBeamsYes, rb4TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb4FalseFloorYes, rb4FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb4Load.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb4Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb4CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb4CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb4CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb4LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb4LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb4NumberOfCarEntrances.Text);//number of car entraces
                if (tb4CarNote.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb4CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb4DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb4DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb4LandingDoorFinishText, rb4LandingDoorFinishOther, rb4LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb4DoorTracksText, rb4DoorTracksAluminium, rb4DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb4AdvancedOpeningNo, rb4AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb4DoorNudgingNo, rb4DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb4CarDoorFinish, rb4CarDoorFinishOther, rb4CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb4CeilingFinishText, rb4CeilingFinishStainlessSteel, rb4CeilingFinishWhite, rb4CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb4FalseCeilingNO, rb4FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb4BumpRailYes, rb4BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb4FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb4FrontWallText, rb4FrontWallStainlessSteel, rb4FrotnWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb4MirrorText, rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb4HandrailText, rb4HandRailStainlesSteel, rb4HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb4SideWallText, rb4SideWallStainlessSteel, rb4SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb4NumbeROfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb4ProtectiveBlanketsYes, rb4ProtetiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb4RearWallText, rb4RearWallOther, rb4RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb4NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb4MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb4AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb4Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb4KeyswitchLocations.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb4COPFinishText, rb4COPFinishStainlessSteel, rb4COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb4LCDColourBlue, rb4LCDColourRed, rb4LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb4ExclusiveServiceNo, rb4ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb4RearDoorKeySwitchNo, rb4RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb4GPOInCarNo, rb4GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb4VoiceAnnunciationYes, rb4VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb4PositionIndicatorTypeSurfaceMount, rb4PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb4FacePlateMaterialText, rb4FacePlateMaterialStainlessSteel, rb4FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb4EmergencyLoweringSystemYes, rb4EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb4OutOfServiceYes, rb4OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 5)
            {
                prefix = "P5";
                #region Page 5 Saving
                f.SaveTbToXML(tb5AuxCOPLocation, tb5CarDepth, tb5CarDoorFinishText,
                    tb5CaRHeight, tb5CarNote, tb5CarWidth, tb5CeilingFinishText, tb5ControlerLocationText, tb5COPFinishText, tb5Designations,
                    tb5DoorHeight, tb5DoorTRacksText, tb5DoorWidth, tb5FacePlateMaterialText, tb5FloorFinish, tb5FrontWallText,
                    tb5HandrailText, tb5Headroom, tb5KetyswitchLocation, tb5LandingDoorFinishText, tb5LiftNumbers, tb5LiftRating,
                    tb5Load, tb5MainCOPLocation, tb5MirrorText, tb5NumberOfCarEntrances, tb5NumberOfCOPs, tb5NumberOfLandingDoors,
                    tb5NumberOfLandings, tb5NumberOIfLEDLights, tb5PitDepth, tb5RearWallText, tb5ShaftDEpth, tb5ShaftWidth,
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
                #region Page 5 Word Export
                f.WordData(prefix + "AE114", tb5LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb5TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb5IndependentServiceYes, rb5IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb5LoadWeighingYes, rb5LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb5ControlerLocationText, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther, rb5ControlerLocationShaft, rb5ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb5FireSErviceNo, rb5FireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb5ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb5PitDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb5PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb5Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb5Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb5NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb5NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb5StructureShaftText, rb5StructureShaftConcrete, rb5StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb5TrimmerBeamsYes, rb5TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb5FalseFloorYes, rb5FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb5Load.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb5Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb5CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb5CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb5CaRHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb5LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb5LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb5NumberOfCarEntrances.Text);//number of car entraces
                if (tb5CarNote.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb5CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb5DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb5DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb5LandingDoorFinishText, rb5LAndingDoorFinishOther, rb5LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb5DoorTRacksText, rb5DoorTracksAluminium, rb5DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb5AdvancedOpeningNo, rb5AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb5DoorNudgingNo, rb5DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb5CarDoorFinishText, rb5CarDoorFinishOther, rb5CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb5CeilingFinishText, rb5CeilingFinishStainlessSteel, rb5CeilingFinishWhite, rb5CeilingFinishMirrorStainlessSTeel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb5FalseCeilingNo, rb5FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb5BumpRailYes, rb5BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb5FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb5FrontWallText, rb5FrontWallStainlessSteel, rb5FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb5MirrorText, rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb5HandrailText, rb5HandRailOther, rb5HandRailStainlessSteel));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb5SideWallText, rb5SideWallStainlessSTeel, rb5SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb5NumberOIfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb5ProtectiveBlanketsYes, rb5ProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb5RearWallText, rb5RearWallOther, rb5RearWallStainlesSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb5NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb5MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb5AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb5Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb5KetyswitchLocation.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb5COPFinishText, rb5COPFinishStainlessSteel, rb5COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb5LCDColourBlue, rb5LCDColourRed, rb5LCDColoiurWHite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb5RearDoorKeySwitchNo, rb5RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb5SecurityKeySwitchNo, rb5SecurityServiceYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb5GPOInCarNo, rb5GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb5VoiceAnnunciationYes, rb5VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb5PositionIndicatorTypeSurfaceMount, rb5PostionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb5FacePlateMaterialText, rb5FacePlateMaterialStainlessSteel, rb5FacePlateMAterialOtjer));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb5EmergencyLoweringSystemYes, rb5EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb5OutOfServiceYes, rb5OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 6)
            {
                prefix = "P6";
                #region Page 6 Saving
                f.SaveTbToXML(tb6AuxCOPLocation, tb6CarDepth, tb6CarDoorFinishText,
                    tb6CarHeight, tb6CarLoad, tb6CarNote, tb6CarSpeed, tb6CarWidth, tb6CeilingFinishText, tb6ControlerLocationText, tb6COPFinishText,
                    tb6Designations, tb6DoorHeight, tb6DoorTracksOther, tb6DoorWidth, tb6FacePlateMaterialText, tb6FloorFinish,
                    tb6FrontWallText, tb6HAndrailText, tb6Headroom, tb6KeySwitchLocation, tb6LandingDoorFinishText, tb6LiftNumbers,
                    tb6LiftRating, tb6MainCOPLocation, tb6MirrorText, tb6NumberOfCarEntrances, tb6NumberOFCOPs, tb6NumberOfLandingDoors,
                    tb6NumberOfLandings, tb6NumberOfLEDLights, tb6PitDepth, tb6RearWallText, tb6ShaftDepth, tb6ShaftWidth,
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
                    rb6StructureShaftCOncrete, rb6StructureShaftOther, rb6TrimmerBeamsNo,
                    rb6TrimmerBeamsYes, rb6VoiceAnnunciationNo, rb6VoiceAnnunciationYes
                    );
                #endregion
                #region Page 6 Word Export
                f.WordData(prefix + "AE114", tb6LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb6TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb6IndependentServiceYes, rb6IndependentNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb6LoadWeighingYes, rb6LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb6ControlerLocationText, rb6ControlerLocationOther, rb6ControlerLocationBottomLanding, rb6ControlerLocationShaft, rb6ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb6FireServiceNo, rb6FireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb6ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb6ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb6PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb6Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb6Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb6NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb6NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb6StructureShaftText, rb6StructureShaftCOncrete, rb6StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb6TrimmerBeamsYes, rb6TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb6FalseFloorYes, rb6FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb6CarLoad.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb6CarSpeed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb6CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb6CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb6CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb6LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb6LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb6NumberOfCarEntrances.Text);//number of car entraces
                if (tb6CarNote.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb6CarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb6DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb6DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb6LandingDoorFinishText, rb6LandingDoorFinishOther, rb6LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb6DoorTypeCentreOpening, rb6DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb6DoorTracksOther, rb6DoorTracksAluminium, rb6DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb6AdvancedOpeningNo, rb6AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb6DoorNudgingNo, rb6DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb6CarDoorFinishText, rb6CarDoorFinishOther, rb6CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb6CeilingFinishText, rb6CeilingFinishStainlessSteel, rb6CeilingFinishWhite, rb6CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb6FalseCeilingNo, rb6FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb6BumpRailYes, rb6BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb6FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb6FrontWallText, rb6FrontWallStainlessSteel, rb6FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb6MirrorText, rb6MirrorFullSize, rb6MirrorHalfSize, rb6MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb6HAndrailText, rb6HandrailStainlessSteel, rb6HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb6SideWallText, rb6SideWallStainlessSteel, rb6SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb6NumberOfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb6ProtectiveBlanketsYes, rb6ProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb6RearWallText, rb6RearWallOther, rb6RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb6NumberOFCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb6MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb6AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb6Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb6KeySwitchLocation.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb6COPFinishText, rb6COPFinishStainlessSteel, rb6COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb6LCDColourBlue, rb6LCDColourRed, rb6LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb6ExclusiveServiceNo, rb6ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb6RearDoorKeySwitchNo, rb6RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb6SecurityKeySwitchNo, rb6SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb6GPOInCarNo, rb6GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb6VoiceAnnunciationYes, rb6VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb6PositionIndicatorTypeSurfaceMount, rb6PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb6FacePlateMaterialText, rb6FacvePlateMaterialStainlessSteel, rb6FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb6EmergencyLoweringSystemYes, rb6EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb6OutOfServiceYes, rb6OutOfServiceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 7)
            {
                prefix = "P7";
                #region Page 7 Saving
                f.SaveTbToXML(tb7AuzCOPLocation, tb7CarDepth, tb7CarDoorFinishText,
                    tb7CarHeight, tb7CarLoad, tb7CarNotes, tb7CarSpeed, tb7CarWidth, tb7CEilingFinishText, tb7ControlerLocationText, tb7COPFinishText,
                    tb7Designations, tb7DoorHeight, tb7DoorTracksText, tb7DoorWidth, tb7FacePlateMaterialText, tb7FloorFinish, tb7FrontWallText,
                    tb7HandrailText, tb7HeadRoom, tb7KeyswitchLocation, tb7LandingDoorFinishText, tb7LiftNumbers, tb7LiftRating,
                    tb7MainCOPLocation, tb7MirrorText, tb7NumberOfCarEntrances, tb7NumberOfCOPs, tb7NumberOfLandingDoors, tb7NumberOfLandings,
                    tb7NumberOfLEDLights, tb7PitDepth, tb7RearWallText, tb7ShaftDepth, tb7ShaftWidth, tb7SideWallText,
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
                #region Page 7 Word Export

                f.WordData(prefix + "AE114", tb7LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb7TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb7IndpendentServiceYes, rb7IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb7LoadWeighingYes, rb7LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb7ControlerLocationText, rb7ControlerLocationBottomLAnding, rb7ControlerLocationOther, rb7ControlerLocationShaft, rb7ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb7FireServiceNo, rb7FireSErviceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb7ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb7ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb7PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb7HeadRoom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb7Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb7NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb7NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb7StructureShaftText, rb7StructureShaftConcrete, rb7StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb7TrimmerBeamsYes, rb7TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb7FalseFloorYes, rb7FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb7CarLoad.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb7CarSpeed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb7CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb7CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb7CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb7LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb7LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb7NumberOfCarEntrances.Text);//number of car entraces
                if (tb7CarNotes.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb7CarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb7DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb7DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb7LandingDoorFinishText, rb7LandingDoorFinishOther, rb7LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb7DoorTypeCentreOpening, rb7DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb7DoorTracksText, rb7DoorTracksAluminium, rb7DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb7AdvancedOpeningNo, rb7AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb7DoorNudgingNo, rb7DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb7CarDoorFinishText, rb7CarDoorFinishOther, rb7CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb7CEilingFinishText, rb7CeilingFinishStainlessSteel, rb7CEilingFinishWhite, rb7CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb7FalseCeilingNo, rb7FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb7BumpRailYes, rb7BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb7FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb7FrontWallText, rb7FrontWallStainlessSteel, rb7FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb7MirrorText, rb7MirrorFullSize, rb7MirrorHalfSize, rb7MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb7HandrailText, rb7HandrailStainlessSteel, rb7HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb7SideWallText, rb7SideWallStainlessSteel, rb7SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb7NumberOfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb7ProtectiveBlanketsYes, rb7ProtctiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb7RearWallText, rb7RearWallOther, rb7RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb7NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb7MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb7AuzCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb7Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb7KeyswitchLocation.Text); //key switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb7COPFinishText, rb7COPFinishStainlessSteel, rb7COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb7LCDColourBlue, rb7LCDColourRed, rb7LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb7ExclusiveServiceNo, rb7ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb7RearDoorKeySwitchNo, rb7RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb7SecurityKeySwitchNo, rb7SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb7GPOInCarNo, rb7GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb7VoiceAnnunciationYes, rb7VoiceAnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb7PositionIndicatorTypeSurfaceMount, rb7PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb7FacePlateMaterialText, rb7FacePlateMaterialStainlessSteel, rb7FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb7EmergencyLoweringSystemYes, rb7EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb7OutOfServiceYes, rb7OutOfSErviceNo));//out of service 

                #endregion
            }
            if (pageTracker >= 8)
            {
                prefix = "P8";
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
                    rb8StructureShaftConcrete, rb8StructureShaftOther, rb8TimmemrBeamsYes,
                    rb8TrimmerBeamsNo, rb8VoiceAnnunicationNo, rb8VoiceAnnunicationYes, tb8DoorNudgingNo, tb8LandingDoorFinishStainlessSteel
                    );

                f.SaveTbToXML(tb8AuxCOPLocation, tb8CarDEpth, tb8CarDoorFinishText,
                    tb8CarHeight, tb8CarWidth, tb8CeilingFinishText, tb8ControlerLocationText, tb8COPFinishText, tb8Desiginations, tb8DoorHeight,
                    tb8DoorTracksText, tb8DoorWidth, tb8FacePlateMaterialText, tb8FloorFinish, tb8FrontWallText, tb8HandrailText,
                    tb8Headroom, tb8KeyswitchLocations, tb8LandingDoorFinishText, tb8LiftCarNotes,
                    tb8LiftNumbers, tb8LiftRating, tb8Load, tb8MainCOPLocation, tb8MirrorText, tb8NumberOfCarEntrances, tb8NumberOfCOPs,
                    tb8NumberOfLandingDoors, tb8NumberOfLandings, tb8NumberofLEDLights, tb8PitDepth, tb8RearWallText,
                    tb8ShaftDepth, tb8ShaftWidth, tb8SideWallText, tb8Speed, tb8StructureShaftText, tb8Travel, tb8TypeOfLift
                    );
                #endregion
                #region Page 8 Word Export
                f.WordData(prefix + "AE114", tb8LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb8TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb8IndependentServiceYes, rb8IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb8LoadWeighingYes, rb8LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb8ControlerLocationText, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther, rb8ControlerLocationShaft, rb8ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb8FireSErviceNo, rb8FireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb8ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb8ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb8PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb8Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb8Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb8NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb8NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb8StructureShaftText, rb8StructureShaftConcrete, rb8StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb8TimmemrBeamsYes, rb8TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb8FalseFloorYes, rb8FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb8Load.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb8Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb8CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb8CarDEpth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb8CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb8LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb8LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb8NumberOfCarEntrances.Text);//number of car entraces
                if (tb8LiftCarNotes.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb8LiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb8DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb8DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb8LandingDoorFinishText, rb8LandingDoorFinishOther, tb8LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb8DoorTracksText, rb8DoorTracksAluminium, rb8DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb8AdvncedOpeningNo, rb8AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb8DoorNudgingYes, tb8DoorNudgingNo));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb8CarDoorFinishText, rb8CarDoorFinishOther, rb8CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb8CeilingFinishText, rb8CeilingFinishStainlessSteel, rb8CeilingFinishWhite, rb8CeilingFinishMirrorStainlessSTeel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb8FalseCeilingNo, rb8FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb8BumpRailYes, rb8BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb8FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb8FrontWallText, rb8FrontWallStainlessSteel, rb8FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb8MirrorText, rb8MirrorFullSize, rb8MirrorHalfSize, rb8MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb8HandrailText, rb8HandRailStainlessSteel, rb8HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb8SideWallText, rb8SideWallStainlessSteel, rb8SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb8NumberofLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb8ProtectiveBlanketsYes, rb8ProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb8RearWallText, rb8RearWallOther, rb8RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb8NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb8MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb8AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb8Desiginations.Text); // designations 
                f.WordData(prefix + "AE191", tb8KeyswitchLocations.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb8COPFinishText, rb8COPFinishStainlessSteel, rb8COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb8LCDColourBlue, rb8LCDColourRed, rb8LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb8RearDoorKeySwitchNo, rb8RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb8GPOInCarNo, rb8GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb8VoiceAnnunicationYes, rb8VoiceAnnunicationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb8PositionIndicatorTypeSurfaceMount, rb8PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb8FacePlateMaterialText, rb8FacePlateMaterialStainlessSteel, rb8FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb8EmergencyLoweringSystemYes, rb8EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb8OutOfSErviceYes, rb8OutOFServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 9)
            {
                prefix = "P9";
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
                    rb9StructureShaftConcrete, rb9StructureShaftOther, rb9TrimmerBeamsNo, rb9TrimmerBeamsYes,
                    rb9VoiceAnnunciationNo, rb9VoiceAnnunciationYes
                    );

                f.SaveTbToXML(tb9AuxCOPLocation, tb9CarDepth, tb9CarDoorFinishText, tb9CarHeight,
                    tb9CarNotes, tb9CarWidth, tb9CeilingFinishText, tb9ControlerLocationText, tb9COPFinishText, tb9Designations, tb9DoorHeight, tb9DoorTracksText,
                    tb9DoorWidth, tb9FacePlateMaterialText, tb9FloorFinish, tb9FrontWallText, tb9HandrailTexrt, tb9Headroom, tb9KeyswitchLocation,
                    tb9LandingDoorFinishText, tb9LiftNumbers, tb9LiftRating, tb9Load, tb9MainCOPLocation, tb9MirrorText, tb9NumberOFCarEntraces,
                    tb9NumberOfCOPs, tb9NumberOfLandingDoors, tb9NumberOfLandings, tb9NumberOfLEDLights, tb9PitDepth, tb9RearWallText,
                    tb9ShaftDepth, tb9ShaftWidth, tb9SideWallText, tb9Speed, tb9StructureShaftText, tb9Travel, tb9TypeOfLift
                    );
                #endregion
                #region Page 9 Word Export
                f.WordData(prefix + "AE114", tb9LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb9TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb9IndependentServiceYes, rb9IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb9LoadWeighingYes, rb9LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb9ControlerLocationText, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther, rb9ControlerLocationShaft, rb9ControlrLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb9FireServiceNo, rb9FireSErviceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb9ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb9ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb9PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb9Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb9Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb9NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb9NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb9StructureShaftText, rb9StructureShaftConcrete, rb9StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb9TrimmerBeamsYes, rb9TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb9FalseFloorYes, rb9FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb9Load.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb9Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb9CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb9CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb9CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb9LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb9LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb9NumberOFCarEntraces.Text);//number of car entraces
                if (tb9CarNotes.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb9CarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb9DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb9DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb9LandingDoorFinishText, rb9LandingDoorFinishOther, rb9LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb9DoorTracksText, rb9DoorTracksAluminium, rb9DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb9AdvancedOpeningNo, rb9AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb9DoorNudgingNo, rb9DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb9CarDoorFinishText, rb9CarDoorFinishOther, rb9CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb9CeilingFinishText, rb9CeilingFinishStainlessSteel, rb9CeilingFinishWhite, rb9CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb9FalseCeilingNo, rb9FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb9BumpRailYes, rb9BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb9FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb9FrontWallText, rb9FrontWallStainlessSteel, rb9FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb9MirrorText, rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb9HandrailTexrt, rb9HandrailStainlessSteel, rb9HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb9SideWallText, rb9SideWallStainlessSteel, rb9SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb9NumberOfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb9ProtectiveBlanketsYes, rb9ProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb9RearWallText, rb9RearWallOther, rb9RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb9NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb9MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb9AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb9Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb9KeyswitchLocation.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb9COPFinishText, rb9COPFinishStainlessSteel, rb9COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb9ExclusiveServiceNo, rb9ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb9RearDoorKeySwitchNo, rb9RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb9GPOInCarNo, rb9GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb9VoiceAnnunciationYes, rb9VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb9PositionIndicatorTypeSurfaceMount, rb9PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb9FacePlateMaterialText, rb9FacePlateMaterialStainlessSteel, rb9FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb9EmergencyLoweringSystemYes, rb9EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb9OutOfServiceYes, rb9OutOfServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 10)
            {
                prefix = "P10";
                #region Page 10 Saving
                f.SaveTbToXML(tb10AuxCOPLocation, tb10CarDepth, tb10CarDoorFinishText,
                    tb10CarHeight, tb10CarWidth, tb10CEilingFinishText, tb10ControlerLocationText, tb10COPFinishText, tb10Desigination, tb10DoorHeight,
                    tb10DoorTracksText, tb10DoorWidth, tb10FacePlateMaterialText, tb10FloorFinish, tb10FrontWallText, tb10HandrailText,
                    tb10Headroom, tb10KeyswitchLocation, tb10LandingDoorFinishText, tb10LiftCarLoad, tb10LiftCarNotes, tb10LiftNumbers,
                    tb10LiftRating, tb10MainCOPLocation, tb10MirrorText, tb10NumberofCarEntrances, tb10NumberOfCOPs, tb10NumberofLandingDoors,
                    tb10NumberofLandings, tb10NumberOfLEDLIghts, tb10PitDepth, tb10RearWallText, tb10ShaftDepth, tb10ShaftWidth,
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
                #region Page 10 Word Export
                f.WordData(prefix + "AE114", tb10LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb10TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb10IndependentServiceYes, rb10IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb10LoadWeighingYes, rb10LoadWEighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb10ControlerLocationText, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther, rb10ControlerLocationShaft, rb10ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb10FireSERviceNo, rb10FireSErviceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb10ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb10ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb10PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb10Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb10Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb10NumberofLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb10NumberofLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb10StructureShaftText, rb10StructureShaftConcrete, rb10StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb10TimmerbeamsYes, rb10TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb10FalseFloorYes, rb10FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb10LiftCarLoad.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb10Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb10CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb10CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb10CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb10LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb10LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb10NumberofCarEntrances.Text);//number of car entraces
                if (tb10LiftCarNotes.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb10LiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb10DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb10DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb10LandingDoorFinishText, rb10LAndingDoorFinishOtherr, rb10LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb10DoorTracksText, rb10DoorTracksAluminium, rb10DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb10AdvancedOpeningNo, rb10AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb10DoorNudgingNo, rb10DoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb10CarDoorFinishText, rb10CarDoorFinishOther, rb10CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb10CEilingFinishText, rb10CeilingFinishStainlessSteel, rb10CeilingFinishWhite, rb10CEilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb10FalseCeilingNo, rb10FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb10BumpRaidYes, rb10BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb10FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb10FrontWallText, rb10FrontWallStainlessSteel, rb10FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb10MirrorText, rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb10HandrailText, rb10HandrailStainlessSteel, rb10HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb10SideWallText, rb10SideWallStainlesSteel, rb10SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb10NumberOfLEDLIghts.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb10ProtectiveBlanketYes, rb10ProtectiveBlanketNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb10RearWallText, rb10RearWallOther, rb10RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb10NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb10MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb10AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb10Desigination.Text); // designations 
                f.WordData(prefix + "AE191", tb10KeyswitchLocation.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb10COPFinishText, rb10COPFinishStainlessSteel, rb10COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb10LCDColourBlue, rb10LCDColourRed, rb10LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb10RearDoorKeySwitchNo, rb10RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb10SecurityKeySwitchNo, rb10SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb10GPOInCarNo, rb10GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb10VoiceAnnunciationYes, rb10VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb10PositionIndicatorTypeSurfaceMount, rb10PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb10FacePlateMaterialText, rb10FacePlateMaterialStainlessSteel, rb10FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb10EmergencyLoweringSystemYes, rb10EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb10OutOFServiceYes, rb10OutOfServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 11)
            {
                prefix = "P11";
                #region Page 11 Saving
                f.SaveTbToXML(tb11AuxCOPLocation, tb11CarDepth, tb11CarDoorFinishText,
                    tb11CarHeight, tb11CarWidth, tb11CeilingFinishText, tb11ControlerLocationText, tb11COPFinishText, tb11Designations, tb11DoorHeight,
                    tb11DoorTracksText, tb11DoorWidth, tb11FaceplateMaterialText, tb11FloorFinish, tb11FrontWallText, tb11HandrailText, tb11Headroom,
                    tb11KeyswitchLocation, tb11LandingDoorFInishOther, tb11LiftCarLoad, tb11LiftCarNote, tb11LiftRating, tb11MainCOPLocation,
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
                #region Page 11 Word Export
                f.WordData(prefix + "AE114", rb11LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb11TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb11IndependentServiceYes, rb11IndependentSErviceNO));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb11LoadWeighingYes, rb11LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb11ControlerLocationText, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther, rb11ControlerLocationShaft, tb11ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb11FireServiceNo, rb11FireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb11ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb11ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb11PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb11Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb11Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb11NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb11NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(rb11StructureShaftText, rb11StructureShaftConcrete, rb11StrructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb11TrimmerBeamsYes, rb11TrimmerBeamNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb11FalseFloorYes, rb11FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb11LiftCarLoad.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb11Speed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb11CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb11CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb11CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb11LiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb11LiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb11NumberofCarEntrances.Text);//number of car entraces
                if (tb11LiftCarNote.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb11LiftCarNote.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb11DoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb11DoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb11LandingDoorFInishOther, rb11LandingDoorFinishOther, rb11LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb11DoorTypeCentreOpening, rb11DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb11DoorTracksText, rb11DoorTracksAluminium, rb11DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb11AdvancedOpeningNo, rb11AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb11DoorNudgingYes, rb11DoorNudgingNo));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb11CarDoorFinishText, rb11CarDoorFinishOther, rb11CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb11CeilingFinishText, rb11CeilingFinishStainlessSteel, rb11CeilingFinishWhite, rb11CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb11FalseCeilingNo, rb11FalseCEilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb11BumpRailYes, rb11BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb11FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb11FrontWallText, rb11FrontWallStainlessSteel, rb11FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb11MirrorText, rb11MirrorFullSize, rb11MirrorHalfSize, rb11MirrorOther));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb11HandrailText, rb11MirrorHalfSize, rb11HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb11SideWallText, rb11SideWallStainlessSteel, rb11SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb11NumberOfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb11ProtectiveBlanketsYes, rb11ProtectiveBlanketNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb11RearWallText, rb11RearWallOther, rb11RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb11NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb11MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb11AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb11Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb11KeyswitchLocation.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb11COPFinishText, rb11COPFinishStainlessSteel, rb11COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb11ExclusiveServiceNo, rb11ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb11RearDoorKeySwitchNo, rb11RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb11GPOInCarNo, rb11GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb11VoiceAnnunciationYes, rb11VoiceAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb11PositionIndicatorTypeSurfaceMount, rb11PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb11FaceplateMaterialText, rb11FacePlateMaterialStainlessSTeel, rb11FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb11EmergencyLoweringSystemYes, rb11EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb11OutOfSErviceYes, rb11OutOfServiceNo));//out of service 
                #endregion
            }
            if (pageTracker >= 12)
            {
                prefix = "P12";
                #region Page 12 Saving
                f.SaveTbToXML(tb12AuxCOPLocation, tb12CarDepth, tb12CarDoorFinishText,
                    tb12CarHeight, tb12CarLiftRating, tb12CarLoad, tb12CarNumberOfCarEntrances, tb12CarSpeed, tb12CarWidth, tb12CeilingFinishText,
                    tb12ControlerLocationText, tb12COPFinishText, tb12Designations, tb12DoorTracksText, tb12FacePlateMaterialText,
                    tb12FloorFinish, tb12FrontWallText, tb12HandrailText, tb12Headroom, tb12KeyswitchLocation, tb12LandingDoorFinishText,
                    tb12LandingDoorHeight, tb12LandingDoorWidth, tb12LiftCarNotes, tb12LiftNumbers, tb12MainCOPLocation, tb12MirrorText,
                    tb12NumberOfCOPs, tb12NumberOfLandingDoors, tb12NumberOfLandings, tb12NumberOfLEDLights, tb12PitDepth,
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
                    rb12StructureShaftConcrete, rb12StructureShaftOther, rb12TrimmerBeamsNo,
                    rb12TrimmerBeamsYes, rb12VoicAnnunciationNo, rb12VoiceAnnuniationYes
                    );
                #endregion
                #region Page 12 Word Export
                f.WordData(prefix + "AE114", tb12LiftNumbers.Text);//lift number
                f.WordData(prefix + "AE115", tb12TypeOfLift.Text);//type of lift
                f.WordData(prefix + "AE215", "Full Collective"); //control type, not changable 
                f.WordData(prefix + "AE116", f.RadioButtonToAsteriskHandeler(rb12IndependentServiceYes, rb12IndependentServiceNo));//independent service
                f.WordData(prefix + "AE117", f.RadioButtonToAsteriskHandeler(rb12LoadWeighingYes, rb12LoadWeighingNo));//load weighing
                f.WordData(prefix + "AE118", f.RadioButtonHandeler(tb12ControlerLocationText, rb12ControlerLocationBottomLanding, rb12ControlerLocationOther, rb12ControlerLocationShaft, rb12ControlerLocationTopLanding));//controler location
                f.WordData(prefix + "AE120", f.RadioButtonHandeler(null, rb12FireServiceNo, rb12FireServiceYes));//fire service
                f.WordData(prefix + "AE123", f.MeasureStringChecker(tb12ShaftWidth.Text, "mm"));// shaft width
                f.WordData(prefix + "AE124", f.MeasureStringChecker(tb12ShaftDepth.Text, "mm"));//shaft depth
                f.WordData(prefix + "AE125", f.MeasureStringChecker(tb12PitDepth.Text, "mm"));//pit depth
                f.WordData(prefix + "AE126", f.MeasureStringChecker(tb12Headroom.Text, "mm"));//headroom
                f.WordData(prefix + "AE127", f.MeasureStringChecker(tb12Travel.Text, "mm"));//travel
                f.WordData(prefix + "AE128", tb12NumberOfLandings.Text);// number of landings
                f.WordData(prefix + "AE129", tb12NumberOfLandingDoors.Text);//number of landing doors 
                f.WordData(prefix + "AE130", f.RadioButtonHandeler(tb12StructureShaftText, rb12StructureShaftConcrete, rb12StructureShaftOther)); //structure shaft 
                f.WordData(prefix + "AE132", f.RadioButtonToAsteriskHandeler(rb12TrimmerBeamsYes, rb12TrimmerBeamsNo));//trimmer beams
                f.WordData(prefix + "AE133", f.RadioButtonToAsteriskHandeler(rb12FalseFloorYes, rb12FalseFloorNo));//false floor
                f.WordData(prefix + "AE134", f.MeasureStringChecker(tb12CarLoad.Text, "kg")); //load
                f.WordData(prefix + "AE135", f.MeasureStringChecker(tb12CarSpeed.Text, "mps"));//speed
                f.WordData(prefix + "AE136", f.MeasureStringChecker(tb12CarWidth.Text, "mm")); // width
                f.WordData(prefix + "AE137", f.MeasureStringChecker(tb12CarDepth.Text, "mm"));//depth
                f.WordData(prefix + "AE138", f.MeasureStringChecker(tb12CarHeight.Text, "mm"));//height
                f.WordData(prefix + "AE139", f.MeasureStringChecker(tb12CarLiftRating.Text, "passengers"));//classification rating
                f.WordData(prefix + "AE113", f.MeasureStringChecker(tb12CarLiftRating.Text, "passenger"));//classification rating
                f.WordData(prefix + "AE140", tb12CarNumberOfCarEntrances.Text);//number of car entraces
                if (tb12LiftCarNotes.Text != "")
                {
                    f.WordData(prefix + "AE142", "NOTE: " + tb12LiftCarNotes.Text + Environment.NewLine);//notes
                }
                else
                {
                    f.WordData(prefix + "AE142", "");//notes
                }
                f.WordData(prefix + "AE143", f.MeasureStringChecker(tb12LandingDoorWidth.Text, "mm"));//door width 
                f.WordData(prefix + "AE144", f.MeasureStringChecker(tb12LandingDoorHeight.Text, "mm")); //door height 
                f.WordData(prefix + "AE147", f.RadioButtonHandeler(tb12LandingDoorFinishText, rb12LandingDoorFinishOther, rb12LandingDoorFinishStainlessSteel));//landing door finish
                f.WordData(prefix + "AE148", f.RadioButtonHandeler(null, rb12DoorTypeCentreOpening, rb12DoorTypeSideOpening));//door type
                f.WordData(prefix + "AE150", f.RadioButtonHandeler(tb12DoorTracksText, rb12DoorTracksAluminium, rb12DoorTracksOther));// door tracks 
                f.WordData(prefix + "AE151", f.RadioButtonHandeler(null, rb12AdvancedOpeningNo, rb12AdvancedOpeningYes));//advanced opening
                f.WordData(prefix + "AE152", f.RadioButtonHandeler(null, rb12LandingDoorNudgingNo, rb12LandingDoorNudgingYes));//door nudging 
                f.WordData(prefix + "AE155", f.RadioButtonHandeler(tb12CarDoorFinishText, rb12CarDoorFinishOther, rb12CarDoorFinishStainlessSteel));//car door finish
                f.WordData(prefix + "AE156", f.RadioButtonHandeler(tb12CeilingFinishText, rb12CeilingFinishStainlessSteel, rb12CeilingFinishWhite, rb12CeilingFinishMirrorStainlessSteel, rbCeilingFinishOther));//ceiling finish
                f.WordData(prefix + "AE157", f.RadioButtonHandeler(null, rb12FalseCeilingNo, rb12FalseCeilingYes));//false ceiling
                f.WordData(prefix + "AE158", f.RadioButtonHandeler(null, rb12BumpRailYes, rb12BumpRailNo));//bump rail
                f.WordData(prefix + "AE159", tb12FloorFinish.Text);//floor
                f.WordData(prefix + "AE160", f.RadioButtonHandeler(tb12FrontWallText, rb12FrontWallStainlessSteel, rb12FrontWallOther));//front wall
                f.WordData(prefix + "AE161", f.RadioButtonHandeler(tb12MirrorText, rb12MirrorFullSize, rb12MirrorHalfSize, rb12MirrorOTher));//mirror
                f.WordData(prefix + "AE162", f.RadioButtonHandeler(tb12HandrailText, rb12HandrailStainlessSTeel, rb12HandrailOther));//handrail
                f.WordData(prefix + "AE163", @"Natural & Mechanical");// ventelation fan
                f.WordData(prefix + "AE164", f.RadioButtonHandeler(tb12SideWallText, rb12SideWallStainlessSteel, rb12SideWallOther)); //side wall 
                f.WordData(prefix + "AE165", tb12NumberOfLEDLights.Text + " LED Lights"); // lighting 
                f.WordData(prefix + "AE167", f.RadioButtonToAsteriskHandeler(rb12ProectiveBlanketsYes, rb12ProtectiveBlanketsNo)); // protective blankets 
                f.WordData(prefix + "AE216", f.RadioButtonHandeler(tb12RearWallText, rb12RearWallOther, rb12RearWallStainlessSteel)); //  rear wall
                f.WordData(prefix + "AE168", tb12NumberOfCOPs.Text); // number of COPS
                f.WordData(prefix + "AE169", tb12MainCOPLocation.Text);// main COP location
                f.WordData(prefix + "AE170", tb12AuxCOPLocation.Text);//aux cop location
                f.WordData(prefix + "AE171", tb12Designations.Text); // designations 
                f.WordData(prefix + "AE191", tb12KeyswitchLocation.Text); //keyt switch location
                f.WordData(prefix + "AE172", f.RadioButtonHandeler(tb12COPFinishText, rb12COPFinishStainlessSteel, rb12COPFinishOther));// COP finish
                f.WordData(prefix + "AE173", "Dual illumination buttons with gong");// button type 
                f.WordData(prefix + "AE174", f.RadioButtonHandeler(null, rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite));// LCD colour
                f.WordData(prefix + "AE183", f.RadioButtonHandeler(null, rb12ExclusiveServiceNo, rb12ExclusiveServiceYes));//exclusive service 
                f.WordData(prefix + "AE184", f.RadioButtonHandeler(null, rb12RearDoorKeySwitchNo, rb12RearDoorKeySwitchYes));// rear door kew switch 
                f.WordData(prefix + "AE186", f.RadioButtonHandeler(null, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes));//security key switch 
                f.WordData(prefix + "AE187", f.RadioButtonHandeler(null, rb12GPOInCarNo, rb12GPOInCarYes));//GPO in car
                f.WordData(prefix + "AE189", f.RadioButtonToAsteriskHandeler(rb12VoiceAnnuniationYes, rb12VoicAnnunciationNo));//voice annunciation 
                f.WordData(prefix + "AE190", f.RadioButtonHandeler(null, rb12PositionIndicatorTypeSurfaceMount, rb12PositionIndicatorTypeFlushMount));// position indicaor type 
                f.WordData(prefix + "AE192", f.RadioButtonHandeler(tb12FacePlateMaterialText, rb12FacePlateMaterialStainlessSteel, rb12FacePlateMaterialOther));//face plate material 
                f.WordData(prefix + "AE193", "Dual illumination buttons with gong");//button type
                f.WordData(prefix + "AE209", f.RadioButtonHandeler(null, rb12EmergencyLoweringSystemYes, rb12EmergencyLoweringSystemNo));// emergency lowering system 
                f.WordData(prefix + "AE210", f.RadioButtonToAsteriskHandeler(rb12OutOfServiceYes, rb12OutOfServiceNo));//out of service 
                #endregion
            }

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
            FillPageWithMainData(2, ref page2Opened, f.loadingPreviousData);
            OpenInfoPage(1);
        }

        private void btPanel3_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(3, ref page3Opened, f.loadingPreviousData);
            OpenInfoPage(2);
        }

        private void btPanel4_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(4, ref page4Opened, f.loadingPreviousData);
            OpenInfoPage(3);
        }

        private void btPanel5_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(5, ref page5Opened, f.loadingPreviousData);
            OpenInfoPage(4);
        }

        private void btPanel6_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(6, ref page6Opened, f.loadingPreviousData);
            OpenInfoPage(5);
        }

        private void btPanel7_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(7, ref page7Opened, f.loadingPreviousData);
            OpenInfoPage(6);
        }

        private void btPanel8_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(8, ref page8Opened, f.loadingPreviousData);
            OpenInfoPage(7);
        }

        private void btPanel9_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(9, ref page9Opened, f.loadingPreviousData);
            OpenInfoPage(8);
        }

        private void btPanel10_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(10, ref page10Opened, f.loadingPreviousData);
            OpenInfoPage(9);
        }

        private void btPanel11_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(11, ref page11Opened, f.loadingPreviousData);
            OpenInfoPage(10);
        }

        private void btPanel12_Click(object sender, EventArgs e)
        {
            FillPageWithMainData(12, ref page12Opened, f.loadingPreviousData);
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

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            //
        }
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

        #region Fill data from 1st page into additional pages
        private void FillPageWithMainData(int pageToBeFilled, ref bool pageOpenedPreviously, bool previousQuoteWasLoaded)
        {
            if (pageOpenedPreviously || previousQuoteWasLoaded)
            {
                return;
            }

            pageToBeFilled--;
            pageOpenedPreviously = true;

            #region Page Data Vars
            Object[][] pageObjects = new object[][]
            {

           new Object[] { tbAuxCOPLocation, tbDepth, tbCarDoorFinish, tbHeight, tbLoad, tbwidth, tbCeilingFinish,
                tbControlerLocation, tbCOPFinish, tbDesignations, tbDoorHeight, tbDoorTracks, tbDoorWidth, tbFacePlateMaterial,
                tbFloorFinish, tbFrontWall, tbHandrail, tbHeadroom, tbKeyswitchLocation, tbLandingDoorFinish, tbLiftNumbers,
                tbLiftRating, tbMainCOPLocation, tbMirror, tbLiftCarNotes, tbNumofCarEntrances, tbNumberOfCOPS, tbNumofLandingDoors,
                tbNumofLandings, tbNumOfLEDLights, tbPitDepth, tbRearWall, tbShaftDepth, tbShaftWidth, tbSideWall, tbSpeed,
                tbStructureShaft, tbTravel, tbTypeofLift, rbAdvancedOpeningNo, rbAdvancedOpeningYes, rbBumpRailNo, rbBumpRailYes,
                rbCarDoorFinishOther, rbCarDoorFInishBrushedStainlessSteel, rbCeilingFinishMirrorStainlessSteel, rbCeilingFinishOther,
                rbCeilingFinishBrushedStasinlessSteel, rbCeilingFinishWhite, rbControlerLocationBottomLanding, rbControlerLocationOther,
                rbControlerlocationShaft, rbControlerLoactionTopLanding, rbCOPFinishOther, rbCOPFinishSatinStainlessSteel, rbDoorNudgingNo,
                rbDoorNudgingYes, rbDoorTracksAnodisedAluminium, rbDoorTracksOther, rbDoorTypeCentreOpening, rbDoorTypeSideOpening,
                rbEmergencyLoweringSystemNo, rbEmergencyLoweringSystemYes, rbExclusiveServiceNo, rbExclusiveServiceYes,
                rbFacePlateMaterialOther, rbFacePlateMaterialSatinStainlessSteel, rbFalseCeilingNo, rbFalseCeilingYes, rbFalseFloorNo, rbFalseFloorYes,
                rbFireServiceNo, rbFireServiceYes, tbFrontWallOther, rbFrontWallBrushedStainlessSteel, rbGPOInCarNo, rbGPOInCarYes, rbHandrailOther,
                rbHandrailBrushedStainlessSTeel, rbIndependentServiceNo, rbIndependentServiceYes, rbLandingDoorFinishOther,
                rbLandingDoorFinishStainlessSteel, rbLEDColourBlue, rbLEDColourRed, rbLEDColourWhite, rbLoadWeighingNo, rbLoadWeighingYes,
                rbMirrorFullSize, rbMirrorHalfSize, rbMirrorOther, rbOutofServiceNo, rbOutofServiceYes, rbPositionIndicatorTypeFlushMount,
                rbPositionIndicatorTypeSurfaceMount, rbProtectiveBlanketsNo, rbProtectriveBlanketsYes, rbRearDoorKeySwitchNo,
                rbRearDoorKeySwitchYes, rbRearWallOther, rbRearWallBrushedStainlessSteel, rbSecurityKeySwitchNo, rbSecurityKeySwitchYes,
                rbSideWallOther, rbSideWallBrushedStainlessSteel, rbStructureShaftConcrete, rbStructureShaftOther, rbTrimmerBeamsNo,
                rbTrimmerBeamsYes, rbVoiceAnnunciationNo, rbVoiceAnnunciationYes },

                 new Object[]  { tb2AuxCOPLocation, tb2CarDepth, tb2CarDoorFinishText, tb2CarHeight, tb2CarLoad, tb2CarWidth, tb2CeilingFinishText,
                    tb2ControlerLocationText, tb2COPFinishText, tb2Designations, tb2DoorHeight, tb2DoorTracksText, tb2DoorWidth, tb2FacePlateMaterialText,
                    tb2FloorFinish, tb2FrontWallText, tb2HandrailText, tb2Headroom, tb2KeyswitchLocation, tb2LandingDoorFinishText, tb2LiftNumbers,
                    tb2LiftRating, tb2MainCOPLocation, tb2MirrorText, tb2Note, tb2NumberOfCarEntrances, tb2NumberOfCOPs, tb2NumberOfLAndingDoors,
                    tb2NumberOfLandings, tb2NumberofLEDLights, tb2PitDepth, tb2RearWallText, tb2ShaftDepth, tb2ShaftWidth, tb2SideWallText, tb2Speed,
                    tb2StructureShaftText, tb2Travel, tb2TypeOfLift, rb2AdvancedOpeningNo, rb2AdvancedOpeningYes, rb2BumpRailNo, rb2BumpRailYes,
                    rb2CarDoorFinishOther, rb2CarDoorFinishStainlessSteel, rb2CeilingFinishMirrorStainlessSteel, rb2CeilingFinishOther,
                    rb2CeilingFinishStainlessSteel, rb2CeilingFinishWhite, rb2ControlerLocationBottomLanding, rb2ControlerLocationOther,
                    rb2ControlerLocationShaft, rb2ControlerLocationTopLanding, rb2COPFinishOther, rb2COPFinishStainlessSTeell, rb2DoorNudgingNo,
                    rb2DoorNudgingYes, rb2DoorTracksAluminium, rb2DoorTracksOther, rb2DoorTypeCEntreOpening, rb2DoorTypeSideOpening,
                    rb2EmergemncyLoweringSystemNo, rb2EmergencyLoweringSystemYes, rb2ExclusiveServiceNo, rb2ExclusiveServiceYes,
                    rb2FacePlateMaterialOther, rb2FacePlateMaterialStainlessSteel, rb2FalseCeilingNo, rb2FalseCeilingYes, rb2FalseFloorNo, rb2FalseFloorYes,
                    rb2FireSErviceNo, rb2FireSErviceYes, rb2FrontWallOther, rb2FrontWallStainlessSteel, rb2GPOInCarNo, rb2GPOInCarYes, rb2HandrailOther,
                    rb2HandRailStainlessSteel, rb2IndependentServiceNo, rb2IndependentServiceYes, rb2LandingDoorFinishOther,
                    rb2LandingDoorFinishStainlessSteel, rb2LCDColourBlue, rb2LCDColourRed, rb2LCDColourWhite, rb2LoadWeighingNo, rb2LoadWeighingYes,
                    rb2MirrorFullSize, rb2MirrorHalfSize, rb2MirrorOther, rb2OutOfServiceNo, rb2OutOfServiceYes, rb2PositionIndicatorTypeFlushMount,
                    rb2PositionIndicatorTypeSurfaceMount, rb2ProtectiveBlanketsNo, rb2ProtectiveBlanketsYes, rb2RearDoorKeySwitchNo,
                    rb2RearDoorKeySwitchYes, rb2RearWallOther, rb2RearWallStainlessSteel, rb2SecurityKeySwitchNo, rb2SecurityKeySwitchYes,
                    rb2SideWallOther, rb2SideWallStainlessSteel, rb2StructureShaftConcrete, rb2StructureShaftOther, rb2TrimmerBeamsNo,
                    rb2TrimmerBeamsYes, rb2VoiceAnnunciationNo, rb2VoiceAnnunciationYes },

           new Object[] { tb3AuxCOPLocation, tb3CarDepth, tb3CarDoorFinishText, tb3CarHeight, tb3Load, tb3CarWidth, tb3CEilingFinishText,
                tb3ControlerLocationText, tb3COPFinishText, tb3Designations, tb3DoorHeight, tb3DoorTracksText, tb3DoorWidth, tb3FacePlaterMaterialText,
                tb3FloorFinish, tb3FrontWallText, tb3HandrailText, tb3HeadRoom, tb3KeyswitchLocation, tb3LandingDoorFinishText, tb3LiftNumbers,
                tb3LiftRating, tb3MainCOPLocation, tb3MirrorText, tb3CarNote, tb3NumberOfCarEntrances, tb3NumberOfCOPs, tb3NumberOfLandingDoors,
                tb3NumberOfLandings, tb3NumberOfLEDLights, tb3PitDepth, tb3RearWallText, tb3ShaftDepth, tb3ShaftWidth, tb3SideWallText, tb3Speed,
                tb3StructureShaftText, tb3Travel, tb3TypeOfLift, rb3AdvancedOpeningNo, rb3AdvancedOpeningYes, rb3BumpRailNo, rb3BumpRailYes,
                rb3CarDoorFinishOther, rb3CarDoorFinishStainlessSteel, rb3MirrorStainlessSteel, rb3CeilingFinishOther,
                rb3CeilingFinishStainlessSteel, rb3CeilingFinishWhite, rb3ControleRLocationBottomLanding, rb3ControlerLocationOther,
                rb3ControlerLocationShaft, rb3ControlerLocationTopLanding, rb3COPFinishOther, rb3COPFinishStainlessSteel, rb3DoorNudgingNo,
                rb3DoorNudgingYes, rb3DoorTracksAluminium, rb3DoorTracksOther, rb3DoorTypeCentreOpening, rb3DoorTypeSideOpening,
                rb3EmergencyLoweringSystemNo, rb3EmergencyLoweringSystemYes, rb3ExclusiveServiceNo, rb3ExclusiveServiceYes,
                rb3FacePlateMaterialOther, rb3FacePlateMaterialStainlessSteel, rb3FalseCeilingNo, rb3FalseCeilingYes, rb3FalseFloorNo, rb3FalseFloorYes,
                rb3FireServieNo, rb3FireServiceYes, rb3FrontWallOther, rb3FrontWallStainlessSteel, rb3GPOInCarNo, rb3GPOInCarYes, rb3HandrailOther,
                rb3HandrailStainlessSteel, rb3IndependentServiceNo, rb3IndependentServiceYes, rb3landingDoorFinishOther,
                rb3LandingDoorFinishStainlessSteel, rb3LCDColourBlue, rb3LCDColourRed, rb3LCDColourWhite, rb3LoadWeighingNo, rb3LoadWeighingYes,
                rb3MirrorFullSize, rb3MirrorHalfSize, rb3MirrorOther, rb3OutOfServiceNo, rb3OutOfSErviceYes, rb3PositionIndicatorTypeFlushMount,
                rb3PositionIndicatorTypeSurfaceMount, rb3ProtectiveBlanketsNo, rb3ProtectiveBlanketsYes, rb3RearDoorKeySwitchNo,
                rb3RearDoorKeySwitchYes, rb3RearWallOther, rb3RearWallStainlessSteel, rb3SecurityKeySwitchNo, rb3SecurityKeySwitchYes,
                rb3SideWallOther, rb3SideWallStainlessSteel, rb3StructureShaftConcrete, rb3StructureShaftOther, rb3TrimmerBeamsNo,
                rb3TrimmerBeamsYes, rb3VoiceAnnunciationNo, rb3VoiceAnnunciationYes },

           new Object[] { tb4AuxCOPLocation, tb4CarDepth, tb4CarDoorFinish, tb4CarHeight, tb4Load, tb4CarWidth, tb4CeilingFinishText,
                tb4ControlerLocationText, tb4COPFinishText, tb4Designations, tb4DoorHeight, tb4DoorTracksText, tb4DoorWidth, tb4FacePlateMaterialText,
                tb4FloorFinish, tb4FrontWallText, tb4HandrailText, tb4Headroom, tb4KeyswitchLocations, tb4LandingDoorFinishText, tb4LiftNumbers,
                tb4LiftRating, tb4MainCOPLocation, tb4MirrorText, tb4CarNote, tb4NumberOfCarEntrances, tb4NumberOfCOPs, tb4NumberOfLandingDoors,
                tb4NumberOfLandings, tb4NumbeROfLEDLights, tb4PitDepth, tb4RearWallText, tb4ShaftDepth, tb4ShaftWidth, tb4SideWallText, tb4Speed,
                tb4StructureShaftText, tb4Travel, tb4TypeOfLift, rb4AdvancedOpeningNo, rb4AdvancedOpeningYes, rb4BumpRailNo, rb4BumpRailYes,
                rb4CarDoorFinishOther, rb4CarDoorFinishStainlessSteel, rb4CeilingFinishMirrorStainlessSteel, rb4CeilingFinishOther,
                rb4CeilingFinishStainlessSteel, rb4CeilingFinishWhite, rb4ControlerLocationBottomLanding, rb4ControlerLocationOther,
                rb4ControlerLocationShaft, rb4ControelrLocationTopLanding, rb4COPFinishOther, rb4COPFinishStainlessSteel, rb4DoorNudgingNo,
                rb4DoorNudgingYes, rb4DoorTracksAluminium, rb4DoorTracksOther, rb4DoorTypeCentreOpening, rb4DoorTypeSideOpening,
                rb4EmergencyLoweringSystemNo, rb4EmergencyLoweringSystemYes, rb4ExclusiveServiceNo, rb4ExclusiveServiceYes,
                rb4FacePlateMaterialOther, rb4FacePlateMaterialStainlessSteel, rb4FalseCeilingNO, rb4FalseCeilingYes, rb4FalseFloorNo, rb4FalseFloorYes,
                rb4FireSErviceNo, rb4FireServiceYes, rb4FrotnWallOther, rb4FrontWallStainlessSteel, rb4GPOInCarNo, rb4GPOInCarYes, rb4HandrailOther,
                rb4HandRailStainlesSteel, IndependentServiceNo, rb4IndependentServiceYes, rb4LandingDoorFinishOther,
                rb4LandingDoorFinishStainlessSteel, rb4LCDColourBlue, rb4LCDColourRed, rb4LCDColourWhite, rb4LoadWeighingNo, rb4LoadWeighingYes,
                rb4MirrorFullSizer, rb4MirrorHalfSize, rb4MirrorOtther, rb4OutOfServiceNo, rb4OutOfServiceYes, rb4PositionIndicatorTypeFlushMount,
                rb4PositionIndicatorTypeSurfaceMount, rb4ProtetiveBlanketsNo, rb4ProtectiveBlanketsYes, rb4RearDoorKeySwitchNo,
                rb4RearDoorKeySwitchYes, rb4RearWallOther, rb4RearWallStainlessSteel, rb4SecurityKeySwitchNo, rb4SecurityKeySwitchYes,
                rb4SideWallOther, rb4SideWallStainlessSteel, rb4StructureShaftConcrete, rb4StructureShaftOther, rb4TrimmerBeamsNo,
                rb4TrimmerBeamsYes, rb4VoiceAnnunciationNo, rb4VoiceAnnunciationYes },

           new Object[] { tb5AuxCOPLocation, tb5CarDepth, tb5CarDoorFinishText, tb5CaRHeight, tb5Load, tb5CarWidth, tb5CeilingFinishText,
                tb5ControlerLocationText, tb5COPFinishText, tb5Designations, tb5DoorHeight, tb5DoorTRacksText, tb5DoorWidth, tb5FacePlateMaterialText,
                tb5FloorFinish, tb5FrontWallText, tb5HandrailText, tb5Headroom, tb5KetyswitchLocation, tb5LandingDoorFinishText, tb5LiftNumbers,
                tb5LiftRating, tb5MainCOPLocation, tb5MirrorText, tb5CarNote, tb5NumberOfCarEntrances, tb5NumberOfCOPs, tb5NumberOfLandingDoors,
                tb5NumberOfLandings, tb5NumberOIfLEDLights, tb5PitDepth, tb5RearWallText, tb5ShaftDEpth, tb5ShaftWidth, tb5SideWallText, tb5Speed,
                tb5StructureShaftText, tb5Travel, tb5TypeOfLift, rb5AdvancedOpeningNo, rb5AdvancedOpeningYes, rb5BumpRailNo, rb5BumpRailYes,
                rb5CarDoorFinishOther, rb5CarDoorFinishStainlessSteel, rb5CeilingFinishMirrorStainlessSTeel, rb5CeilingFinishOther,
                rb5CeilingFinishStainlessSteel, rb5CeilingFinishWhite, rb5ControlerLocationBottomLanding, rb5ControlerLocationOther,
                rb5ControlerLocationShaft, rb5ControlerLocationTopLanding, rb5COPFinishOther, rb5COPFinishStainlessSteel, rb5DoorNudgingNo,
                rb5DoorNudgingYes, rb5DoorTracksAluminium, rb5DoorTracksOther, rb5DoorTypeCentreOpening, rb5DoorTypeSideOpening,
                rb5EmergencyLoweringSystemNo, rb5EmergencyLoweringSystemYes, rb5ExclusiveServiceNo, rb5ExclusiveServiceYes,
                rb5FacePlateMAterialOtjer, rb5FacePlateMaterialStainlessSteel, rb5FalseCeilingNo, rb5FalseCeilingYes, rb5FalseFloorNo, rb5FalseFloorYes,
                rb5FireSErviceNo, rb5FireServiceYes, rb5FrontWallOther, rb5FrontWallStainlessSteel, rb5GPOInCarNo, rb5GPOInCarYes, rb5HandRailOther,
                rb5HandRailStainlessSteel, rb5IndependentServiceNo, rb5IndependentServiceYes, rb5LAndingDoorFinishOther,
                rb5LandingDoorFinishStainlessSteel, rb5LCDColourBlue, rb5LCDColourRed, rb5LCDColoiurWHite, rb5LoadWeighingNo, rb5LoadWeighingYes,
                rb5MirrorFullSize, rb5MirrorHalfSize, rb5MirrorOther, rb5OutOfServiceNo, rb5OutOfServiceYes, rb5PostionIndicatorTypeFlushMount,
                rb5PositionIndicatorTypeSurfaceMount, rb5ProtectiveBlanketsNo, rb5ProtectiveBlanketsYes, rb5RearDoorKeySwitchNo,
                rb5RearDoorKeySwitchYes, rb5RearWallOther, rb5RearWallStainlesSteel, rb5SecurityKeySwitchNo, rb5SecurityServiceYes,
                rb5SideWallOther, rb5SideWallStainlessSTeel, rb5StructureShaftConcrete, rb5StructureShaftOther, rb5TrimmerBeamsNo,
                rb5TrimmerBeamsYes, rb5VoiceAnnunciationNo, rb5VoiceAnnunciationYes },

           new Object[] { tb6AuxCOPLocation, tb6CarDepth, tb6CarDoorFinishText, tb6CarHeight, tb6CarLoad, tb6CarWidth, tb6CeilingFinishText,
                tb6ControlerLocationText, tb6COPFinishText, tb6Designations, tb6DoorHeight, tb6DoorTracksOther, tb6DoorWidth, tb6FacePlateMaterialText,
                tb6FloorFinish, tb6FrontWallText, tb6HAndrailText, tb6Headroom, tb6KeySwitchLocation, tb6LandingDoorFinishText, tb6LiftNumbers,
                tb6LiftRating, tb6MainCOPLocation, tb6MirrorText, tb6CarNote, tb6NumberOfCarEntrances, tb6NumberOFCOPs, tb6NumberOfLandingDoors,
                tb6NumberOfLandings, tb6NumberOfLEDLights, tb6PitDepth, tb6RearWallText, tb6ShaftDepth, tb6ShaftWidth, tb6SideWallText, tb6CarSpeed,
                tb6StructureShaftText, tb6Travel, tb6TypeOfLift, rb6AdvancedOpeningNo, rb6AdvancedOpeningYes, rb6BumpRailNo, rb6BumpRailYes,
                rb6CarDoorFinishOther, rb6CarDoorFinishStainlessSteel, rb6CeilingFinishMirrorStainlessSteel, rb6CeilingFinishOther,
                rb6CeilingFinishStainlessSteel, rb6CeilingFinishWhite, rb6ControlerLocationBottomLanding, rb6ControlerLocationOther,
                rb6ControlerLocationShaft, rb6ControlerLocationTopLanding, rb6COPFinishOther, rb6COPFinishStainlessSteel, rb6DoorNudgingNo,
                rb6DoorNudgingYes, rb6DoorTracksAluminium, rb6DoorTracksOther, rb6DoorTypeCentreOpening, rb6DoorTypeSideOpening,
                rb6EmergencyLoweringSystemNo, rb6EmergencyLoweringSystemYes, rb6ExclusiveServiceNo, rb6ExclusiveServiceYes,
                rb6FacePlateMaterialOther, rb6FacvePlateMaterialStainlessSteel, rb6FalseCeilingNo, rb6FalseCeilingYes, rb6FalseFloorNo, rb6FalseFloorYes,
                rb6FireServiceNo, rb6FireServiceYes, rb6FrontWallOther, rb6FrontWallStainlessSteel, rb6GPOInCarNo, rb6GPOInCarYes, rb6HandrailOther,
                rb6HandrailStainlessSteel, rb6IndependentNo, rb6IndependentServiceYes, rb6LandingDoorFinishOther,
                rb6LandingDoorFinishStainlessSteel, rb6LCDColourBlue, rb6LCDColourRed, rb6LCDColourWhite, rb6LoadWeighingNo, rb6LoadWeighingYes,
                rb6MirrorFullSize, rb6MirrorHalfSize, rb6MirrorOther, rb6OutOfServiceNo, rb6OutOfServiceYes, rb6PositionIndicatorTypeFlushMount,
                rb6PositionIndicatorTypeSurfaceMount, rb6ProtectiveBlanketsNo, rb6ProtectiveBlanketsYes, rb6RearDoorKeySwitchNo,
                rb6RearDoorKeySwitchYes, rb6RearWallOther, rb6RearWallStainlessSteel, rb6SecurityKeySwitchNo, rb6SecurityKeySwitchYes,
                rb6SideWallOther, rb6SideWallStainlessSteel, rb6StructureShaftCOncrete, rb6StructureShaftOther, rb6TrimmerBeamsNo,
                rb6TrimmerBeamsYes, rb6VoiceAnnunciationNo, rb6VoiceAnnunciationYes },

           new Object[] { tb7AuzCOPLocation, tb7CarDepth, tb7CarDoorFinishText, tb7CarHeight, tb7CarLoad, tb7CarWidth, tb7CEilingFinishText,
                tb7ControlerLocationText, tb7COPFinishText, tb7Designations, tb7DoorHeight, tb7DoorTracksText, tb7DoorWidth, tb7FacePlateMaterialText,
                tb7FloorFinish, tb7FrontWallText, tb7HandrailText, tb7HeadRoom, tb7KeyswitchLocation, tb7LandingDoorFinishText, tb7LiftNumbers,
                tb7LiftRating, tb7MainCOPLocation, tb7MirrorText, tb7CarNotes, tb7NumberOfCarEntrances, tb7NumberOfCOPs, tb7NumberOfLandingDoors,
                tb7NumberOfLandings, tb7NumberOfLEDLights, tb7PitDepth, tb7RearWallText, tb7ShaftDepth, tb7ShaftWidth, tb7SideWallText, tb7CarSpeed,
                tb7StructureShaftText, tb7Travel, tb7TypeOfLift, rb7AdvancedOpeningNo, rb7AdvancedOpeningYes, rb7BumpRailNo, rb7BumpRailYes,
                rb7CarDoorFinishOther, rb7CarDoorFinishStainlessSteel, rb7CeilingFinishMirrorStainlessSteel, rb7CeilingFinishOther,
                rb7CeilingFinishStainlessSteel, rb7CEilingFinishWhite, rb7ControlerLocationBottomLAnding, rb7ControlerLocationOther,
                rb7ControlerLocationShaft, rb7ControlerLocationTopLanding, rb7COPFinishOther, rb7COPFinishStainlessSteel, rb7DoorNudgingNo,
                rb7DoorNudgingYes, rb7DoorTracksAluminium, rb7DoorTracksOther, rb7DoorTypeCentreOpening, rb7DoorTypeSideOpening,
                rb7EmergencyLoweringSystemNo, rb7EmergencyLoweringSystemYes, rb7ExclusiveServiceNo, rb7ExclusiveServiceYes,
                rb7FacePlateMaterialOther, rb7FacePlateMaterialStainlessSteel, rb7FalseCeilingNo, rb7FalseCeilingYes, rb7FalseFloorNo, rb7FalseFloorYes,
                rb7FireServiceNo, rb7FireSErviceYes, rb7FrontWallOther, rb7FrontWallStainlessSteel, rb7GPOInCarNo, rb7GPOInCarYes, rb7HandrailOther,
                rb7HandrailStainlessSteel, rb7IndependentServiceNo, rb7IndpendentServiceYes, rb7LandingDoorFinishOther,
                rb7LandingDoorFinishStainlessSteel, rb7LCDColourBlue, rb7LCDColourRed, rb7LCDColourWhite, rb7LoadWeighingNo, rb7LoadWeighingYes,
                rb7MirrorFullSize, rb7MirrorHalfSize, rb7MirrorOther, rb7OutOfSErviceNo, rb7OutOfServiceYes, rb7PositionIndicatorTypeFlushMount,
                rb7PositionIndicatorTypeSurfaceMount, rb7ProtctiveBlanketsNo, rb7ProtectiveBlanketsYes, rb7RearDoorKeySwitchNo,
                rb7RearDoorKeySwitchYes, rb7RearWallOther, rb7RearWallStainlessSteel, rb7SecurityKeySwitchNo, rb7SecurityKeySwitchYes,
                rb7SideWallOther, rb7SideWallStainlessSteel, rb7StructureShaftConcrete, rb7StructureShaftOther, rb7TrimmerBeamsNo,
                rb7TrimmerBeamsYes, rb7VoiceAnunciationNo, rb7VoiceAnnunciationYes },

           new Object[] { tb8AuxCOPLocation, tb8CarDEpth, tb8CarDoorFinishText, tb8CarHeight, tb8Load, tb8CarWidth, tb8CeilingFinishText,
                tb8ControlerLocationText, tb8COPFinishText, tb8Desiginations, tb8DoorHeight, tb8DoorTracksText, tb8DoorWidth, tb8FacePlateMaterialText,
                tb8FloorFinish, tb8FrontWallText, tb8HandrailText, tb8Headroom, tb8KeyswitchLocations, tb8LandingDoorFinishText, tb8LiftNumbers,
                tb8LiftRating, tb8MainCOPLocation, tb8MirrorText, tb8LiftCarNotes, tb8NumberOfCarEntrances, tb8NumberOfCOPs, tb8NumberOfLandingDoors,
                tb8NumberOfLandings, tb8NumberofLEDLights, tb8PitDepth, tb8RearWallText, tb8ShaftDepth, tb8ShaftWidth, tb8SideWallText, tb8Speed,
                tb8StructureShaftText, tb8Travel, tb8TypeOfLift, rb8AdvncedOpeningNo, rb8AdvancedOpeningYes, rb8BumpRailNo, rb8BumpRailYes,
                rb8CarDoorFinishOther, rb8CarDoorFinishStainlessSteel, rb8CeilingFinishMirrorStainlessSTeel, rb8CeilingFinishOther,
                rb8CeilingFinishStainlessSteel, rb8CeilingFinishWhite, rb8ControlerLocationBottomLanding, rb8ControlerLocationOther,
                rb8ControlerLocationShaft, rb8ControlerLocationTopLanding, rb8COPFinishOther, rb8COPFinishStainlessSteel, tb8DoorNudgingNo,
                rb8DoorNudgingYes, rb8DoorTracksAluminium, rb8DoorTracksOther, rb8DoorTypeCentreOpening, rb8DoorTypeSideOpening,
                rb8EmergencyLoweringSystemNo, rb8EmergencyLoweringSystemYes, rb8ExclusiveServiceNo, rb8ExclusiveServiceYes,
                rb8FacePlateMaterialOther, rb8FacePlateMaterialStainlessSteel, rb8FalseCeilingNo, rb8FalseCeilingYes, rb8FalseFloorNo, rb8FalseFloorYes,
                rb8FireSErviceNo, rb8FireServiceYes, rb8FrontWallOther, rb8FrontWallStainlessSteel, rb8GPOInCarNo, rb8GPOInCarYes, rb8HandrailOther,
                rb8HandRailStainlessSteel, rb8IndependentServiceNo, rb8IndependentServiceYes, rb8LandingDoorFinishOther,
                tb8LandingDoorFinishStainlessSteel, rb8LCDColourBlue, rb8LCDColourRed, rb8LCDColourWhite, rb8LoadWeighingNo, rb8LoadWeighingYes,
                rb8MirrorFullSize, rb8MirrorHalfSize, rb8MirrorOther, rb8OutOFServiceNo, rb8OutOfSErviceYes, rb8PositionIndicatorTypeFlushMount,
                rb8PositionIndicatorTypeSurfaceMount, rb8ProtectiveBlanketsNo, rb8ProtectiveBlanketsYes, rb8RearDoorKeySwitchNo,
                rb8RearDoorKeySwitchYes, rb8RearWallOther, rb8RearWallStainlessSteel, rb8SecurityKeySwitchNo, rb8SecurityKeySwitchYes,
                rb8SideWallOther, rb8SideWallStainlessSteel, rb8StructureShaftConcrete, rb8StructureShaftOther, rb8TrimmerBeamsNo,
                rbTrimmerBeamsYes, rb8VoiceAnnunicationNo, rb8VoiceAnnunicationYes },

           new Object[] { tb9AuxCOPLocation, tb9CarDepth, tb9CarDoorFinishText, tb9CarHeight, tb9Load, tb9CarWidth, tb9CeilingFinishText,
                tb9ControlerLocationText, tb9COPFinishText, tb9Designations, tb9DoorHeight, tb9DoorTracksText, tb9DoorWidth, tb9FacePlateMaterialText,
                tb9FloorFinish, tb9FrontWallText, tb9HandrailTexrt, tb9Headroom, tb9KeyswitchLocation, tb9LandingDoorFinishText, tb9LiftNumbers,
                tb9LiftRating, tb9MainCOPLocation, tb9MirrorText, tb9CarNotes, tb9NumberOFCarEntraces, tb9NumberOfCOPs, tb9NumberOfLandingDoors,
                tb9NumberOfLandings, tb9NumberOfLEDLights, tb9PitDepth, tb9RearWallText, tb9ShaftDepth, tb9ShaftWidth, tb9SideWallText, tb9Speed,
                tb9StructureShaftText, tb9Travel, tb9TypeOfLift, rb9AdvancedOpeningNo, rb9AdvancedOpeningYes, rb9BumpRailNo, rb9BumpRailYes,
                rb9CarDoorFinishOther, rb9CarDoorFinishStainlessSteel, rb9CeilingFinishMirrorStainlessSteel, rb9CeilingFinishOther,
                rb9CeilingFinishStainlessSteel, rb9CeilingFinishWhite, rb9ControlerLocationBottomLanding, rb9ControlerLocationOther,
                rb9ControlerLocationShaft, rb9ControlrLocationTopLanding, rb9COPFinishOther, rb9COPFinishStainlessSteel, rb9DoorNudgingNo,
                rb9DoorNudgingYes, rb9DoorTracksAluminium, rb9DoorTracksOther, rb9DoorTypeCentreOpening, rb9DoorTypeSideOpening,
                rb9EmergencyLoweringSystemNo, rb9EmergencyLoweringSystemYes, rb9ExclusiveServiceNo, rb9ExclusiveServiceYes,
                rb9FacePlateMaterialOther, rb9FacePlateMaterialStainlessSteel, rb9FalseCeilingNo, rb9FalseCeilingYes, rb9FalseFloorNo, rb9FalseFloorYes,
                rb9FireServiceNo, rb9FireSErviceYes, rb9FrontWallOther, rb9FrontWallStainlessSteel, rb9GPOInCarNo, rb9GPOInCarYes, rb9HandrailOther,
                rb9HandrailStainlessSteel, rb9IndependentServiceNo, rb9IndependentServiceYes, rb9LandingDoorFinishOther,
                rb9LandingDoorFinishStainlessSteel, rb9LCDColourBlue, rb9LCDColourRed, rb9LCDColourWhite, rb9LoadWeighingNo, rb9LoadWeighingYes,
                rb9MirrorFullSize, rb9MirrorHalfSize, rb9MirrorOther, rb9OutOfServiceNo, rb9OutOfServiceYes, rb9PositionIndicatorTypeFlushMount,
                rb9PositionIndicatorTypeSurfaceMount, rb9ProtectiveBlanketsNo, rb9ProtectiveBlanketsYes, rb9RearDoorKeySwitchNo,
                rb9RearDoorKeySwitchYes, rb9RearWallOther, rb9RearWallStainlessSteel, rb9SecurityKeySwitchNo, rb9SecurityKeySwitchYes,
                rb9SideWallOther, rb9SideWallStainlessSteel, rb9StructureShaftConcrete, rb9StructureShaftOther, rb9TrimmerBeamsNo,
                rb9TrimmerBeamsYes, rb9VoiceAnnunciationNo, rb9VoiceAnnunciationYes },

           new Object[] { tb10AuxCOPLocation, tb10CarDepth, tb10CarDoorFinishText, tb10CarHeight, tb10LiftCarLoad, tb10CarWidth, tb10CEilingFinishText,
                tb10ControlerLocationText, tb10COPFinishText, tb10Desigination, tb10DoorHeight, tb10DoorTracksText, tb10DoorWidth, tb10FacePlateMaterialText,
                tb10FloorFinish, tb10FrontWallText, tb10HandrailText, tb10Headroom, tb10KeyswitchLocation, tb10LandingDoorFinishText, tb10LiftNumbers,
                tb10LiftRating, tb10MainCOPLocation, tb10MirrorText, tb10LiftCarNotes, tb10NumberofCarEntrances, tb10NumberOfCOPs, tb10NumberofLandingDoors,
                tb10NumberofLandings, tb10NumberOfLEDLIghts, tb10PitDepth, tb10RearWallText, tb10ShaftDepth, tb10ShaftWidth, tb10SideWallText, tb10Speed,
                tb10StructureShaftText, tb10Travel, tb10TypeOfLift, rb10AdvancedOpeningNo, rb10AdvancedOpeningYes, rb10BumpRailNo, rb10BumpRaidYes,
                rb10CarDoorFinishOther, rb10CarDoorFinishStainlessSteel, rb10CEilingFinishMirrorStainlessSteel, rb10CEilingFinishOther,
                rb10CeilingFinishStainlessSteel, rb10CeilingFinishWhite, rb10ControlerLocationBottomLanding, rb10ControlerLocationOther,
                rb10ControlerLocationShaft, rb10ControlerLocationTopLanding, rb10COPFinishOther, rb10COPFinishStainlessSteel, rb10DoorNudgingNo,
                rb10DoorNudgingYes, rb10DoorTracksAluminium, rb10DoorTracksOther, rb10DoorTypeCentreOpening, rb10DoorTypeSideOpening,
                rb10EmergencyLoweringSystemNo, rb10EmergencyLoweringSystemYes, rb10ExclusiveSErviceNo, rb10ExclusiveServiceYes,
                rb10FacePlateMaterialOther, rb10FacePlateMaterialStainlessSteel, rb10FalseCeilingNo, rb10FalseCeilingYes, rb10FalseFloorNo, rb10FalseFloorYes,
                rb10FireSERviceNo, rb10FireSErviceYes, rb10FrontWallOther, rb10FrontWallStainlessSteel, rb10GPOInCarNo, rb10GPOInCarYes, rb10HandrailOther,
                rb10HandrailStainlessSteel, rb10IndependentServiceNo, rb10IndependentServiceYes, rb10LAndingDoorFinishOtherr,
                rb10LandingDoorFinishStainlessSteel, rb10LCDColourBlue, rb10LCDColourRed, rb10LCDColourWhite, rb10LoadWEighingNo, rb10LoadWeighingYes,
                rb10MirrorFullSize, rb10MirrorHalfSize, rb10MirrorOther, rb10OutOfServiceNo, rb10OutOFServiceYes, rb10PositionIndicatorTypeFlushMount,
                rb10PositionIndicatorTypeSurfaceMount, rb10ProtectiveBlanketNo, rb10ProtectiveBlanketYes, rb10RearDoorKeySwitchNo,
                rb10RearDoorKeySwitchYes, rb10RearWallOther, rb10RearWallStainlessSteel, rb10SecurityKeySwitchNo, rb10SecurityKeySwitchYes,
                rb10SideWallOther, rb10SideWallStainlesSteel, rb10StructureShaftConcrete, rb10StructureShaftOther, rb10TrimmerBeamsNo,
                rb10TimmerbeamsYes, rb10VoiceAnnunciationNo, rb10VoiceAnnunciationYes },

           new Object[] { tb11AuxCOPLocation, tb11CarDepth, tb11CarDoorFinishText, tb11CarHeight, tb11LiftCarLoad, tb11CarWidth, tb11CeilingFinishText,
                tb11ControlerLocationText, tb11COPFinishText, tb11Designations, tb11DoorHeight, tb11DoorTracksText, tb11DoorWidth, tb11FaceplateMaterialText,
                tb11FloorFinish, tb11FrontWallText, tb11HandrailText, tb11Headroom, tb11KeyswitchLocation, tb11LandingDoorFInishOther, rb11LiftNumbers,
                tb11LiftRating, tb11MainCOPLocation, tb11MirrorText, tb11LiftCarNote, tb11NumberofCarEntrances, tb11NumberOfCOPs, tb11NumberOfLandingDoors,
                tb11NumberOfLandings, tb11NumberOfLEDLights, tb11PitDepth, tb11RearWallText, tb11ShaftDepth, tb11ShaftWidth, tb11SideWallText, tb11Speed,
                rb11StructureShaftText, tb11Travel, tb11TypeOfLift, rb11AdvancedOpeningNo, rb11AdvancedOpeningYes, rb11BumpRailNo, rb11BumpRailYes,
                rb11CarDoorFinishOther, rb11CarDoorFinishStainlessSteel, rb11CeilingFinishMirrorStainlessSteel, rb11CeilingFinishOther,
                rb11CeilingFinishStainlessSteel, rb11CeilingFinishWhite, rb11ControlerLocationBottomLanding, rb11ControlerLocationOther,
                rb11ControlerLocationShaft, tb11ControlerLocationTopLanding, rb11COPFinishOther, rb11COPFinishStainlessSteel, rb11DoorNudgingNo,
                rb11DoorNudgingYes, rb11DoorTracksAluminium, rb11DoorTracksOther, rb11DoorTypeCentreOpening, rb11DoorTypeSideOpening,
                rb11EmergencyLoweringSystemNo, rb11EmergencyLoweringSystemYes, rb11ExclusiveServiceNo, rb11ExclusiveServiceYes,
                rb11FacePlateMaterialOther, rb11FacePlateMaterialStainlessSTeel, rb11FalseCeilingNo, rb11FalseCEilingYes, rb11FalseFloorNo, rb11FalseFloorYes,
                rb11FireServiceNo, rb11FireServiceYes, rb11FrontWallOther, rb11FrontWallStainlessSteel, rb11GPOInCarNo, rb11GPOInCarYes, rb11HandrailOther,
                rb11HandrailStainlessSteel, rb11IndependentSErviceNO, rb11IndependentServiceYes, rb11LandingDoorFinishOther,
                rb11LandingDoorFinishStainlessSteel, rb11LCDColourBlue, rb11LCDColourRed, rb11LCDColourWhite, rb11LoadWeighingNo, rb11LoadWeighingYes,
                rb11MirrorFullSize, rb11MirrorHalfSize, rb11MirrorOther, rb11OutOfServiceNo, rb11OutOfSErviceYes, rb11PositionIndicatorTypeFlushMount,
                rb11PositionIndicatorTypeSurfaceMount, rb11ProtectiveBlanketNo, rb11ProtectiveBlanketsYes, rb11RearDoorKeySwitchNo,
                rb11RearDoorKeySwitchYes, rb11RearWallOther, rb11RearWallStainlessSteel, rb11SecurityKeySwitchNo, rb11SecurityKeySwitchYes,
                rb11SideWallOther, rb11SideWallStainlessSteel, rb11StructureShaftConcrete, rb11StrructureShaftOther, rb11TrimmerBeamNo,
                rb11TrimmerBeamsYes, rb11VoiceAnnunciationNo, rb11VoiceAnnunciationYes },

           new Object[] { tb12AuxCOPLocation, tb12CarDepth, tb12CarDoorFinishText, tb12CarHeight, tb12CarLoad, tb12CarWidth, tb12CeilingFinishText,
                tb12ControlerLocationText, tb12COPFinishText, tb12Designations, tb12LandingDoorHeight, tb12DoorTracksText, tb12LandingDoorWidth, tb12FacePlateMaterialText,
                tb12FloorFinish, tb12FrontWallText, tb12HandrailText, tb12Headroom, tb12KeyswitchLocation, tb12LandingDoorFinishText, tb12LiftNumbers,
                tb12CarLiftRating, tb12MainCOPLocation, tb12MirrorText, tb12LiftCarNotes, tb12CarNumberOfCarEntrances, tb12NumberOfCOPs, tb12NumberOfLandingDoors,
                tb12NumberOfLandings, tb12NumberOfLEDLights, tb12PitDepth, tb12RearWallText, tb12ShaftDepth, tb12ShaftWidth, tb12SideWallText, tb12CarSpeed,
                tb12StructureShaftText, tb12Travel, tb12TypeOfLift, rb12AdvancedOpeningNo, rb12AdvancedOpeningYes, rb12BumpRailNo, rb12BumpRailYes,
                rb12CarDoorFinishOther, rb12CarDoorFinishStainlessSteel, rb12CeilingFinishMirrorStainlessSteel, rb12CeilingFinishOther,
                rb12CeilingFinishStainlessSteel, rb12CeilingFinishWhite, rb12ControlerLocationBottomLanding, rb12ControlerLocationOther,
                rb12ControlerLocationShaft, rb12ControlerLocationTopLanding, rb12COPFinishOther, rb12COPFinishStainlessSteel, rb12LandingDoorNudgingNo,
                rb12LandingDoorNudgingYes, rb12DoorTracksAluminium, rb12DoorTracksOther, rb12DoorTypeCentreOpening, rb12DoorTypeSideOpening,
                rb12EmergencyLoweringSystemNo, rb12EmergencyLoweringSystemYes, rb12ExclusiveServiceNo, rb12ExclusiveServiceYes,
                rb12FacePlateMaterialOther, rb12FacePlateMaterialStainlessSteel, rb12FalseCeilingNo, rb12FalseCeilingYes, rb12FalseFloorNo, rb12FalseFloorYes,
                rb12FireServiceNo, rb12FireServiceYes, rb12FrontWallOther, rb12FrontWallStainlessSteel, rb12GPOInCarNo, rb12GPOInCarYes, rb12HandrailOther,
                rb12HandrailStainlessSTeel, rb12IndependentServiceNo, rb12IndependentServiceYes, rb12LandingDoorFinishOther,
                rb12LandingDoorFinishStainlessSteel, rb12LCDColourBlue, rb12LCDColourRed, rb12LCDColourWhite, rb12LoadWeighingNo, rb12LoadWeighingYes,
                rb12MirrorFullSize, rb12MirrorHalfSize, rb12MirrorOTher, rb12OutOfServiceNo, rb12OutOfServiceYes, rb12PositionIndicatorTypeFlushMount,
                rb12PositionIndicatorTypeSurfaceMount, rb12ProtectiveBlanketsNo, rb12ProectiveBlanketsYes, rb12RearDoorKeySwitchNo,
                rb12RearDoorKeySwitchYes, rb12RearWallOther, rb12RearWallStainlessSteel, rb12SecurityKeySwitchNo, rb12SecurityKeySwitchYes,
                rb12SideWallOther, rb12SideWallStainlessSteel, rb12StructureShaftConcrete, rb12StructureShaftOther, rb12TrimmerBeamsNo,
                rb12TrimmerBeamsYes, rb12VoicAnnunciationNo, rb12VoiceAnnuniationYes }
        };

            #endregion

            for (int i = 0; i < pageObjects[pageToBeFilled].Length; i++)
            {
                if (pageObjects[pageToBeFilled][i] is TextBox)
                {
                    TextBox sourceTextBox = (TextBox)pageObjects[0][i];
                    TextBox recipientTextBox = (TextBox)pageObjects[pageToBeFilled][i];

                    recipientTextBox.Text = sourceTextBox.Text;
                }
                else if (pageObjects[pageToBeFilled][i] is RadioButton)
                {
                    RadioButton sourceRadioButton = (RadioButton)pageObjects[0][i];
                    RadioButton recipientRadioButton = (RadioButton)pageObjects[pageToBeFilled][i];

                    recipientRadioButton.Checked = sourceRadioButton.Checked;
                }
                else if (pageObjects[pageToBeFilled][i] is CheckBox)
                {
                    CheckBox sourceCheckBox = (CheckBox)pageObjects[0][i];
                    CheckBox recipientCheckBox = (CheckBox)pageObjects[pageToBeFilled][i];

                    recipientCheckBox.Checked = sourceCheckBox.Checked;
                }
            }
        }
        #endregion

    }
}
