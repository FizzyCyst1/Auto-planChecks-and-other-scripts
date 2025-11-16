/*******************************************************************************************************
* Namespace:       BeginPlanCheck
* Application:     BeginPlanCheck
* Purpose:         It seems I have committed the cardinal sin of automating an inherently inefficient process. This is an intermediate script towards more automated plan-checking. 
*                  This simple script will identify the plan-check type required, e.g.:
*                   OTB1 or DTB1 Truebeam VMAT/IMRT/Conformal/SABR/SRS, Orthovoltage, Halcyon VMAT/IMRT and move the relevant plan-check template file (excel) used by physics.
*                   It will pre-fill the patient ID, plan-name and date within the excel physics plan-checklist and save a copy in the correct folder location on Ndrive.
*                   This script will also (eventually ... in version 2.1!) send the patient plan to Mobius (unless an Orthovoltage plan). 
*                   It is intended to be run once by a physicist for every patient at the commencement of a physics plan-check
* Author:          Joshua Hiatt 
* Date:            02/07/2024
* Comments:        
* 
* 
* 
* *****************************************************************************************************
* Checked BY:      
* Date:            
* Comments:        
* 
* 
*  *****************************************************************************************************/

//******************************************************************************************************
// Minor edits:
//*****************************************************************************************************

//07/07/2025: Updated comment, to include plan-check version.

//** The treatment types with individual checklists are as follows:
//OTB1:
// 1. Conformal
// 2. Electron
// 3. Field in Field
// 4. Hybrid
// 5. HyperArc
// 6. IMRT
// 7. VMAT
// 8. Liver SBRT
// 9. Non-Liver SABR



//Ortho:
// 10. Ortho

//Halcyon:
// 11. IMRT
// 12. VMAT

//(repeat for DTB1? unless different plan-checklists).
// 13. Conformal
// 14. Electron
// 15. Field in Field
// 16. Hybrid
// 17. HyperArc
// 18. IMRT
// 19. VMAT
// 20. Liver SBRT
// 21. Non-Liver SABR

// 22. Halcyon Hybrid.

//23. Dubbo Ortho.

//**///



using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Numerics;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Web;
using System.Windows.Interop;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Text.RegularExpressions;
//using System.Windows;
//using Window = System.Windows.Window;




// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.1.1.1")]
[assembly: AssemblyFileVersion("1.1.1.1")]
[assembly: AssemblyInformationalVersion("2.12")]

// TODO: Uncomment the following line if the script requires write access.
// [assembly: ESAPIScript(IsWriteable = true)]

namespace BeginPlanCheck2
{
    class Program
    {
        /***Defining local variables: ****/
        #region Local Variables
        private static string _patientId; //private (to this class, so only visible to this class/?method).
        private static string _courseId;
        private static string _planId;
        private static string _structureSetId;
        private static string _currentUserId;
        private static string _usern; //current username;


        private static string tmtmach = string.Empty;
        private static int planchecktype = 0;
        private static bool isElectron = false;
        private static bool isOHAL = false;
        private static bool isOTB1 = false;
        private static bool isDTB1 = false;
        private static bool isVMAT = false;
        private static bool isHyperArc = false;
        private static bool isSABR = false;
        private static bool isFIF = false;
        private static bool isIMRT = false;
        private static bool ishybrd = false;
        private static bool isconfrmal = false;
        private static bool isLiverSABR = false;
        private static bool isDubboOrtho = false;
        private static bool beamalign = false;
        private static bool JT_isOn_allBeams = true; //Start as true, as some tests such as halcyon this isn't applicable for and isn't changed.
        //private static bool isSABR_check2 = false; //if other sabr tests don't pick up sabr status.
        private static bool allsametmtmach = false;



        private static string destinationFolder = string.Empty; //@"\\Ordcari-mfs901\va_data$\ProgramData\Vision\PublishedScripts\In Development\JH in development uSe aT oWn RiSk\testfolders\OTB1"; //e.g. patient QA folder.
        private static string templateLocation = string.Empty; //@"\\Ordcari-mfs901\va_data$\ProgramData\Vision\PublishedScripts\In Development\JH in development uSe aT oWn RiSk\testfolders\OTB1\VMAT\MACRO_TESTING2_ACUROS_16100-Eclipse VMAT Physics Checks V1.8.xlsm"; // e.g. location of relevant patient QA list.
        private static string filename = string.Empty;

        private static bool bigiftrue = false;

        private static string ticksymbol = "\u2713\u2071"; //\xB3"; //should be characters for tick symbol in excel; .
        
        private static string crossymbol = "X\u2071";//\xB3; //

        //Dubbo check-box specific, in array form:
        private static string[] Dubbo_row_strings = new string[7];
        private static int[] Dubbo_column_ints = new int[7];
        private static string[] Dubbo_TRUE_false_checkbox = new string[7];
        private static string[] Dubbo_box_captions = new string[7];
        //
        private static string Dubbo_Rxapprvd = "false";
        private static string Dubbo_Planapprvd = "false";
        private static string Dubbo_dosecalalgo = "false";
        private static string Dubbo_pho_opt = "false";
        private static string Dubbo_DG = "false";
        private static string Dubbo_He = "false";
        private static string Dubbo_BAlign = "false";
        //

        //Set default values, only changed within switch cases if they differ from default.
        private static double desired = 0.125; //Desired dose-grid in cm. Set to default of 0.125 for SABR/SRS (although protocol is 0.125); and set to 0.2 cm for VMAT/IMRT std..
        private static string desiredphotoncalcalgo = "Acuros_16100";
        private static string pho_opt = string.Empty;
        private static string reviewer = string.Empty; //Plan approver or reviewer; (assinged to non-empty only if plan is reviewed/approved by an RO).
        


        /*** Set ESAPI variables ****/
        private static Patient patient;
        private static Course courSelec;
        private static PlanSetup planSelec;
        private static ReferencePoint refSelec;

        #endregion

        /*** Ensure Single Threaded ****/
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                /*** Check for arguments supplied ****/
                if (args.Any())
                {
                    if (args.First().Split(';').Count() == 6)
                    {
                        _patientId = args.First().Split(';').First();
                        _courseId = args.First().Split(';').ElementAt(1);
                        _planId = args.First().Split(';').ElementAt(2);
                        _structureSetId = args.First().Split(';').ElementAt(3);
                        _currentUserId = args.First().Split(';').ElementAt(4);
                        // _usern = args.First().Split(';').Last();
                        _usern = args.First().Split(';').ElementAt(5).ToString()+' '+args.Last().ToString();//.Split(';').ElementAt(5);//+" "+args.Last().Split(';').ElementAt(5); //Should give firstName LastName; if they are seperate.
                    }
                }



                using (VMS.TPS.Common.Model.API.Application app = VMS.TPS.Common.Model.API.Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
                Console.ReadLine();
            }
        }
        static void Execute(VMS.TPS.Common.Model.API.Application app)
        {
            Patient patient = app.OpenPatientById(_patientId);//open using supplied.
            courSelec = patient.Courses.FirstOrDefault(x => x.Id.Equals(_courseId));
            planSelec = courSelec.PlanSetups.FirstOrDefault(x => x.Id == _planId);


            //Default row,column locations for excel plan-check documents. The switch case will define new values if they differ from standard.
            string colV = "C";
            int rowV_patId = 8;
            int rowV_planid = 10;
            int rowV_userid = 11;
            int rowV_date1 = 41;
            string colV_date1 = "D";
            int rowV_dosecalalgo = 19;
            int rowV_DG = 21;
            int row_comment = 27;
            string colV_comment = "B";
            int rowPlanapprvd = 14;
            int rowHe = 21; //Hetereogenety is ON test.
            int row_BAlign = 24; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
            int rowV_pho_opt = 20; //PO_16100;
            int rowRxapprvd = 12;//RX approver
            int rowAF = 24; //Autofeathering (OHAL only)
            int row_mlc = 17; //MLC type is DoseDynamic | VMAT (OHAL only).
            int row_samemach = 25; //Halcyon only.
            //




            Console.WriteLine("Starting begin plan-check.");







          

            //** Identify treatment technique - pull the last beam**//

            Beam beam = planSelec.Beams.Last(x => !x.IsSetupField);

            //*** Identify treating machine **//
            tmtmach = beam.TreatmentUnit.Id;
            isOHAL = tmtmach.Contains("OHAL");
            isDTB1 = tmtmach.Contains("DTB1");
            isOTB1 = tmtmach.Contains("OTB1");

            //Determine if a hybrid plan, by defining an VMAT beam and IMRT beam that aren't setup fields. Now if both of these are not empty then it is a hybrid.//


            Beam beamarry1 = null;
            Beam beamarry2 = null;

            try
            {
                beamarry1 = planSelec.Beams.First(x => !x.IsSetupField && x.MLCPlanType.ToString().Contains("VMAT")); //
            }
            catch

            { beamarry1 = null; }

            try
            {
                beamarry2 = planSelec.Beams.First(x => !x.IsSetupField && x.MLCPlanType.ToString().Contains("DoseDynamic")); //!x.IsSetupField & 
            }
            catch
            { beamarry2 = null; }


            if ((beamarry1 != null) && (beamarry2 != null))
            {
                ishybrd = true;
                planchecktype = isDTB1 ? 16 : isOHAL ? 22 : 4; //DTB1 Hybrid, if not, then is it Halcyon HB?, if not then must be OTB1 hybrid.

            }

            Console.WriteLine(beam.Technique.ToString());
            Console.WriteLine(planSelec.Name.ToString().ToLower());
            //** if not a hybrid plan, and there are non-setup fields in the plan, determine the plan type: **//

            if (beam != null) //&& (ishybrd != true)
            {
                isElectron = beam.EnergyModeDisplayName.Contains("E");
                tmtmach = beam.TreatmentUnit.Id;
                isOHAL = tmtmach.Contains("OHAL");
                isDTB1 = tmtmach.Contains("DTB1");
                isOTB1 = tmtmach.Contains("OTB1");
                isVMAT = beam.MLCPlanType.ToString().Contains("VMAT");
                isIMRT = beam.MLCPlanType.ToString().Contains("DoseDynamic");
                //isPal = beam.Plan.PlanIntent.ToString().ToLower().Contains("pallative");
                Console.WriteLine(beam.MLCPlanType.ToString()); //This is the 'Treatment Type' parameter viewable in the Rx.
                isFIF = beam.MLCPlanType.ToString().Contains("DoseDynamic") & beam.ControlPoints.Count < 12; //Crude measure of Field in field, must be dose dynamic and have very few control points.
                isconfrmal = beam.MLCPlanType.ToString().Contains("Static");
                try
                {
                    string isSABRR = beam.Plan.RTPrescription.Technique.ToString();// delete once trouble shooting finished?
                    Console.WriteLine($"The prescription technique is: {isSABRR}");
                    isSABR = beam.Plan.RTPrescription.Technique.ToString().ToLower().Contains("stereotactic") ; //(beam.Technique.ToString().Contains("SRS Arc Therapy-I") | beam.Technique.ToString().Contains("SRS ARC"));// & ~isPal;
                    isLiverSABR = (planSelec.Id.ToString().ToLower().Contains("liver") & isSABR) || (beam.Plan.RTPrescription.Site.ToString().ToLower().Contains("liver") & isSABR); //** Needs to distinguish liver sabr and non-liver sabr. Crude test looking at planId. Note

                    bool SRT_criteria = (beam.Plan.RTPrescription.NumberOfFractions > 1 & beam.Plan.DosePerFraction.Dose >= 6 & !isIMRT);

                    if (isSABR == false) //try test again, as some slipping through the cracks. But have sabr as the name.
                    {
                        isSABR = (planSelec.Id.ToString().ToLower().Contains("sabr") & !isLiverSABR) | planSelec.Id.ToString().Contains("SRT") | SRT_criteria; //Is not liversabr,
                                    //and the plan ID contains sabr; but the Rx and beam type are not stereo. OR planID contains in all upper case "SRT". OR prescription meets SRT criteria.                                                                                                                                                               //beams may not always be SRS Arc Therapy, can also just be standard selection.. in which case this test will not identify. **//
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Perhaps no perscription has been set?");
                    Console.WriteLine(ex.ToString());
                    string isSABRR = "Nope.";
                    isSABR = false;
                }


                
                isDubboOrtho = tmtmach.Contains("DUBBO ORTHO");

                isHyperArc = beam.Technique.ToString().ToLower().Contains("srs hyper") | beam.Technique.ToString().ToLower().Contains("srs Hyper Arc") | beam.ToleranceTableLabel.ToString().ToLower().Contains("hyperarc");

                Console.WriteLine(beam.Technique.ToString());

                // NilMLC = beam.MLCPlanType.ToString().Contains("NotDefined"); //If MLC is undefined, then use Jaws to pull field-size.

                if (isOTB1 && (ishybrd != true))
                {
                    planchecktype = isElectron ? 2 : isFIF ? 3 : isIMRT ? 6 : isconfrmal ? 1 : isHyperArc ? 5 : isLiverSABR ? 8 : isSABR ? 9 : isVMAT ? 7 : 10;  // ternary condition if its an Electron - = 2; if no than next test.
                                                                                                                                                                 //if it's not electron, than is it FIF? = 3; or 
                                                                                                                                                                 //If it's not FIF, is it IMRT? = 6;
                                                                                                                                                                 //If no, than is it conformal? = 1;
                                                                                                                                                                 //if nah, then,... uhh is it HyperArc? = 5;
                                                                                                                                                                 //if nope then is it Liver SABR? 8,
                                                                                                                                                                 //if nope, then is it Sabr? 9,
                                                                                                                                                                 //if no, then is it VMAT? 7,
                                                                                                                                                                 // else it must be .... Ortho? Case 10.
                                                                                                                                                                 // But this Ortho condition can actually never be reached unless the machine type is also OTB1 - which it won't be.... but ortho plan-checks are rare.
                }                                                                                                                                                   //so would more likely be an error case. Have resolved this by having alternative test for ortho units.
                else if (isDTB1 && (ishybrd != true))
                {
                    planchecktype = isElectron ? 14 : isFIF ? 15 : isIMRT ? 18 : isconfrmal ? 13 : isHyperArc ? 17 : isLiverSABR ? 20 : isSABR ? 21 : isVMAT ? 19 : 23;
                }
                else if (isOHAL && (ishybrd != true))
                {
                    planchecktype = isIMRT ? 11 : 12; //Either IMRT or VMAT (SABR not yet introduced).
                }
                else if (ishybrd != true)
                {
                    //Must be ortho
                    planchecktype = isDubboOrtho ? 23 : 10; //Have identifier between the two. 23 is Dubbo ortho, 10 is orange ortho.
                }

                

                Console.WriteLine($"The plan check case type is: {planchecktype}");
                //Switch case to assign plan-check type (identify numbers 1-23...)
                //No need to duplicate switch again later, could just identify export location in one fell swoop.
                switch (planchecktype)
                {
                    case 1:
                        Console.WriteLine("Conformal");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\1. Conformal Patient QA\! Template Do Not Delete\ACUROS 16100 - Eclipse MULTIPLAN RTPS  Physics Checks V1.6.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\1. Conformal Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_Conformal_PhysicsChecks.xlsx"; //Name of copied template
                        desired = 0.2; // desired DG size in cm.

                        colV = "D";
                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 10;
                        rowV_date1 = 30;
                        colV_date1 = "D"; //No date entered for conformal.
                        rowV_dosecalalgo = 18;
                        rowV_DG = 19;
                        row_comment = 26;
                        colV_comment = "B";
                        rowPlanapprvd = 14;
                        rowHe = 20; //Hetereogenety is ON test.
                        row_BAlign = 23; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        rowV_pho_opt = 33; //Not in conformal checklist. //PO_16100;
                        rowRxapprvd = 11;//RX approver




                        break; //break to stop the switch case from further executing.
                    case 2:
                        Console.WriteLine("Electron");
                        templateLocation = @"\\"; //Where the template is sourced from.
                        destinationFolder = @"\\"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_Electron_PhysicsChecks.xlsx";
                        desired = 0.2; // desired DG size in cm.
                        break;
                    case 3:
                        Console.WriteLine("Field in Field");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\2. Field in Field Patient QA\! Template Do Not Delete\ACUROS  16100 -FinF Eclipse RTPS Physics Checks v6.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\2. Field in Field Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_FiF_PhysicsChecks.xlsx";
                        desired = 0.2; // desired DG size in cm.
                        colV = "D";
                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 11;
                        rowV_date1 = 33; //No date.
                        colV_date1 = "D";
                        rowV_dosecalalgo = 19;
                        rowV_DG = 20;
                        row_comment = 28;
                        colV_comment = "B";
                        rowPlanapprvd = 15;
                        rowHe = 21; //Hetereogenety is ON test.
                        row_BAlign = 25; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        rowV_pho_opt = 36; //No Photon Optimiser in FiF //PO_16100;
                        rowRxapprvd = 12;//RX approver

                        break;
                    case 4:
                        Console.WriteLine("Hybrid");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\10. IMRT and VMAT Hybrid Plan Patient QA\! Template DO NOT DELETE\ACUROS 16100-Eclipse IMRT+VMAT Hybrid Plan Physics Checks V.6.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\10. IMRT and VMAT Hybrid Plan Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_Hybrid_PhysicsChecks.xlsx";
                        desired = 0.2; // desired DG size in cm.
                        row_BAlign = 26;
                        rowV_date1 = 43;
                        rowV_dosecalalgo = 20;
                        colV_date1 = "D";
                        rowPlanapprvd = 16;
                        rowHe = 23;
                        rowV_DG = 22;
                        row_comment = 28;
                        colV_comment = "B";

                        rowRxapprvd = 12;//RX approver
                        rowPlanapprvd = 16;
                        rowV_pho_opt = 21;
                        row_BAlign = 26; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        break;
                    case 5:
                        Console.WriteLine("HyperArc");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\PATIENT ID_ HyperArc ACUROS_16100-Eclipse Physics Checks V1.0.xlsm"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\12. HyperArc QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_HyperArc_PhysicsChecks.xlsm";

                        desired = 0.1;
                        colV = "C";
                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 11;
                        rowV_date1 = 51;
                        colV_date1 = "D";
                        rowV_dosecalalgo = 28;
                        rowV_DG = 30;
                        row_comment = 35;
                        colV_comment = "B";
                        rowPlanapprvd = 22;
                        rowHe = 32; //Hetereogenety is ON test.
                        row_BAlign = 23; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        rowV_pho_opt = 29; //PO_16100;
                        rowRxapprvd = 21;//RX approver

                        //


                        break;
                    case 6:
                        Console.WriteLine("IMRT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\ACUROS 16100-Eclipse IMRT Physics Checks V.10.xlsm"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\3. IMRT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_IMRT_PhysicsChecks.xlsm";
                        
                        rowV_date1 = 42;
                        colV_date1 = "D";
                        //colV_comment = "B";
                        desired = 0.2; // desired DG size in cm.
                        rowPlanapprvd = 15;
                        rowHe = 22;
                        row_BAlign = 25;

                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 11;
                        rowV_dosecalalgo = 19;
                        rowV_DG = 21;
                        row_comment = 27;
                        colV_comment = "B";
                        rowV_pho_opt = 20; //PO_16100;
                        rowRxapprvd = 12;//RX approver


                        break;
                    case 7:
                        Console.WriteLine("VMAT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\ACUROS_16100-Eclipse VMAT Physics Checks V1.8.xlsm"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\5. VMAT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        Console.WriteLine("OTB1");
                        // Console.ReadLine();
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_VMAT_PhysicsChecks.xlsm";
                        desired = 0.2; // desired DG size in cm.
                        rowV_dosecalalgo = 18;
                        rowV_DG = 20;
                        row_comment = 26;
                        colV_comment = "B";
                        rowV_pho_opt = 19; //PO_16100;
                        break;
                    case 8:
                        Console.WriteLine("Liver SBRT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\Template_ Liver SABR ACUROS_16100-Eclipse SABR Physics Checks V1.0.xlsm"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\13. Liver SBRT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_Liver_SABR_PhysicsChecks.xlsm";
                        desired = 0.125; // desired DG size in cm.
                        rowV_dosecalalgo = 29;
                        rowV_DG = 31;
                        row_comment = 35;
                        colV_comment = "B";
                        rowPlanapprvd = 20;
                        rowHe = 32;
                        rowV_date1 = 12;
                        colV_date1 = "D";
                        //colV_date2 = "C";
                        row_BAlign = 21; //Needs to also check that it's contained WITHIN PTV Structure.
                        rowV_pho_opt = 30;
                        rowRxapprvd = 19; //RX approver by RO.
                        break;
                    case 9:
                        Console.WriteLine("Non-Liver SABR");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\Template_ SABR ACUROS_16100-Eclipse SABR Physics Checks V2.0.xlsm"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\11. Non-Liver SABR Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_SABR_PhysicsChecks.xlsm";
                        rowPlanapprvd = 19;
                        desired = 0.125; // desired DG size in cm.
                        rowV_dosecalalgo = 25;
                        rowHe = 28;
                        row_comment = 31;
                        rowV_date1 = 47;
                        colV_comment = "B";
                        colV_date1 = "D";
                        rowV_DG = 27;
                        row_BAlign = 20; //Needs to also check that it's contained WITHIN PTV Structure.
                        rowV_pho_opt = 26;
                        rowRxapprvd = 18; //RX approver by RO.
                        break;
                    case 10:
                        Console.WriteLine("Ortho");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\4. Orthovoltage Patient QA\! Template Do Not Delete\CWCCC Measurement Based Orthovoltage MU Check v.1.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\4. Orthovoltage Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_Ortho_PhysicsMU.xlsx";
                        colV = "B";
                        rowV_patId = 4;
                        rowV_planid = 5;
                        rowV_userid = 32;
                        rowV_date1 = 31;
                        colV_date1 = "B";
                        row_comment = 35;
                        colV_comment = "A";
                        rowRxapprvd = 6;//RX approver
                                        //

                        break;
                    case 11:
                        Console.WriteLine("Halcyon IMRT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Halcyon\ACUROS 16100-Eclipse IMRT Physics Checks_Halcyon_V.10.xlsm"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\15. Halcyon Patient QA\!Halcyon_IMRT patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_HalcyonIMRT_PhysicsChecks.xlsm";
                        desired = 0.2; // desired DG size in cm.

                        rowHe = 24;
                        row_BAlign = 29;
                        rowAF = 25; //Autofeathering.
                        row_comment = 31;

                        rowV_date1 = 46;
                        colV_date1 = "D";
                        rowV_dosecalalgo = 20;
                        rowPlanapprvd = 15;

                        rowV_DG = 23;

                        colV_comment = "B";

                        rowHe = 24; //Hetereogenety is ON test.
                        row_BAlign = 29; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        rowV_pho_opt = 21; //PO_16100;
                        rowRxapprvd = 12;//RX approver

                        row_mlc = 17; //MLC type is DoseDynamic | VMAT (OHAL only).
                        row_samemach = 26; //Halcyon only.
                        break;
                    case 12:
                        Console.WriteLine("Halcyon VMAT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Halcyon\ACUROS_16100-Eclipse VMAT Physics Checks_ Halcyon V1.8.xlsm"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\15. Halcyon Patient QA\!Halcyon_VMAT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_HalcyonVMAT_PhysicsChecks.xlsm";
                        desired = 0.2; // desired DG size in cm.
                        row_comment = 30;
                        rowV_DG = 22;
                        rowHe = 23;
                        row_BAlign = 28;
                        rowV_pho_opt = 20;
                        row_mlc = 16;
                        row_samemach = 25; //Halcyon only.
                        break;
                    case 13:
                        Console.WriteLine("DTB Conformal");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\1. Conformal Patient QA\! Template Do Not Delete\ACUROS 16100 - Eclipse MULTIPLAN RTPS  Physics Checks V1.6.xlsx"; //Where the template is sourced from.//templateLocation = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\1. Conformal Patient QA\! Template Do Not Delete\WCCD ACUROS  16100 - Eclipse MULTIPLAN RTPS  Physics Checks Conformal Plans V1.4.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\1. Conformal Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_WCCD_ACUROS_16100-Eclipse_Conformal_PhysicsChecks.xlsx";
                        desired = 0.2; // desired DG size in cm.

                        colV = "D";
                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 10;
                        rowV_date1 = 30;
                        colV_date1 = "D"; //No date entered for conformal.
                        rowV_dosecalalgo = 18;
                        rowV_DG = 19;
                        row_comment = 26;
                        colV_comment = "B";
                        rowPlanapprvd = 14;
                        rowHe = 20; //Hetereogenety is ON test.
                        row_BAlign = 23; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        rowV_pho_opt = 33; //Not in conformal checklist. //PO_16100;
                        rowRxapprvd = 11;//RX approver

                        break;
                    case 14:
                        Console.WriteLine("DTB Electron");
                        templateLocation = @"\\"; //Where the template is sourced from.
                        destinationFolder = @"\\"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_Electron_PhysicsChecks.xlsx";
                        break;
                    case 15:
                        Console.WriteLine("DTB Field in Field");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\2. Field in Field Patient QA\! Template Do Not Delete\ACUROS  16100 -FinF Eclipse RTPS Physics Checks v6.xlsx"; //Where the template is sourced from.//templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\2. Field in Field Patient QA\! Template Do Not Delete\ACUROS  16100 -FinF Eclipse RTPS Physics Checks v6.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\2. Field in Field Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_FiF_PhysicsChecks.xlsx";

                        desired = 0.2; // desired DG size in cm.
                        colV = "D";
                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 11;
                        rowV_date1 = 33; //No date.
                        colV_date1 = "D";
                        rowV_dosecalalgo = 19;
                        rowV_DG = 20;
                        row_comment = 28;
                        colV_comment = "B";
                        rowPlanapprvd = 15;
                        rowHe = 21; //Hetereogenety is ON test.
                        row_BAlign = 25; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        rowV_pho_opt = 36; //No Photon Optimiser in FiF //PO_16100;
                        rowRxapprvd = 12;//RX approver

                        break;
                    case 16:
                        Console.WriteLine("DTB Hybrid");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\10. IMRT and VMAT Hybrid Plan Patient QA\! Template DO NOT DELETE\ACUROS 16100-Eclipse IMRT+VMAT Hybrid Plan Physics Checks V.6.xlsx"; //Where the template is sourced from.//templateLocation = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\4. IMRT and VMAT Hybrid Plan Patient QA\! Template\WCCD ACUROS 16100-Eclipse IMRT+VMAT Hybrid Plan Physics Checks V.3.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\4. IMRT and VMAT Hybrid Plan Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_Hybrid_PhysicsChecks.xlsx";

                        desired = 0.2; // desired DG size in cm.
                        row_BAlign = 26;
                        rowV_date1 = 43;
                        rowV_dosecalalgo = 20;
                        colV_date1 = "D";
                        rowPlanapprvd = 16;
                        rowHe = 23;
                        rowV_DG = 22;
                        row_comment = 28;
                        colV_comment = "B";

                        rowRxapprvd = 12;//RX approver
                        rowPlanapprvd = 16;
                        rowV_pho_opt = 21;
                        row_BAlign = 26; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).

                        break;
                    case 17:
                        Console.WriteLine("DTB HyperArc");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\PATIENT ID_ HyperArc ACUROS_16100-Eclipse Physics Checks V1.0.xlsm";//@"N:\Team\Orange\Radonc\Physics\2. Patient QA\12. HyperArc QA\! Template Do Not Delete\PATIENT ID_ HyperArc ACUROS_16100-Eclipse Physics Checks V1.0.xlsx"; //Where the template is sourced from. templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\12. HyperArc QA\! Template Do Not Delete\PATIENT ID_ HyperArc ACUROS_16100-Eclipse Physics Checks V1.0.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\9. SRT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_HyperArc_PhysicsChecks.xlsm";

                        desired = 0.1;
                        colV = "C";
                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 11;
                        rowV_date1 = 51;
                        colV_date1 = "D";
                        rowV_dosecalalgo = 28;
                        rowV_DG = 30;
                        row_comment = 35;
                        colV_comment = "B";
                        rowPlanapprvd = 22;
                        rowHe = 32; //Hetereogenety is ON test.
                        row_BAlign = 23; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).
                        rowV_pho_opt = 29; //PO_16100;
                        rowRxapprvd = 21;//RX approver

                        break;
                    case 18:
                        Console.WriteLine("DTB IMRT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\ACUROS 16100-Eclipse IMRT Physics Checks V.10.xlsm"; //@"N:\Team\Orange\Radonc\Physics\2. Patient QA\3. IMRT Patient QA\! Template Do Not Delete\ACUROS 16100-Eclipse IMRT Physics Checks V.10.xlsx"; //Where the template is sourced from. templateLocation = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\3. IMRT Patient QA\! Template Do Not Delete\ACUROS 16100-Eclipse IMRT Physics Checks V.10.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\3. IMRT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_IMRT_PhysicsChecks.xlsm";
                       
                        rowV_date1 = 42;
                        colV_date1 = "D";
                        //colV_comment = "B";
                        desired = 0.2; // desired DG size in cm.
                        rowPlanapprvd = 15;
                        rowHe = 22;
                        row_BAlign = 25;

                        rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 11;
                        rowV_dosecalalgo = 19;
                        rowV_DG = 21;
                        row_comment = 27;
                        colV_comment = "B";
                        rowV_pho_opt = 20; //PO_16100;
                        rowRxapprvd = 12;//RX approver

                        break;
                    case 19:
                        Console.WriteLine("DTB VMAT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\ACUROS_16100-Eclipse VMAT Physics Checks V1.8.xlsm"; // @"N:\Team\Orange\Radonc\Physics\2. Patient QA\5. VMAT Patient QA\! VMAT QA TEMPLATE\ACUROS_16100-Eclipse VMAT Physics Checks V1.8.xlsx"; //Where the template is sourced from. templateLocation = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\5. VMAT Patient QA\! VMAT QA TEMPLATE\WCCD ACUROS_16100-Eclipse VMAT Physics Checks V1.8.2.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\5. VMAT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_WCCD_ACUROS_16100-Eclipse_PhysicsChecks.xlsm";
                        
                        desired = 0.2; // desired DG size in cm.
                        rowV_dosecalalgo = 18;
                        rowV_DG = 20;
                        row_comment = 26;
                        colV_comment = "B";
                        rowV_pho_opt = 19; //PO_16100;



                        /*rowV_patId = 8;
                        rowV_planid = 10;
                        rowV_userid = 11;
                        rowV_date1 = 41;

                        rowRxapprvd = 12;//RX approver
                        rowPlanapprvd = 14;
                        colV_date1 = "D";
                        rowV_dosecalalgo = 18;
                        rowV_pho_opt = 19; //PO_16100;
                        rowV_DG = 20;
                        rowHe = 21; //Hetereogenety is ON test.
                        row_BAlign = 24; //(x = x1, y = y1, z=z1) for all treatment beams (same iso.)).


                        row_comment = 26;
                        colV_comment = "B";
                        desired = 0.2; // desired DG size in cm.


                        //Arrays for check-boxes only:
                        string[] Dubbo_row_strings = new string[] { "D", "D", "D", "D", "D", "D", "D" }; //These arrays must all contain the same number of elements.
                        int[] Dubbo_column_ints = new int[] { rowRxapprvd, rowPlanapprvd, rowV_dosecalalgo, rowV_pho_opt, rowV_DG, rowHe, row_BAlign };
                        string[] Dubbo_TRUE_false_checkbox = new string[] { Dubbo_Rxapprvd, Dubbo_Planapprvd, Dubbo_dosecalalgo, Dubbo_pho_opt, Dubbo_DG, Dubbo_He, Dubbo_BAlign };
                        string[] Dubbo_box_captions = new string[] { "", "", "", "", "", "", "" }; */

                        break;
                    case 20:
                        Console.WriteLine("DTB Liver SBRT");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\Template_ Liver SABR ACUROS_16100-Eclipse SABR Physics Checks V1.0.xlsm"; //@"N:\Team\Orange\Radonc\Physics\2. Patient QA\13. Liver SBRT Patient QA\! Template Do Not Delete\Template_ Liver SABR ACUROS_16100-Eclipse SABR Physics Checks V1.0.xlsx"; //Where the template is sourced from. templateLocation = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\10. Liver SBRT Patient QA\! Template Do Not Delete\Template_ Liver SABR ACUROS_16100-Eclipse SABR Physics Checks V1.0.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\10. Liver SBRT Patient QA";
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_LiverSABR_PhysicsChecks.xlsm";
                                                
                        desired = 0.125; // desired DG size in cm.
                        rowV_dosecalalgo = 29;
                        rowV_DG = 31;
                        row_comment = 35;
                        colV_comment = "B";
                        rowPlanapprvd = 20;
                        rowHe = 32;
                        rowV_date1 = 12;
                        colV_date1 = "D";
                        //colV_date2 = "C";
                        row_BAlign = 21; //Needs to also check that it's contained WITHIN PTV Structure.
                        rowV_pho_opt = 30;
                        rowRxapprvd = 19; //RX approver by RO.
                        break;
                    case 21:
                        Console.WriteLine("DTB Non-Liver SABR");
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\!ScriptTemplates_DoNotModify\Truebeam\Template_ SABR ACUROS_16100-Eclipse SABR Physics Checks V2.0.xlsm"; //@"N:\Team\Orange\Radonc\Physics\2. Patient QA\11. Non-Liver SABR Patient QA\! Template Do Not Delete\Template_ SABR ACUROS_16100-Eclipse SABR Physics Checks V2.0.xlsx"; //Where the template is sourced from.
                        //templateLocation = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\9. SRT Patient QA\!Template DO NOT DELETE\Template_ SABR ACUROS_16100-Eclipse SABR Physics Checks V2.0.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\9. SRT Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_ACUROS_16100-Eclipse_SABR_PhysicsChecks.xlsm";

                        
                        rowPlanapprvd = 19;
                        desired = 0.125; // desired DG size in cm.
                        rowV_dosecalalgo = 25;
                        rowHe = 28;
                        row_comment = 31;
                        rowV_date1 = 47;
                        colV_comment = "B";
                        colV_date1 = "D";
                        rowV_DG = 27;
                        row_BAlign = 20; //Needs to also check that it's contained WITHIN PTV Structure.
                        rowV_pho_opt = 26;
                        rowRxapprvd = 18; //RX approver by RO.

                        break;
                    case 22:
                        Console.WriteLine("Halcyon HYBRID"); //!!! Halcyon hybrid plans technically possible!!!? //
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\15. Halcyon Patient QA\!Halcyon_VMAT Patient QA\! Template Do Not Delete\ACUROS_16100-Eclipse VMAT Physics Checks_ Halcyon V1.8.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\2. Patient QA\15. Halcyon Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"OHAL_HYBRID_{patient.Id}_{_planId}_ACUROS_16100-Eclipse_HalcyonHybrid_PhysicsChecks.xlsx";
                        break;
                    case 23:
                        Console.WriteLine("Dubbo Ortho"); //
                        templateLocation = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\6. Orthovoltage Patient QA\Template - DO NOT DELETE\WCCD AAPM TG61 Orthovoltage MU Check v.3.4 - template.xlsx"; //Where the template is sourced from.
                        destinationFolder = @"N:\Team\Orange\Radonc\Physics\1.  Documents\Dubbo WCCD plans\1 WCCD Patient QA\6. Orthovoltage Patient QA"; //Where a copy of the template is to be put. Also create a new directory if one does not already exist with MRN.
                        filename = $"{patient.Id}_{_planId}_WCCD_AAPM_TG61_OrthovoltagePhysics_MU_Check.xlsx";
                        break;
                }
            }
            //
            Console.WriteLine(planchecktype.ToString());
            //Console.ReadLine();

            

            string fn1 = CreateFolderPtID(_patientId, destinationFolder, _courseId, _planId); //Create folder - if not already present.
            string fn = String.Join(@"\", fn1, filename);
            Console.WriteLine(fn1);
            Console.WriteLine(fn);
            //**Now need to copy the template spreadsheet to the folder.**//
            MoveTemplateSpreadsheet(templateLocation, fn);


            //
            WriteToCell(fn, colV, rowV_patId, _patientId); //Now write to the spreadsheet, the patient ID.
            WriteToCell(fn, colV, rowV_planid, _planId); //planID.

            //Console.WriteLine(_currentUserId);
            //Console.WriteLine(_usern);

            //Check if is Eugene
            if (_currentUserId == @"nswhealth\60027842" || _currentUserId == "60027842")
            {
                WriteToCell(fn, colV, rowV_userid, "ET");
            }
            else
            {
                WriteToCell(fn, colV, rowV_userid, CheckingFunctions.initialS(_usern, _currentUserId));
            }

            WriteToCell(fn, colV_date1, rowV_date1, CheckingFunctions.Getdayte()); //Give the date

            string valuenotation = "\u2071 Checked by AutoPhysics script: version 2.12, script approved on 07/07/2025.";
            WriteToCell(fn, colV_comment, row_comment, valuenotation); //Write the label for things check automatically. row_comment


            //Check all the things and write to sheet as go:
            var vartemp = CheckingFunctions.GetHetreogeneity(planSelec);
            Console.WriteLine(vartemp);

            //Check Photon Optimiser:

            string pho_opt = CheckingFunctions.GetOptimiserCalcAlgo(planSelec);
            if (String.Equals(pho_opt, "PO_16100"))
            {
                WriteToCell(fn, colV_date1, rowV_pho_opt, ticksymbol); //dosecalalgo
                Dubbo_pho_opt = "TRUE";
            }
            else
            {
                WriteToCell(fn, colV_date1, rowV_pho_opt, crossymbol + $" {pho_opt}");
            }

            //Check Photon Dose calculation algorithm.
            string dgca = CheckingFunctions.GetVolDoseCalcAlgo(planSelec);

            // Console.ReadLine();
            if (String.Equals(dgca, desiredphotoncalcalgo))
            {
                WriteToCell(fn, colV_date1, rowV_dosecalalgo, ticksymbol); //dosecalalgo
                Dubbo_dosecalalgo = "TRUE";
            }
            else
            {
                WriteToCell(fn, colV_date1, rowV_dosecalalgo, crossymbol);
            }

            bigiftrue = false;
            //Dosegrid:
            if (planchecktype == 17 | planchecktype == 5) //SRS HyperArc for DTB/OTB. Then also need to check the SRSHyperArc DG.
            {
                bigiftrue = (CheckingFunctions.CheckDoseGirdSize(desired, CheckingFunctions.GetDoseGird(planSelec)) & CheckingFunctions.CheckDoseGirdSize(desired, CheckingFunctions.GetHyperArcDoseGird(planSelec)));
            }
            else
            {
                bigiftrue = CheckingFunctions.CheckDoseGirdSize(desired, CheckingFunctions.GetDoseGird(planSelec));
            }
            //string value_dg = "FAIL";
            if (bigiftrue == true)
            {
                WriteToCell(fn, colV_date1, rowV_DG, ticksymbol);
                Dubbo_DG = "TRUE";
            }
            else
            {
                WriteToCell(fn, colV_date1, rowV_DG, crossymbol + "Incorrect dose-grid size for protocol AND OR HyperArc Dose-grid differs. Review.");
            }


            //Perscription Approval, and checks last modifier is an RO (authorised user);

            string[] rx_stat = CheckingFunctions.GetPersciptionStatus(courSelec);
            bool boolget = rx_stat[2].ToString().ToLower().Contains("true");

            // Console.WriteLine(rx_stat[2], rx_stat[0], rx_stat[1]);

            if (boolget == true)
            {
                WriteToCell(fn, colV_date1, rowRxapprvd, ticksymbol);
                Dubbo_Rxapprvd = "TRUE";
            }
            else
            {

                WriteToCell(fn, colV_date1, rowRxapprvd, crossymbol);
                

            }





            //Treatment Approval/Reviewed Status, and ifdone by authorised user (An RO):
            //string combinedString = string.Empty;
            //string combinedString2 = string.Empty;
            //string[] aprvl = { combinedString, combinedString2 };
            //aprvl = CheckingFunctions.GetPlanApprvlHistry(planSelec); //Returns string (joined array) ; with first containing plan approval status (well the history thereof); and the second containing the person associated with each change in the status.

            bool reviewerTrue = CheckingFunctions.PlanReviewedOrApproved(planSelec); //Ensure plan-reviewer is in authorised list AND plan not subsequently edited.

            Console.WriteLine(reviewerTrue);


            reviewer = CheckingFunctions.GetNameOfPlanReviewerOrApprover(planSelec); //returns empty string if plan reviewer/approver is not in authorised list.


            //bool isPln_apprvd = //CheckingFunctions.PlanReviewedOrApproved(combinedString, combinedString2); //Returns true if reviewed/approved by appropriate person.
            if (string.IsNullOrEmpty(reviewer))
            {
                WriteToCell(fn, colV_date1, rowPlanapprvd, crossymbol);
            }
            else if (reviewerTrue == true)
            {

                WriteToCell(fn, colV_date1, rowPlanapprvd, ticksymbol);
                Dubbo_Planapprvd = "TRUE";

            }


            // reviewer = CheckingFunctions.GetNameOfPlanReviewerOrApprover(planSelec);
            // Console.WriteLine(reviewer);

            //Hetreogeniety correction ON/OFF
            string hetrogenitycor = CheckingFunctions.GetHetreogeneity(planSelec);
            if (CheckingFunctions.GetHetreogeneity(planSelec).Contains("ON"))
            {
                WriteToCell(fn, colV_date1, rowHe, ticksymbol);
                Dubbo_He = "TRUE";
            }
            else
            {
                WriteToCell(fn, colV_date1, rowHe, crossymbol);
                Console.WriteLine(CheckingFunctions.GetHetreogeneity(planSelec));
            }

            //Autofeathing correction ON/OFF
            string af = CheckingFunctions.GetAutoFeather(planSelec);
            if ((af.Contains("AutoFeathering, On") && (planchecktype == 11 | planchecktype == 12)))
            {
                WriteToCell(fn, colV_date1, rowAF, ticksymbol);
            }
            else if (planchecktype == 11 | planchecktype == 12)
            {
                WriteToCell(fn, colV_date1, rowAF, crossymbol);
                Console.WriteLine(CheckingFunctions.GetHetreogeneity(planSelec));
            }




            //If IMRT than do the lostMU factor searchings and writings:
            string ss = string.Empty;
            string ss2 = string.Empty;
            string ss3 = string.Empty;

            if (isIMRT | planchecktype == 4)
            {
                int indx = 0;
                foreach (var a in planSelec.Beams.Where(x => !x.IsSetupField).Reverse()) // Exclude setupfields and Reverse order to get it in format users expect.
                {
                    if (a.MLCPlanType.ToString().Contains("DoseDynamic"))
                    {
                        try
                        {


                            double MU = CheckingFunctions.GetBeamMU(a);
                            string lostMUfactor = CheckingFunctions.GetBeamlogs(a, "Information: Lost MU factor for carriage group");
                            ss = lostMUfactor.Substring(lostMUfactor.LastIndexOf("= ") + 1);
                            if (lostMUfactor == string.Empty)
                            {
                                lostMUfactor = CheckingFunctions.GetBeamlogs(a, "Information: LostMUFactor"); //Such as Hybrid for OTB1 case.
                                ss = lostMUfactor.Substring(lostMUfactor.LastIndexOf("r ") + 1);
                            }

                            string MaxMU = CheckingFunctions.GetBeamlogs(a, "Maximum MU for carriage group");
                            ss2 = MaxMU.Substring(MaxMU.LastIndexOf("= ") + 1);

                            Console.WriteLine(ss);
                            Console.WriteLine(ss2);

                            if (MaxMU == string.Empty)
                            {
                                MaxMU = CheckingFunctions.GetBeamlogs(a, "Information: Maximum MUs"); //Such as Hybrid for OTB1 case.
                                ss2 = MaxMU.Substring(MaxMU.LastIndexOf(": ") + 1);
                            }

                            CheckingFunctions.WriteToCellOnSheet(fn, 2, "A", (3 + indx), a.Id); //Nope, just pull the field ID. //  Write the beam Name to Spreadsheet. B.C. when we iterate 0,1 is highest number field ID. a.BeamNumber.ToString()
                            CheckingFunctions.WriteToCellOnSheet(fn, 2, "B", (3 + indx), MU.ToString());
                            CheckingFunctions.WriteToCellOnSheet(fn, 2, "C", (3 + indx), ss);
                            CheckingFunctions.WriteToCellOnSheet(fn, 2, "D", (3 + indx), ss2);

                            //Now for Jaw-tracking:
                            if (tmtmach != "OHAL1")
                            {
                                string JT = CheckingFunctions.GetBeamlogs(a, "Information: Jaw Tracking: "); //Will be "Information: Jaw Tracking: disabled"; if not turned on. or  enabled, if turned on.
                                ss3 = JT.Substring(JT.LastIndexOf(": ") + 2); //Can test if this is enabled or disabled - then just need to show where it is failed in the plan.
                                if (JT != string.Empty)
                                {
                                    CheckingFunctions.WriteToCellOnSheet(fn, 2, "G", (3 + indx), JT);
                                    if (ss3 == "disabled")
                                    {
                                        JT_isOn_allBeams = false;
                                        break;
                                    }
                                }
                            }


                            //Console.WriteLine(lostMUfactor.Substring(lostMUfactor.LastIndexOf("= ")+1));

                            //Console.WriteLine($"The mus for beam {a.Name} is {MU}, the lostmufactor(ornear?) is {lostMUfactor}, the maximumMUs are {MaxMU}.");
                            //Then write these to sheet for each beam, on "Lost MU Check" excel sheet.
                            //BeamMU is "B3" (up to "B7"); lostMU starts at "C3", and MaxMU starts at "D3";
                            //indx++;
                        }

                        catch
                        { //must be a setup field or non-imrt beam. Do nothing.
                            Console.WriteLine($"{a.Name} is a Non-IMRT beam; OR LMC in plan not run!! please check.");

                        }


                    }

                    indx++; //still want to increment, in-case setupfield or non-IMRT beam is in the middle of the fields.

                }
            }


            
            //Need to check jaw-tracking for all Truebeam vmat fields (VMAT/SABR/SRS):
            if (planchecktype == 7 | planchecktype == 8 | planchecktype == 9 | planchecktype == 17 | planchecktype == 19 | planchecktype == 21)
            {
                int indx = 0;
                foreach (var a in planSelec.Beams.Where(x => !x.IsSetupField).Reverse()) // Exclude setupfields
                {
                    if (a.MLCPlanType.ToString().Contains("VMAT"))
                    {
                        try
                        {
                            JT_isOn_allBeams = false; // fail it, unless found in logs.
                            //Jaw-tracking:
                            string JT = CheckingFunctions.GetBeamlogs(a, "Information: Jaw tracking is "); //Will be "Information: Jaw Tracking: disabled"; if not turned on. or  enabled, if turned on.
                            ss3 = JT.Substring(JT.LastIndexOf("is ") + 3); //Can test if this is enabled or disabled/on or off - then just need to show where it is failed in the plan.
                            //Console.WriteLine(ss3);
                            //Console.WriteLine(JT);
                            if (ss3 == "on.")
                            {
                                JT_isOn_allBeams = true;
                            }
                            else if (ss3 == "off.")
                            {
                                JT_isOn_allBeams = false;
                                break; //Terminates the foreach loop. And does so with JT test set to false.
                            }



                        }

                        catch
                        { //must be a setup field or non-imrt beam. Do nothing.
                            Console.WriteLine($"{a.Name} Jaw-tracking check failed for this beam. Please Check manually.");

                        }


                    }

                    indx++; //still want to increment, in-case setupfield or non-IMRT beam is in the middle of the fields.

                }
            }





            //Beam position is same spot. //Note for HyperArc this check-list item is to also ensure the beams are all 6FFF AND at same spot, however to have identified the correct plan-check type
            //The treatment technique MUST have been HyperArc; which means the beams are 6FFF (unless somehow somebody made a mixed energy plan - which think you can't do).
            bigiftrue = CheckingFunctions.beamAlign(planSelec);
            //string value_dg = "FAIL";
            if (bigiftrue == true)
            {
                WriteToCell(fn, colV_date1, row_BAlign, ticksymbol);
                Dubbo_BAlign = "TRUE";
            }
            else
            {
                WriteToCell(fn, colV_date1, row_BAlign, crossymbol);
            }

            ///Now write all the values too sheet 3 (or 1 in case of liver sabr/non-liver sabr, which is case 8/9) of the workbook, so can review manually output of script. Or Hybrid.
            int sheettowrite = 3;
            if (planchecktype == 8 | planchecktype == 9 | planchecktype == 20 | planchecktype == 21)
            {
                sheettowrite = 2;
            }


            //MLC is VMAT or Dose-dynamic (Halcyon tests only). Also check that treatment machine is same for all beams (and is halcyon).
            if (planchecktype == 12 | planchecktype == 11) //Halcyon VMAT | Hal IMRT
            {
                if (CheckingFunctions.AllBeamsDoseDynamicOrVMAT(planSelec) == true)
                {
                    WriteToCell(fn, colV_date1, row_mlc, ticksymbol);
                }
                else
                {
                    WriteToCell(fn, colV_date1, row_mlc, ticksymbol);
                }

                if (allsametmtmach = CheckingFunctions.AllBeamsSameTmtMachine(planSelec, "OHAL1") == true)
                {
                    WriteToCell(fn, colV_date1, row_samemach, ticksymbol);
                }
                else
                {
                    WriteToCell(fn, colV_date1, row_samemach, ticksymbol);
                }





            }


            ///

            //


            ///Define row,col and values arrays:
            string[] values = { $"{rx_stat[0]}, {rx_stat[1]}, {rx_stat[2]}", reviewer, "", "", dgca, CheckingFunctions.GetOptimiserCalcAlgo(planSelec), CheckingFunctions.GetDoseGird(planSelec).ToString(), hetrogenitycor, CheckingFunctions.beamAlign(planSelec).ToString(), CheckingFunctions.GetHyperArcDoseGird(planSelec).ToString(), CheckingFunctions.GetAutoFeather(planSelec), JT_isOn_allBeams.ToString()};

            string[] testitems = { "The prescription is approved by RO", "The plan is Reviewed or Approved by RO", "Check the Convolution Kernel is Qr40f and Imaging device is Orange CT", "Physical Material table -Acuros XB-13.5", "The dose calculation algorithm is Acuros_16100", "The dose optimiser is PO_16100", "The dose grid resolution is 0.2 cm (or <= 0.125 cm for SABR)", "Heterogeneity Correction in ON", "Beam placement all aligned with the same point", "SRS/HyperArc-Dose grid (applicable for HA plans) is: ", "Autofeathering settings (if applicable for OHAL): ", "Jaw tracking is on for all Truebeam VMAT/IMRT fields (Ignore for OHAL1): " }; //s = {"A", "A"};
            //Array.Copy(testitems, rows, values.Length); //{testitems.Length}; //"A"
            string[] rows = Enumerable.Repeat("A", values.Length).ToArray(); //These lines generate row-column values of A1,A2,A3,...,etc... for use in writing to sheet3.
            string[] rows2 = Enumerable.Repeat("B", values.Length).ToArray();
            int[] cols = Enumerable.Repeat(1, values.Length).ToArray();
            int count = 0;
            foreach (int a in cols) {
                cols[count] = count + 1;
                count++;
            }

            //;
            foreach (var a in values)
            {
                Console.WriteLine(a);
            }


            //Now write them to sheet 3 of excel, in cells A1, through to Ax;
            CheckingFunctions.WriteLotsOfValuesToCellsOnSheet(fn, sheettowrite, rows, cols, testitems);
            CheckingFunctions.WriteLotsOfValuesToCellsOnSheet(fn, sheettowrite, rows2, cols, values);


            //**Dubbo Section **//
            //This region we write the specific values required for the check-boxes for Dubbo sheets:



          /*  if (planchecktype > 12) //if plan-check type case is 13 or above - it's a Dubbo plan - uses Dubbo spreadsheets. Write the check-boxes:
            {
                string[] Dubbo_TRUE_false_checkbox = new string[] { Dubbo_Rxapprvd, Dubbo_Planapprvd, Dubbo_dosecalalgo, Dubbo_pho_opt, Dubbo_DG, Dubbo_He, Dubbo_BAlign }; //Need to redeclare, now that values have been changed.
                CheckingFunctions.WriteLotsOfTickBoxs_Dubbo_ToCells(fn, Dubbo_row_strings, Dubbo_column_ints, Dubbo_TRUE_false_checkbox, Dubbo_box_captions);

            }
          */
            ////




            //**Other Functions**//


            //**Function for the creation of the patient folder in the relevant patient QA directory**//

            string CreateFolderPtID(string pat_ID, string folderlocDestination, string courseId, string planId)
            {   //input the patient ID, the destination that the folder should go to, and the planname (if required for multiple courses/plans).

                //Check if location already has a folder with name matching pat_ID; if not create one.
                //First get all folder names in folderlocDestination directory.
                // string[] foldaNamez = Directory.GetDirectories(folderlocDestination);
                //Join folderlocationDestinationPath together:
                string[] arayy = { folderlocDestination, pat_ID };
                string[] arayy2 = { folderlocDestination, pat_ID, courseId, planId };
                string newdir = String.Empty;
                string sep = @"\";
                string jj = String.Join(sep, arayy);
                string jj2 = String.Join(sep, arayy2);
                //var newdir = new DirectoryInfo(jj); //Only needed because I've poorly constructed the return values in the "if statement".

                if (!Directory.Exists(jj))
                {
                    Directory.CreateDirectory(jj); //if it doesn't exist, create this one.
                                                   // var newdir2 = new DirectoryInfo(jj);
                    newdir = jj;
                }
                else if (!Directory.Exists(jj2))
                {
                    Directory.CreateDirectory(jj2); //if it doesn't, check that an alternative one hasn't been made (e.g. script already been run). If it hasn't then create another directory,
                                                    //with format folderlocDestionat/patientID/courseId/planId; % Should be unique each time for each new plan/course.
                                                    //var newdir2 =  new DirectoryInfo(jj2);
                    newdir = jj2;
                }
                else
                {
                    Console.WriteLine("Directories already exist. Will not create a new one.");
                    //var newdir2 =  new DirectoryInfo(jj);
                }
                // return newdir;
                return newdir;

            }

            //  == pat_ID





            //**Move copy of template spreadsheet there. **//
            void MoveTemplateSpreadsheet(string templatelocation, string copyNameIncludingFullPath)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(templatelocation, true);

                workbook.SaveAs(copyNameIncludingFullPath);
                workbook.Close(copyNameIncludingFullPath);

            }



            //** Write values to copy of spreadsheet **// 
            //** Define function to write the values to copy of the excel document **//

            void WriteToCell(string fileName, string row, int col, string value)
            {
                //e.g. WriteToCell('myworkbook.xls', 'A', 1, 'write this value there');
                //  using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(fileName, true))
                // {
                //     //  WorksheetPart worksheetPart = GetWorksheetPartByName(spreadsheet, "sheet1");

                //     spreadsheet.WorkbookPart
                //     WorksheetPart worksheetA = spreadsheet.WorkbookPart();
                // }
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Open(fileName, true);
                //Excel.Workbook workbook = excelApp.WorkbookOpen(fileName);
                //Excel.Worksheet worksheet1 = ((Excel.Worksheet)Application.ActiveWorkbook.Worksheets[1]);
                Excel.Worksheet worksheet = (workbook.Application.ActiveWorkbook.Worksheets[1]); //Should give first active worksheet.

                //now write too it:
                // worksheet.Cell(row, col).Value = value;
                string colstring = col.ToString();
                string rss = String.Join(string.Empty, row, colstring); //Aren't in write format for function import.

                Excel.Range ra = worksheet.Range[rss]; //get the range.
                ra.Value2 = value; //Now write it to the desired.

                workbook.Save();
                //Savin' & closing time.
                workbook.Close();

                //  var workbook = new XLWorkbook(fileName);
                // var worksheet = workbook.Worksheets.Worksheet(1);
                //worksheet.Cell(row, col).Value = value;
                //workbook.SaveAs(fileName);
            }


            //

            //isOHAL = tmtmach.Contains("OHAL");



            //** if NOT ortho, export that plan to Mobius.**//
            if (planchecktype != 23 && planchecktype != 10) //(tmtmach != "SXR")
            {
                //Send to Mobius here!

            }




            // var bbbb = planSelec.PhotonCalculationOptions.ToArray();

            // foreach (var a in bbbb)
            //{
            //PO = plan.PhotonCalculationOptions.ToArray()[1].ToString();
            //   Console.WriteLine(a.ToString());
            //
            // }

            // Console.ReadLine();
            //var bbbb = plan.GetCalculationOptions(OptimizationOption)
            //var bbbb = planSelec.GetCalculationOptions("PO_16100");

            // foreach (var a in bbbb)
            // {
            //PO = plan.PhotonCalculationOptions.ToArray()[1].ToString();
            //    Console.WriteLine(a.ToString());

            // }

            // Console.ReadLine();







            try
            {
                //LOGGING: write to logging spreadsheet:
                int logR = CheckingFunctions.GetEmptyCellRow(@"N:\Team\Orange\Radonc\Physics\Projects\039_AutoPhysics_PlanChecks\plan-checkScriptTesting_Log_v1.xlsx", 1);
                Console.WriteLine(logR);
                //Console.ReadLine();

                string[] values2 = { $"{CheckingFunctions.Getdayte()}, {CheckingFunctions.initialS(_usern, _currentUserId)}, {_patientId}, {_planId}, {planchecktype}" };
                string[] cols2 = Enumerable.Repeat("A", values2.Length).ToArray();

                //string[] cols2 = new string[5] { "A", "B", "C", "D", "E" };
                int[] rows3 = { logR, logR, logR, logR, logR };

                CheckingFunctions.WriteLotsOfValuesToCellsOnSheet(@"N:\Team\Orange\Radonc\Physics\Projects\039_AutoPhysics_PlanChecks\plan-checkScriptTesting_Log_v1.xlsx", 1, cols2, rows3, values2);

                //Console.ReadLine();



                //If it was successful in creating file in folder location, open said floc:
                if (File.Exists(fn))
                {
                    Process.Start("explorer.exe", fn1); //Open Folder loc
                    Process.Start("explorer.exe", fn); //Open actual file. Let's GO!!
                }

            }
            catch {

                Console.WriteLine("Logging to log spreadsheet failed. Maybe somebody moved it or has it open.");
            }

        }
    }
}
