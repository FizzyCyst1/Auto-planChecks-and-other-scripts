/*******************************************************************************************************
* Namespace:       SunDoseProfileExport
* Application:     SunDoseProfileExport
* Purpose:         To export beam profiles from TPS in Sun Dose friendly format
* Author:          Joshua Hiatt
* Date:            28/5/2024
* Comments:        
* 
* 
* 
* *****************************************************************************************************
* Checked BY:      Scott Piggott
* Date:            29/05/2024
* Comments:        
* 
* 
*  *****************************************************************************************************/

//******************************************************************************************************
//04/10/2024: bug-fix, when selecting Crossplane profiles, using max normalisation; was exporting in-line profiles. Have since fixed issue.

//04/07/2024: Added electron functionality, PDI to PDD = true for electrons.
// No longer writes values to file for profile/PDD points that have NaN value.
// Pulls applicator size when electrons.
//
// Minor edits on 04/06/2024, By Joshua Hiatt. Added "-FFF" to string replace, for naming of energies for OHAL1 plans.
// Added correct field-sizes pulled from plans. And SSD. To parameters in SunDose format.

//24/06/2024 - Fixed a bug affecting asymmetric field-sizes, whereby the crossline profiles would be incorrect and would be duplicates of the inline.
//Due to the way C# handles DoseProfile variables. Fixed by redefining each dose-profile prior to export.
//*****************************************************************************************************


using System;
using System.Linq;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.IO;
using System.Windows.Documents;
using System.Collections.Generic;
using System.Windows.Media.TextFormatting;
using System.Xml.Linq;




/*** Ensure script is not writeable ****/
//[assembly: ESAPIScript(IsWriteable = true)]

namespace SunDoseProfileExport2
{
    class Program
    {
        #region Local Variables
        private static string _patientId; //private (to this class, so only visible to this class/?method).
        private static string _courseId;
        private static string _planId;
        private static string _structureSetId;
        private static string _currentUserId;
        private static string type = string.Empty;
        private static string tmtmach = string.Empty;
        private static string ssd = string.Empty;
        private static string xfs = string.Empty;
        private static string yfs = string.Empty;
        private static string xfs_mlc = string.Empty;
        private static string yfs_mlc = string.Empty;
        private static string mlc2 = string.Empty;
        private static string sumx_mlc = string.Empty;
        private static string sumy_mlc = string.Empty;
        private static string modality = string.Empty;
        static string tf_pdi = string.Empty;
        private static string rtype = string.Empty;
        private static bool isElectron = false;
        private static bool NilMLC = false;
        private static int symdist = 0;
        private static DoseProfile proff;
        private static DoseProfile proffI;
        private static double normaval = 1;
        private static double normavalINPLANE = 1;
        private static int conttta = 0;
        private static int pdproSa = 5;
        private static int pdproSa_pdd = 5;
        private static double maxyv = 0;
        private static double maxyvv = 0;
        private static int contta = 0;
        private static double mvI = 0;
        private static double mvv = 0;
        private static double[] usedis;
        private static string destinationFolder = @"\\Ordcmdc-mis001.nswhealth.net\wnswlhd\Team\Orange\Radonc\RT - Physics\TPS_Scans_ForSunDoseImport\"; //Not this one: //@"C:\temp\TPS_Scans_ForSunDoseImport\"; //;
        //ThIS one works if you want it direct (non-citrix mode:)
        //private static string destinationFolder = Path.Combine(Path.GetTempPath(), "TPS_Scans_Exports");


        private static int dista = 1;
        private static int stepsize = 1;

        /*** Set ESAPI variables ****/
        private static Patient patient;
        private static Course courSelec;
        private static PlanSetup planSelec;
        private static ReferencePoint refSelec;
        private static VVector reffy;
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
                    if (args.First().Split(';').Count() == 5)
                    {
                        _patientId = args.First().Split(';').First();
                        _courseId = args.First().Split(';').ElementAt(1);
                        _planId = args.First().Split(';').ElementAt(2);
                        _structureSetId = args.First().Split(';').ElementAt(3);
                        _currentUserId = args.First().Split(';').Last();
                    }
                }

                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());

            }
            Console.ReadLine();
        }

        /// <summary>
        /// Executes application code taking ESAPI application as parameter
        /// </summary>
        /// <param name="app">ESAPI application</param>
        static void Execute(Application app)
        {
            /*** Check for patient ID otherwise ask user for it ****/
            if (String.IsNullOrEmpty(_patientId))
            {
                Console.WriteLine("Version 1.04.");
                Console.WriteLine("Modified export folder to Cdrive of local server.");
                Console.WriteLine("This script application allows you to export dose profiles and or PDDs from a plan into a format compatible for import into SunDose.");
                Console.WriteLine("///Requirements: Plan must contain a reference point (this will act as the start location for the relevant profile) within the body and dose must be calculated ////.");
                Console.WriteLine("  "); //CW+TAB+TAB;
                Console.WriteLine("Enter patient Id:"); //CW+TAB+TAB;
                _patientId = Console.ReadLine();
            }
            /*** Open patient with ID supplied ****/
            Patient patient = app.OpenPatientById(_patientId);
            if (patient == null) { Console.WriteLine($"No patient with Id {_patientId}"); return; }
            Console.WriteLine($"Patient open: {patient.Name}");
            Console.WriteLine("Version 1.04.");
            Console.WriteLine("Modified export folder to Cdrive of local server.");
            Console.WriteLine("This script application allows you to export dose profiles and or PDDs from a plan into a format compatible for import into SunDose.");
            Console.WriteLine("///Requirements: Plan must contain a reference point (this will act as the start location for the relevant profile) within the body and dose must be calculated ////.");
            //Console.WriteLine("Version 1.02.");
            //Directory.CreateDirectory(destinationFolder);
            /*** Set course to null and attempt to retrieve via plan ID if not null ****/
            courSelec = null;
            if (!string.IsNullOrEmpty(_courseId))
            {
                courSelec = patient.Courses.FirstOrDefault(x => x.Id.Equals(_courseId));
            }
            else
            {
                /*** Display courses for user to select ****/
                var courseIDList = new List<string>();
                foreach(var crse in patient.Courses) courseIDList.Add(crse.Id);
                int ccSelect = GetUserListSelection(courseIDList);

                /*** Display Selected Course ****/
                courSelec = patient.Courses.ElementAt(ccSelect);
                Console.WriteLine($"You have selected, {ccSelect}. - {courSelec}");
            }

            /*** Set course to null and attempt to retrieve via plan ID if not null ****/
            planSelec = null;
            if (!string.IsNullOrEmpty(_planId))
            {
                planSelec = courSelec.PlanSetups.FirstOrDefault(x => x.Id == _planId);
            }
            else
            {
                /*** Display plans for user to select ****/
                var planIDList = new List<string>();
                foreach (var plan in courSelec.PlanSetups) planIDList.Add(plan.Id);
                int ppSelect = GetUserListSelection(planIDList);

                /*** Display Selected Plan ****/
                planSelec = courSelec.PlanSetups.ElementAt(ppSelect); //Plan selected
                Console.WriteLine($"You have selected, {ppSelect}. - {planSelec}");
            }

            /*** Select Reference Point ****/
            refSelec = null;
            /*** Display plans for user to select ****/
            var referenceIDList = new List<string>();
            foreach (var ccr in patient.ReferencePoints)
            {
                string refPt = (ccr.HasLocation(planSelec)) ? ccr.Id : $"{ccr.Id} - INVALID - is not contained in this plan.";
                referenceIDList.Add(refPt);
            }
            int rpSelect = GetUserListSelection(referenceIDList);

            /*** Get the selected reference point ****/
            refSelec = patient.ReferencePoints.ElementAt(rpSelect); 
            Console.WriteLine($"You have selected, {rpSelect}. - {refSelec}");

            Boolean validref = false;
            if (!refSelec.HasLocation(planSelec))
            {
                validref = (refSelec == patient.ReferencePoints.ElementAt(rpSelect));
                Console.WriteLine("This reference point is not contained within the selected plan. Please select another."); return;
            }

            /*** Ask users to select from list of output types ****/
            Console.WriteLine($"Would you like a PDD or profiles?");
            string[] cwl = { "PDD", "Both In-Plane/Crossplane Profiles", "Crossplane", "Inplane" };
            var cwlList = new List<string>(cwl);
            int pdproS = GetUserListSelection(cwlList);

            /*** Select selected string for export type ****/
            var pdproSelec = cwl[pdproS]; //Option selected
            Console.WriteLine($"You have selected, {pdproS}. - {pdproSelec}");

            /*** Ensure plan is set to absolute dose mode ****/
            planSelec.DoseValuePresentation = DoseValuePresentation.Absolute;

            /*** Set the reference point location as VVector ****/
            reffy = refSelec.GetReferencePointLocation(planSelec);

            /*** Run sub when PDDs selected ****/
            if (pdproS == 0)
            {
                ProcessPDD();
            } else
            {
                /*** Get the distance for profile ****/
                Console.WriteLine($"Please input the TOTAL distance in millimeters for the profile (e.g. if you wish to have a total profile of 12 cm, input in 120. This will give the dose profile +60 mm to either side of the chosen reference point.");
                int dist = -1; 
                if (!int.TryParse(Console.ReadLine(), out dist))
                {
                    /*** Ensure entry is numeric otherwise error ****/
                    Console.WriteLine("Invalid Selection"); return; 
                }
                /*** Entry is in millimetres ****/
                 symdist = (dist / 2);

                /*** Get Profile step size  ****/
                Console.WriteLine($"Please input the profile step-size in 'mm', e.g. 1 mm, 2 mm.");
                int dista = -1; 
                if (!int.TryParse(Console.ReadLine(), out dista))
                {
                    /*** Ensure entry is numeric otherwise error ****/
                    Console.WriteLine("Invalid Selection"); return; 
                }
                /*** Entry is in millimetres ****/
                stepsize = dista;

              

                /*** Get inline start and stop points ****/
                VVector startInline = new VVector(reffy.x, reffy.y, (reffy.z + symdist));
                VVector stopInline = new VVector(reffy.x, reffy.y, (reffy.z - symdist));
                usedis = new double[(int)Math.Ceiling((startInline - stopInline).Length / stepsize)];

                /*** Get Profile using start and stop points and double array ****/
                proffI = planSelec.Dose.GetDoseProfile(startInline, stopInline, preallocatedBuffer: usedis);

                /*** Display inplane profile point value to user ****/
                Console.WriteLine($"Inplane profiles (assuming HFS setup) are [Gy]:");

                conttta = 0;
                maxyv = 0;
                mvI = 0;
                foreach (var v in proffI)
                {
                    Console.WriteLine($"{proffI[conttta].Value}");
                    if ((double)proffI[conttta].Value > maxyv)
                    {
                        mvI = proffI[conttta].Value;
                        maxyv = mvI;

                    }
                    conttta++;
                }

                Console.WriteLine($" ");
                Console.WriteLine($" ");

                /*** Display crossplane profile point to user ****/
                Console.WriteLine($"Crossplane profiles (assuming HFS setup) [Gy]:");

                contta = 0;
                maxyvv = 0;
                mvv = 0;

                /*** Add the vectors (writing the start/stop points in vector format) to get start/stop points from the ref point vector). ****/
                /*** ASSUMES HFS and reffy is in DICOM co-ordinates ****/
                VVector startCrossline = new VVector((reffy.x - symdist), reffy.y, reffy.z);
                VVector stopCrossline = new VVector((reffy.x + symdist), reffy.y, reffy.z);

                /*** calculate double array for points ****/
                usedis = new double[(int)Math.Ceiling((startCrossline - stopCrossline).Length / stepsize)];

                /*** Co-ordinates are in DICOM (pretty sure), get dose profiles command requires dicom co-ords.. Can use DICOMToUser(), command to convert to user co-ords; ****/
                /*** Which is what Eclipse will display on the screen. But as we are pulling the co-ords in dicom already - direct from the ref. pt. we can skip that step. ****/

                /*** Get Profile using start and stop points and double array ****/
                proff = planSelec.Dose.GetDoseProfile(startCrossline, stopCrossline, preallocatedBuffer: usedis);

                foreach (var vv in proff)
                {
                    Console.WriteLine($"{(proff[contta].Value)}"); //  /normaval
                    if (proff[contta].Value > maxyvv)
                    {
                        mvv = proff[contta].Value;
                        maxyvv = mvv;

                    }
                    contta++;
                }

                Console.WriteLine($" ");
                Console.WriteLine($"The units are in: {proff.Unit}");

                /*** Normalisation ****/
                /*** create string array to select normalisation value ****/
                /*
                //[0] - None (units will maybe be in cGy? or in %, depending on plan settings? not sure). 
                //[1] - Maximum value anywhere along the extraced PDD/Profile
                //[3] - Value at the selected reference point
                //[4] - Custom value - will be next prompted to input a value to normalise to. */

                Console.WriteLine($"What normalisation value would you like to use?");
                string[] cwla = { "None", "Max value along profile", "Value at the reference point", "Custom Value - if selected will be prompted to input." };
                var cwlaList = new List<string>(cwla);
                pdproSa = GetUserListSelection(cwlaList);

                /*** Display selected value ****/
                var normavalname = cwla[pdproS]; //Option selected
                Console.WriteLine($"You have selected, {pdproSa}. - {normavalname}");

                /*** Set the normal values  ****/
                normaval = 1;
                normavalINPLANE = 1;

                /*** If custom normalisation selected ****/
                if (pdproSa == 3)
                {
                    /*** Have forced displayed units to whatever the system has as the absolute units. For our centre this is Gy (but others may have cGy).  ****/
                    Console.WriteLine($"You have selected custom normalisation. Input the normalisation value in units of UNITs: Gy."); 
                    double normya = -1; //Do this, in case user inputs a letter or something.
                    if (!double.TryParse(Console.ReadLine(), out normya))
                    {
                        /*** Ensure selection is numeric otherwise error ****/
                        Console.WriteLine("Invalid Selection"); return; 
                    }
                    /*** Set normal values to be custom value ****/
                    normaval = normya; //Custom normalisation value in Gy.
                    normavalINPLANE = normaval;
                }
                /*** if no normalisation selected ****/
                else if (pdproSa == 0)
                {
                    Console.WriteLine($"You have selected No normalisation. The profile units will remain in Gray.");
                    /*** Set normal values to 100 ****/
                    normaval = 100;
                    normavalINPLANE = normaval;
                }
                /*** if reference point selected ****/
                else if (pdproSa == 2)
                {
                    decimal dd = (proff.Count / 2);
                    int ah1 = (int)Math.Round(dd, 0);


                    Console.WriteLine($"You have selected to normalise to the value at the reference point. Normalisation values will be:");
                    Console.WriteLine($" {proff[ah1].Value} , units of: {proff.Unit}"); //proff[contta].Value

                    /*** set mornal value as value at ref point ****/
                    normaval = (double)proff[ah1].Value;
                    normavalINPLANE = normaval;
                }
                /*** if max value selected ****/
                else if (pdproSa == 1)
                {
                    Console.WriteLine($"You have selected to normalise to the max value along the profile. Normalisation values will be:");
                    Console.WriteLine($"Crossplane {mvv} ,units of: {proff.Unit}");
                    Console.WriteLine($"Inplane {mvI} ,units of: {proffI.Unit}");
                    //Console.WriteLine($"Crossplane{PDDprof.Max()} ,units of:{PDDprof.Unit}");
                    normaval = (double)mvv;
                    normavalINPLANE = (double)mvI;
                }

                /*** apply the normalisation to cross plane values and re-express the values as a percentage: ****/
                Console.WriteLine($"Crossplane profiles (assuming HFS setup) are (normalised) - expressed as a %:");
                contta = 0;
                foreach (var v in proff)
                {
                    Console.WriteLine($"{((proff[contta].Value * 100) / normaval)}"); //  /normaval
                    contta++;
                }
                /*** apply the normalisation to inplane values and re-express the values as a percentage: ****/
                Console.WriteLine($"Inplane profiles are (normalised) - expressed as a %:");
                conttta = 0;
                foreach (var vv in proffI)
                {
                    Console.WriteLine($"{((proffI[conttta].Value * 100) / normavalINPLANE)}");
                    conttta++;
                }

                /*** Ask user to select output to sundose compatible format ****/
                Console.WriteLine($" ");
                Console.WriteLine($"Would you like to export these profiles to a SunDose compatible format?");
                Console.WriteLine($" ");
                string[] exporta = { "Yes", "No", "Bonus option: export profiles FOR ALL reference points in the plan - with identical settings." };
                var exportaList = new List<string>(exporta);
                int qqq = GetUserListSelection(exportaList);

                /*** Get the user decision to export to SunDose format ****/
                var decision = exporta[qqq]; 
                Console.WriteLine($"You have selected, {decision}. - {exporta[qqq]}");

                /*** Get beam information from plan ****/
                type = string.Empty;
                tmtmach = string.Empty;
                modality = string.Empty;
                rtype = string.Empty;
                ssd = string.Empty;
                xfs = string.Empty;
                yfs = string.Empty;
                isElectron = false;
                xfs_mlc = string.Empty;
                yfs_mlc = string.Empty;


                Beam beam = planSelec.Beams.Last(x => !x.IsSetupField);
                
                if (beam != null)
                {
                    isElectron = beam.EnergyModeDisplayName.Contains("E");
                    tf_pdi = isElectron ? "True" : "False";
                    tmtmach = beam.TreatmentUnit.Id;
                    NilMLC = beam.MLCPlanType.ToString().Contains("NotDefined"); //If MLC is undefined, then use Jaws to pull field-size.
                    ssd = (beam.SSD/10).ToString(); //SSD in cm.
                    //xfs = Math.Abs((beam.ControlPoints.First().JawPositions.X1*2/10)).ToString(); //X jaws, needs to be in units of cm, and for whole field. (is assumed symmetric).
                    //yfs = Math.Abs(beam.ControlPoints.First().JawPositions.Y1*2/10).ToString(); //Needs to be in units of cm, and for whole field. (is assumed symmetric).
                    type = isElectron ? beam.EnergyModeDisplayName.Replace("E", "") : beam.EnergyModeDisplayName.Replace("X", "").Replace("-FFF","").Replace("FFF", ""); //Added .Replace("-FFF","")
                    rtype = isElectron ? "Scattering Foil" : beam.EnergyModeDisplayName.Contains("FFF") ? "FFF" : "FF";
                    modality = isElectron ? "Electron" : "Photon";
                    xfs = (Math.Abs(beam.ControlPoints.First().JawPositions.X1)  / 10).ToString();
                    yfs = (Math.Abs(beam.ControlPoints.First().JawPositions.Y1)  / 10).ToString();
                    xfs_mlc = NilMLC ? xfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30])  / 10).ToString(); // If no MLC, pull jaws, convert to cm for entire field. Assumes centred and symmetric jaws. If MLCs, assume these define field, pull from one bank, should be ~central MLC (e.g. for 28 leaf pairs for OHAL1 is ~14; for 60 leaf pairs for OTB1, is not as robust). User can edit if programmitc approach gets it wrong. Once again assumes centred and symmetric. Script usually run for simple static centred open fields.
                    yfs_mlc = NilMLC ? yfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30]) / 10).ToString();
                    sumx_mlc = NilMLC ? xfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30]) *2 / 10).ToString();
                    sumy_mlc = NilMLC ? yfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30]) * 2 / 10).ToString();
                    mlc2 = yfs_mlc; //NilMLC ? xfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30]) * 2 / 10).ToString(); //Broken code to fix SNX bug.
                    if (isElectron) {
                        xfs = beam.Applicator.Id.ToString().Replace("A", "");
                        yfs = beam.Applicator.Id.ToString().Replace("A", "");
                    }
                }



                /*** Export for SunDose ****/
                if (qqq == 0)
                {
                    Console.WriteLine($"...");
                    /*** Set Filename for inplane/crossplane ****/
                    if (pdproS == 2)
                    {
                        string fileName = $"{planSelec}_{refSelec}_TPSexport_Crossline.snctxt"; //{cwl.ElementAt(pdproS)}
                        fileName = SanitizeFileName(fileName);
                        /*** Set Folder name ****/
                        string fullPath = destinationFolder + fileName;
                        /*** Create directory to save file ****/
                        if (!Directory.Exists(destinationFolder)) Directory.CreateDirectory(destinationFolder);

                        /*** Write data to file ****/
                        StreamWriter sw = new StreamWriter(fullPath);
                        sw.WriteLine("Tab-Delimited Scan Output");
                        sw.WriteLine("FILE HEADER");
                        sw.WriteLine($"File Name\t{fileName}");
                        sw.WriteLine("File Date\t02/20/2077 00:01");
                        sw.WriteLine($"File Export Version\t4.0.1.8\t\r\nFile Version\t5.0\t\r\nFile Scan Count\t1\t\r\nBEGIN SCAN\r\nSummary Comments\t\t\r\nSummary Beam Type\t{modality}\t\r\nBeam Type\t{modality}\t\r\nEnergy (MV / MeV)\t{type}\t\r\nRate Type\t{rtype}\t\r\nSummary Energy (MV/MeV)\t{type}\t\r\nSummary FieldSize X (cm)\t{sumx_mlc}\t\r\nSummary FieldSize Y (cm)\t{sumy_mlc}\t\r\nSummary Wedge Type\tOpen Field\t\r\nSummary Wedge Angle (degrees)\t0.00\t\r\nSummary Scan Type\t{cwl.ElementAt(pdproS)}");
                        sw.WriteLine($"Is PDI to PDD\t{tf_pdi}\t\r\n\r\nBEGIN DOSE TABLE\t\t\r\nAction\tSmooth\t\r\n\tX (cm)\tY (cm)\tZ (cm)\tRelative Dose (%)");

                        /*** Enumerate over dose points here ****/
                        /*** NOTE MUST CONVERT TO WATER-TANK CO-ORDINATES. //So going from DICOM to SUNDOSE WT ****/
                        int co = 0;
                        try
                        {
                            VVector startCrosslineexp = new VVector((reffy.x - symdist), reffy.y, reffy.z); //required to redefine doseprofile prior to export.
                            VVector stopCrosslineexp = new VVector((reffy.x + symdist), reffy.y, reffy.z);
                            usedis = new double[(int)Math.Ceiling((startCrosslineexp - stopCrosslineexp).Length / stepsize)];
                            proff = planSelec.Dose.GetDoseProfile(startCrosslineexp, stopCrosslineexp, preallocatedBuffer: usedis);


                            foreach (var qwerty in proff)
                            {
                                /*** Water tank co-ords are x = crossline, y = inline, z = depth, units are 'cm' so must also convert from mm. ****/
                                /*** Whereas DICOM co-ordinates (in HFS) are x = crossline, y = depth, z = inline. ****/

                                bool res = Double.IsNaN(proff[co].Value); //If it's NaN give a true value. Will invert it in the next line to determine if we should save the value to file.
                                if (!res)
                                {
                                    sw.WriteLine($"\t{Math.Round(proff[co].Position.x / 10, 4)}\t{Math.Round(proff[co].Position.z / 10, 4)}\t{Math.Round(reffy.y, 4)}\t{Math.Round(proff[co].Value * 100 / normaval, 4)}");
                                }
                                co++;
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("No Crosslines written due to error. Most likely none were created. Full error message; The process failed: {0}", e.ToString());
                            throw;
                        }

                        /*** Add raw data value to footer ****/
                        sw.WriteLine($"\r\n\r\nEND DOSE TABLE\r\nSCAN HEADER\r\nFACILITY INFORMATION\r\nInstitution\tCWCCC\t\r\nDelivery System\t{tmtmach}\t\r\nDelivery System Manufacturer\t\t\r\nDelivery System Model #\t\t\r\nDelivery System Serial #\tH195586\t\r\nField Detector Model #\tCC13 #96910 Orange\t\r\nField Detector Serial #\t96910\t\r\nReference Detector Model #\tCC13 #96830 Orange\t\r\nReference Detector Serial #\t96830\t\r\nSNC EQUIPMENT\r\nApplication Programming Interface\t4.0.1.8\t\r\nHardware Device Interface\t4.0.0.354\t\r\nSunSCAN 3D Model #\tNot Available\t\r\nSunSCAN 3D Serial #\t64368013\t\r\nSunSCAN 3D Firmware\t1.0.4.2\t\r\nDrive Factor Diameter (pulse/mm)\t0\t\r\nDrive Factor Vertical (pulse/mm)\t0\t\r\nDrive Factor Ring (pulse/mm)\t0\t\r\nElectrometer Model #\tNot Available\t\r\nElectrometer Serial #\t64368013\t\r\nElectrometer Firmware\t1.0.1.3\t\r\nLeveling Platform Model #\tNot Available\t\r\nLeveling Platform Serial #\tNot Available\t\r\nLift Table Model #\tNot Available\t\r\nLift Table Serial #\tNot Available\t\r\nReservoir Model #\tNot Available\t\r\nReservoir Serial #\tNot Available\t\r\nBootloader Firmware\tNot Available\t\r\nError Status Register\tNot Available\t\r\nFormatted Application Version\tNot Available\t\r\nFormatted Bootblock Version\tNot Available\t\r\nFormatted PMD OS Firmware Version\tNot Available\t\r\nHardware Version\tNot Available\t\r\nDELIVERY SYSTEM\r\nGantry Angle (degrees)\t0\t\r\nCollimator Angle (degrees)\t0.00\t\r\nCollimation Type\tJaws and MLC\t\r\nWedge Type\tOpen Field\t\r\nWedge Angle (degrees)\t0.00\t\r\nWedge Direction\t\t\r\nField Size X (cm)\t{sumx_mlc}\t\r\nField Size Y (cm)\t{sumy_mlc}\t\r\nField Shape\tSquare\t\r\nMeasurement Unit\tcm\t\r\nCollimator Position Jaws X1 (cm) {xfs}\t\t\r\nCollimator Position Jaws X2 (cm) {xfs}\t\t\r\nCollimator Position Jaws Y1 (cm) {yfs}\t\t\r\nCollimator Position Jaws Y2 (cm) {yfs}\t\t\r\nCollimator Position MLC X1 (cm) {xfs_mlc}\t\t\r\nCollimator Position MLC X2 (cm) {mlc2}\t\t\r\nCollimator Position MLC Y1 (cm) {yfs_mlc}\t\t\r\nCollimator Position MLC Y2 (cm) {mlc2}\t\t\r\nSETUP PARAMETERS\r\nRing Center (cm)\t25.218\t\r\nAngle Offset (degrees)\t4.53\t\r\nHysteresis Minus (cm)\t0\t\r\nMEASUREMENT DETAILS\r\nComments\t\t\r\nScan Id\t3384\t\r\nScan Date\t02/17/2024 03:57\t\r\nScan Type\tCrossline\t\r\nScan Medium\tWater\t\r\nSource to Surface Distance (cm)\t{ssd}\t\r\nIon Chamber Equivalent Model\t\t\r\nScan Source\tSunSCAN 3D\t\r\nSunSCAN\tTrue\t\r\nMeasurement Mode\tContinuous\t\r\nScan Speed (cm/s)\t0.05 cm/second\t\r\nStyle\t7\t\r\nEMF Spacing (cm)\t0.050\t\r\nOptimized Rotation\tTrue\t\r\nDiameter Drive Scan Direction\tFalse\t\r\nAdditional Scan Range (cm)\t5.00\t\r\nIntegrated Measurement\tFalse\t\r\nEffective Point of Measurement (cm)\t0.15\t\r\nDetector Bias Voltage (V)\t304.34\t\r\nReference Detector Bias Voltage (V)\t302.80\t\r\nField Background Rate (counts/update)\t0.005\t\r\nReference Background Rate (counts/update)\t0.00207035175879397\t\r\nNormalization Value (Field/Reference)\t0.802148539760211\t\r\nPulse Normalized\tFalse\t\r\nMeasurement Current\tDynamic\t\r\nOverscan Amount\t\t\r\nOffset Detector Holder\tFalse\t\r\n\r\n\r\nBEGIN RAW DATA\t\t\t\t\r\n\tSequence\tTheta\tDiameter\tX (cm)\tY (cm)\tZ (cm)\tDelta Time (s)\tReference Cumulative Counts\tField Cumulative Counts\tCorrected Ratio\tCumulative Pulses\tVoltage (V)\tElectrometer Timer (s)\tMotion Timer (s)\tPressure\tInternal Temperature\tExternal Temperature\tPlus 5 Sensor\tReference: Pulse Size (counts)\tReference: Timed Measurement\tReference: Cumulative Buckets\tReference: Cumulative Counts (counts)\tReference: Cumulative Counts Corrected (counts)\tReference: Number of Updates when Measuring Charge\tReference: Number of Updates when Measuring Bucket Charge\tReference: Rail\tField: Pulse Size (counts)\tField: Timed Measurement\tField: Cumulative Buckets\tField: Cumulative Counts (counts)\tField: Cumulative Counts Corrected (counts)\tField: Number of Updates when Measuring Charge\tField: Number of Updates when Measuring Bucket Charge\tField: Rail\tTPR Sensor Calibrated\tTPR Sensor Raw\tTPR Buildup\tReference Voltage (V)\tActual X (cm)\tActual Y (cm)\tActual Z (cm)\tApplied Pulse Count\tRails Hit\r\n\t1\t0\t18.051\t18.051\t0\t2.8\t0.1\t81935\t467\t0.00563875845749496\t12\t304.2573\t0.1\t0.0989\t0\t0\t-273.15\t0\t81961\t0\t0\t81935\t81935\t1000\t0\t0\t497\t0\t0\t467\t467\t1000\t0\t0\t0\t0\t5.394\t303.0365\t18.051\t0\t2.8\t0\tFalse\r\n\t2\t0\t18.051\t18.051\t0\t2.8\t0.05\t116054\t655\t0.00543701832617846\t17\t304.3488\t0.15\t0.145\t0\t0\t-273.15\t0\t116102\t0\t0\t116054\t116054\t1500\t0\t0\t691\t0\t0\t655\t655\t1500\t0\t0\t0\t0\t5.394\t302.8992\t18.051\t0\t2.8\t0\tFalse");
                        sw.WriteLine("END RAW DATA\r\nEND SCAN\t");
                        sw.WriteLine($"Created By: {_currentUserId}");
                        sw.Close();


                        Console.WriteLine($"Success, profiles/etc.. exported to: {fullPath}");
                    }

                    if (pdproS == 3)
                    {
                        /*** Set inline filename ****/
                        string fileName2 = $"{planSelec}_{refSelec}_TPSexport_Inline.snctxt";
                        fileName2 = SanitizeFileName(fileName2);

                        /*** Set Folder name ****/
                        string fullPath2 = destinationFolder + fileName2;
                    /*** Create directory to save file ****/
                    if (!Directory.Exists(destinationFolder)) Directory.CreateDirectory(destinationFolder);

                    
                        /*** Write data to file ****/
                        StreamWriter sw2 = new StreamWriter(fullPath2);
                        sw2.WriteLine("Tab-Delimited Scan Output");
                        sw2.WriteLine("FILE HEADER");
                        sw2.WriteLine($"File Name\t{fileName2}");
                        sw2.WriteLine("File Date\t02/20/2077 00:02");
                        sw2.WriteLine($"File Export Version\t4.0.1.8\t\r\nFile Version\t5.0\t\r\nFile Scan Count\t1\t\r\nBEGIN SCAN\r\nSummary Comments\t\t\r\nSummary Beam Type\t{modality}\t\r\nBeam Type\t{modality}\t\r\nEnergy (MV / MeV)\t{type}\t\r\nRate Type\t{rtype}\t\r\nSummary Energy (MV/MeV)\t{type}\t\r\nSummary FieldSize X (cm)\t{xfs}\t\r\nSummary FieldSize Y (cm)\t{yfs}\t\r\nSummary Wedge Type\tOpen Field\t\r\nSummary Wedge Angle (degrees)\t0.00\t\r\nSummary Scan Type\tInline");
                        sw2.WriteLine($"Is PDI to PDD\t{tf_pdi}\t\r\n\r\nBEGIN DOSE TABLE\t\t\r\nAction\tSmooth\t\r\n\tX (cm)\tY (cm)\tZ (cm)\tRelative Dose (%)");

                        /*** Enumerate over points in plane ****/
                        /*** NOTE MUST CONVERT TO WATER-TANK CO-ORDINATES. //So going from DICOM to SUNDOSE WT. ****/
                        int co = 0;
                        try
                        {
                            VVector startInlineexp = new VVector(reffy.x, reffy.y, (reffy.z + symdist)); //redefine prior to export. Unsure why needed to do these but fixes the bug -JH.
                            VVector stopInlineexp = new VVector(reffy.x, reffy.y, (reffy.z - symdist));
                            double[] usedis2 = new double[(int)Math.Ceiling((startInlineexp - stopInlineexp).Length / stepsize)];
                            /*** Get Profile using start and stop points and double array ****/
                            proffI = planSelec.Dose.GetDoseProfile(startInlineexp, stopInlineexp, preallocatedBuffer: usedis2);

                            /*** Water tank co-ords are x = crossline, y = inline, z = depth, units are 'cm' so must also convert from mm. ****/
                            /*** whereas DICOM co-ordinates (in HFS) are x = crossline, y = depth, z = inline. ****/
                            foreach (var qwerty in proffI) //enumerate over -plane.//NOTE MUST CONVERT TO WATER-TANK CO-ORDINATES. //So going from DICOM to SUNDOSE WT.
                            {
                                bool res = Double.IsNaN(proffI[co].Value); //If it's NaN give a true value. Will invert it in the next line to determine if we should save the value to file.
                                if (!res)
                                {
                                    sw2.WriteLine($"\t{Math.Round(proffI[co].Position.x / 10, 4)}\t{Math.Round(proffI[co].Position.z / 10, 4)}\t{Math.Round(reffy.y, 1)}\t{Math.Round(proffI[co].Value * 100 / normavalINPLANE, 4)}"); //
                                }
                                co++;
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("No Inlines written due to error. Most likely none were created. Full error message; The process failed: {0}", e.ToString());
                            throw;
                        }

                        /*** Add raw data to the footer ****/
                        sw2.WriteLine($"\r\n\r\nEND DOSE TABLE\r\nSCAN HEADER\r\nFACILITY INFORMATION\r\nInstitution\tCWCCC\t\r\nDelivery System\t{tmtmach}\t\r\nDelivery System Manufacturer\t\t\r\nDelivery System Model #\t\t\r\nDelivery System Serial #\tH195586\t\r\nField Detector Model #\tCC13 #96910 Orange\t\r\nField Detector Serial #\t96910\t\r\nReference Detector Model #\tCC13 #96830 Orange\t\r\nReference Detector Serial #\t96830\t\r\nSNC EQUIPMENT\r\nApplication Programming Interface\t4.0.1.8\t\r\nHardware Device Interface\t4.0.0.354\t\r\nSunSCAN 3D Model #\tNot Available\t\r\nSunSCAN 3D Serial #\t64368013\t\r\nSunSCAN 3D Firmware\t1.0.4.2\t\r\nDrive Factor Diameter (pulse/mm)\t0\t\r\nDrive Factor Vertical (pulse/mm)\t0\t\r\nDrive Factor Ring (pulse/mm)\t0\t\r\nElectrometer Model #\tNot Available\t\r\nElectrometer Serial #\t64368013\t\r\nElectrometer Firmware\t1.0.1.3\t\r\nLeveling Platform Model #\tNot Available\t\r\nLeveling Platform Serial #\tNot Available\t\r\nLift Table Model #\tNot Available\t\r\nLift Table Serial #\tNot Available\t\r\nReservoir Model #\tNot Available\t\r\nReservoir Serial #\tNot Available\t\r\nBootloader Firmware\tNot Available\t\r\nError Status Register\tNot Available\t\r\nFormatted Application Version\tNot Available\t\r\nFormatted Bootblock Version\tNot Available\t\r\nFormatted PMD OS Firmware Version\tNot Available\t\r\nHardware Version\tNot Available\t\r\nDELIVERY SYSTEM\r\nGantry Angle (degrees)\t0\t\r\nCollimator Angle (degrees)\t0.00\t\r\nCollimation Type\tJaws and MLC\t\r\nWedge Type\tOpen Field\t\r\nWedge Angle (degrees)\t0.00\t\r\nWedge Direction\t\t\r\nField Size X (cm)\t{xfs}\t\r\nField Size Y (cm)\t{yfs}\t\r\nField Shape\tSquare\t\r\nMeasurement Unit\tcm\t\r\nCollimator Position Jaws X1 (cm) {xfs}\t\t\r\nCollimator Position Jaws X2 (cm) {xfs}\t\t\r\nCollimator Position Jaws Y1 (cm) {yfs}\t\t\r\nCollimator Position Jaws Y2 (cm) {yfs}\t\t\r\nCollimator Position MLC X1 (cm) {xfs_mlc}\t\t\r\nCollimator Position MLC X2 (cm) {mlc2}\t\t\r\nCollimator Position MLC Y1 (cm) {yfs_mlc}\t\t\r\nCollimator Position MLC Y2 (cm) {mlc2}\t\t\r\nSETUP PARAMETERS\r\nRing Center (cm)\t0\t\r\nAngle Offset (degrees)\t0\t\r\nHysteresis Minus (cm)\t0\t\r\nMEASUREMENT DETAILS\r\nComments\t\t\r\nScan Id\t3384\t\r\nScan Date\t02/17/2024 03:57\t\r\nScan Type\tInline\t\r\nScan Medium\tWater\t\r\nSource to Surface Distance (cm)\t{ssd}\t\r\nIon Chamber Equivalent Model\t\t\r\nScan Source\tSunSCAN 3D\t\r\nSunSCAN\tTrue\t\r\nMeasurement Mode\tContinuous\t\r\nScan Speed (cm/s)\t0.05 cm/second\t\r\nStyle\t7\t\r\nEMF Spacing (cm)\t0.050\t\r\nOptimized Rotation\tTrue\t\r\nDiameter Drive Scan Direction\tFalse\t\r\nAdditional Scan Range (cm)\t5.00\t\r\nIntegrated Measurement\tFalse\t\r\nEffective Point of Measurement (cm)\t0.15\t\r\nDetector Bias Voltage (V)\t304.34\t\r\nReference Detector Bias Voltage (V)\t302.80\t\r\nField Background Rate (counts/update)\t0.005\t\r\nReference Background Rate (counts/update)\t0.00207035175879397\t\r\nNormalization Value (Field/Reference)\t0.802148539760211\t\r\nPulse Normalized\tFalse\t\r\nMeasurement Current\tDynamic\t\r\nOverscan Amount\t\t\r\nOffset Detector Holder\tFalse\t\r\n\r\n\r\nBEGIN RAW DATA\t\t\t\t\r\n\tSequence\tTheta\tDiameter\tX (cm)\tY (cm)\tZ (cm)\tDelta Time (s)\tReference Cumulative Counts\tField Cumulative Counts\tCorrected Ratio\tCumulative Pulses\tVoltage (V)\tElectrometer Timer (s)\tMotion Timer (s)\tPressure\tInternal Temperature\tExternal Temperature\tPlus 5 Sensor\tReference: Pulse Size (counts)\tReference: Timed Measurement\tReference: Cumulative Buckets\tReference: Cumulative Counts (counts)\tReference: Cumulative Counts Corrected (counts)\tReference: Number of Updates when Measuring Charge\tReference: Number of Updates when Measuring Bucket Charge\tReference: Rail\tField: Pulse Size (counts)\tField: Timed Measurement\tField: Cumulative Buckets\tField: Cumulative Counts (counts)\tField: Cumulative Counts Corrected (counts)\tField: Number of Updates when Measuring Charge\tField: Number of Updates when Measuring Bucket Charge\tField: Rail\tTPR Sensor Calibrated\tTPR Sensor Raw\tTPR Buildup\tReference Voltage (V)\tActual X (cm)\tActual Y (cm)\tActual Z (cm)\tApplied Pulse Count\tRails Hit\r\n\t1\t0\t18.051\t18.051\t0\t2.8\t0.1\t81935\t467\t0.00563875845749496\t12\t304.2573\t0.1\t0.0989\t0\t0\t-273.15\t0\t81961\t0\t0\t81935\t81935\t1000\t0\t0\t497\t0\t0\t467\t467\t1000\t0\t0\t0\t0\t5.394\t303.0365\t18.051\t0\t2.8\t0\tFalse\r\n\t2\t0\t18.051\t18.051\t0\t2.8\t0.05\t116054\t655\t0.00543701832617846\t17\t304.3488\t0.15\t0.145\t0\t0\t-273.15\t0\t116102\t0\t0\t116054\t116054\t1500\t0\t0\t691\t0\t0\t655\t655\t1500\t0\t0\t0\t0\t5.394\t302.8992\t18.051\t0\t2.8\t0\tFalse");
                        sw2.WriteLine("END RAW DATA\r\nEND SCAN\t");
                        sw2.WriteLine($"Created By: {_currentUserId}");
                        sw2.Close();
                    }
                }               
                /*** Get the selection from user and act accordingly ****/
               else if (qqq == 2)
                {
                    ExportAllProfileWithSameSettings();
                }
                else if (qqq != 0)
                {
                    /*** No export required ****/
                    Console.WriteLine($"Ok, have nice day!");
                }
            }
        }

        static void ProcessPDD()
        {
            /*** get reference point for PDD ****/
            Console.WriteLine($"Please input the TOTAL depth from the reference point in millimeters for the PDD (e.g. 300.");
            int dist_pdd = -1; 
            if (!int.TryParse(Console.ReadLine(), out dist_pdd))
            {
                /*** Ensure selection is numeric otherwise error ****/
                Console.WriteLine("Invalid Selection"); return; //If they didn't put in a numeric. Gives errar.
            }

            /*** Get user to set step size ****/
            Console.WriteLine($"Please input the PDD step-size in 'mm', whole integers only, e.g. 1 mm, 2 mm.");
            int dista_pdd = -1; 
            if (!int.TryParse(Console.ReadLine(), out dista_pdd))
            {
                /*** Ensure selection is numeric otherwise error ****/
                Console.WriteLine("Invalid Selection"); return; 
            }
            var stepsize_pdd = dista_pdd;

            /*** Add the vectors (writing the start/stop points in vector format) to get start/stop points from the ref point vector). ****/
            /*** ASSUMES HFS and reffy is in DICOM co-ordinates ****/
            VVector startPDD = new VVector(reffy.x, reffy.y, reffy.z);
            VVector stopPDD = new VVector(reffy.x, reffy.y + dist_pdd, reffy.z);

            /*** set double array from start and stop collection points ****/
            double[] usedis_pdd = new double[(int)Math.Ceiling((stopPDD - startPDD).Length / stepsize_pdd)];

            /*** Co-ordinates are in DICOM (pretty sure), get dose profiles command requires dicom co-ords.. Can use DICOMToUser(), command to convert to user co-ords; ****/
            /*** Which is what Eclipse will display on the screen. But as we are pulling the co-ords in dicom already - direct from the ref. pt. we can skip that step. ****/

            /*** Get the profile based on start, stop and double array of points ****/
            var pdd_d = planSelec.Dose.GetDoseProfile(startPDD, stopPDD, preallocatedBuffer: usedis_pdd);
            Console.WriteLine($"(Un-normalised DepthDose), from reference point down (assuming HFS setup):");

            /*** Display all points values for profile found ****/
            int conttta_pdd = 0;
            double maxyv_pdd = 0;
            double mvI_pdd = 0;
            foreach (var v_pdd in pdd_d)
            {
                Console.WriteLine($"{pdd_d[conttta_pdd].Value}");
                if ((double)pdd_d[conttta_pdd].Value > maxyv_pdd)
                {
                    mvI_pdd = pdd_d[conttta_pdd].Value;
                    maxyv_pdd = mvI_pdd;

                }
                conttta_pdd++;
            }

            
            Console.WriteLine($" ");
            Console.WriteLine($" ");

            /*** Create user list for normalisation ****/
            Console.WriteLine($"What normalisation value would you like to use?");
            string[] cwla_pdd = { "None", "Max value along PDD", "Value at the reference point", "Custom Value - if selected will be prompted to input." };
            var cwla_pddList = new List<string>(cwla_pdd);
            int pdproSa_pdd = GetUserListSelection(cwla_pddList);

            /*** Get option selected ****/
            var normavalname_pdd = cwla_pdd[pdproSa_pdd]; 
            Console.WriteLine($"You have selected, {pdproSa_pdd}. - {normavalname_pdd}");

            /*** Set default normalisation value ****/
            double normaval_pdd = 1;

            /*** Id custom value chosen ****/
            if (pdproSa_pdd == 3)
            {
                /*** Have forced displayed units to whatever the system has as the absolute units. For our centre this is Gy (but others may have cGy). ****/
                Console.WriteLine($"You have selected custom normalisation. Input the normalisation value in units of UNITs: Gy."); 
                int normya_pdd = -1; 
                if (!int.TryParse(Console.ReadLine(), out normya_pdd))
                {
                    /*** Ensure selection is numeric otherwise error ****/
                    Console.WriteLine("Invalid Selection"); return;

                }
                /*** Custom normalisation value in Gy. ****/
                normaval_pdd = normya_pdd; 
                                            
            }
            /*** if no normalisation selected ****/
            else if (pdproSa_pdd == 0)
            {
                /*** Set the mornalisation value to 100 ****/
                Console.WriteLine($"You have selected No normalisation. The profile units will remain in Gray.");
                normaval_pdd = 100;
            }
            /*** if refernce point chosen ****/
            else if (pdproSa_pdd == 2)
            {
               // decimal dd_pdd = (pdd_d.Count());
                //int ah1_pdd = (int)Math.Round(dd_pdd, 0);


                Console.WriteLine($"You have selected to normalise to the value at the reference point. Normalisation values will be:");
                Console.WriteLine($" {pdd_d[0].Value} , units of: {pdd_d.Unit}"); //proff[contta].Value

                /*** set the normalisation value to be that at reference point ****/
                normaval_pdd = (double)pdd_d[0].Value;
            }
            /*** if max point chosen ****/
            else if (pdproSa_pdd == 1)
            {
                Console.WriteLine($"You have selected to normalise to the max value along the profile. Normalisation values will be:");

                Console.WriteLine($"PDD{maxyv_pdd} ,units of:{pdd_d.Unit}");
                /*** set normalistion value to max point ****/
                normaval_pdd = (double)mvI_pdd;
                // normavalINPLANE = (double)mvI_pdd;
            }

            /*** apply the normalisation and re-express the values as a percentage: ****/
            Console.WriteLine($"PDD (assuming HFS setup) and (normalised) - expressed as a %:");
            int contta_pdd = 0;
            foreach (var v_pdd in pdd_d)
            {
               Console.WriteLine($"{((pdd_d[contta_pdd].Value * 100) / normaval_pdd)}"); //Double.IsNaN(v_pdd) ?  : 
                contta_pdd++;
            }

            /*** User to select whether to export as SunDose format ****/
            Console.WriteLine($" ");
            Console.WriteLine($"Would you like to export these profiles to a SunDose compatible format?");
            Console.WriteLine($" ");
            string[] exporta_pdd = { "Yes", "No" };
            var exporta_pddList = new List<string>(exporta_pdd);
            int qqq_pdd = GetUserListSelection(exporta_pddList);

            /*** Get user Selection ****/
            var decision_pdd = exporta_pdd[qqq_pdd]; 
            Console.WriteLine($"You have selected, {decision_pdd}. - {exporta_pdd[qqq_pdd]}");

            /*** Get beam information from plan ****/
            type = string.Empty;
            tmtmach = string.Empty;
            modality = string.Empty;
            rtype = string.Empty;
            isElectron = false;
            NilMLC = false;

            Beam beam = planSelec.Beams.Last(x => !x.IsSetupField);

            if (beam != null)
            {
                isElectron = beam.EnergyModeDisplayName.Contains("E");
                tf_pdi = isElectron ? "True" : "False";
                NilMLC = beam.MLCPlanType.ToString().Contains("NotDefined"); //If MLC is undefined, then use Jaws to pull field-size. N.B. OHAL1 is currently undefined even if MLCs are used. Works for OTB1 though.
                ssd = (beam.SSD / 10).ToString(); //SSD in cm.
                tmtmach = beam.TreatmentUnit.Id;
                type = isElectron ? beam.EnergyModeDisplayName.Replace("E", "") : beam.EnergyModeDisplayName.Replace("X", "").Replace("-FFF", "").Replace("FFF", "");
                rtype = isElectron ? "Scattering Foil" : beam.EnergyModeDisplayName.Contains("FFF") ? "FFF" : "FF";
                modality = isElectron ? "Electron" : "Photon";
                xfs = (Math.Abs(beam.ControlPoints.First().JawPositions.X1)  / 10).ToString();
                yfs = (Math.Abs(beam.ControlPoints.First().JawPositions.Y1)  / 10).ToString();
                xfs_mlc = NilMLC ? xfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30])/10).ToString(); // If no MLC, pull jaws, convert to cm for entire field. Assumes centred and symmetric jaws. If MLCs, assume these define field, pull from one bank, should be ~central MLC (e.g. for 28 leaf pairs for OHAL1 is ~14; for 60 leaf pairs for OTB1, is not as robust). User can edit if programmitc approach gets it wrong. Once again assumes centred and symmetric. Script usually run for simple static centred open fields.
                yfs_mlc = NilMLC ? yfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30])/10).ToString();
                sumx_mlc = NilMLC ? xfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30]) * 2 / 10).ToString();
                sumy_mlc = NilMLC ? yfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30]) * 2 / 10).ToString();
                mlc2 = yfs_mlc; //NilMLC ? xfs : Math.Abs((beam.ControlPoints.First().LeafPositions[0, 30]) * 2 / 10).ToString(); //Broken code to fix SNX bug.
                if (isElectron)
                {
                    xfs = beam.Applicator.Id.ToString().Replace("A", "");
                    yfs = beam.Applicator.Id.ToString().Replace("A", "");
                }

            }

            /*** Now to write whatever was created into the required formatting for SunDose. ****/
            if (qqq_pdd == 0)
            {
                Console.WriteLine($"...");
                /*** Set PDD File name format ****/
                string fileName = $"{planSelec}_{refSelec}_TPSexport_PDD.snctxt";
                fileName = SanitizeFileName(fileName);
                /*** Set Folder name ****/
                string fullPath = destinationFolder + fileName; 
                Console.WriteLine("Attempting to write to:");
                
                Console.WriteLine (fullPath );
                Console.WriteLine("fileName is:");

                Console.WriteLine(fileName);
                /*** Create directory to save file ****/
                if (!Directory.Exists(destinationFolder)) Directory.CreateDirectory(destinationFolder);

                /*** Write data to file ****/
                
                StreamWriter sw = new StreamWriter(fullPath);
                sw.WriteLine("Tab-Delimited Scan Output");
                sw.WriteLine("FILE HEADER");
                sw.WriteLine($"File Name\t{fileName}");
                sw.WriteLine("File Date\t02/20/2077 00:02");
                sw.WriteLine($"File Export Version\t4.0.1.8\t\r\nFile Version\t5.0\t\r\nFile Scan Count\t1\t\r\nBEGIN SCAN\r\nSummary Comments\t\t\r\nSummary Beam Type\t{modality}\t\r\nBeam Type\t{modality}\t\r\nEnergy (MV / MeV)\t{type}\t\r\nRate Type\t{rtype}\t\r\nSummary Energy (MV/MeV)\t{type}\t\r\nSummary FieldSize X (cm)\t{sumx_mlc}\t\r\nSummary FieldSize Y (cm)\t{sumy_mlc}\t\r\nSummary Wedge Type\tOpen Field\t\r\nSummary Wedge Angle (degrees)\t0.00\t\r\nSummary Scan Type\tDepth Scan");
                sw.WriteLine($"Is PDI to PDD\t{tf_pdi}\t\r\n\r\nBEGIN DOSE TABLE\t\t\r\nAction\tSmooth\t\r\n\tX (cm)\tY (cm)\tZ (cm)\tRelative Dose (%)");

                /*** Enumerate over points and write to file ****/
                int co = 0;
                double adval = -1 * pdd_d[0].Position.y;

                /*** Write data to file ****/
                /*** NOTE MUST CONVERT TO WATER-TANK CO-ORDINATES. //So going from DICOM to SUNDOSE WT. ****/
                foreach (var qwerty in pdd_d) 
                {
                    /*** Water tank co-ords are x = crossline, y = inline, z = depth, units are 'cm' so must also convert from mm. ****/
                    /*** Water tank co-ords are x = crossline, y = inline, z = depth, units are 'cm' so must also convert from mm. ****/
                    bool res = Double.IsNaN(pdd_d[co].Value); //If it's NaN give a true value. Will invert it in the next line to determine if we should save the value to file.


                    try
                    {
                        if (!res)
                        {
                            sw.WriteLine($"\t{Math.Round(pdd_d[co].Position.x / 10, 4)}\t{Math.Round(pdd_d[co].Position.z / 10, 4)}\t{Math.Round((pdd_d[co].Position.y + adval) / 10, 4)}\t{Math.Round(pdd_d[co].Value * 100 / normaval_pdd, 4)}"); //Get rid of these values - replace with actual ones. Keep tab formatting same.
                        }
                        co++;

                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("No PDDs written due to error. Most likely none were create. Full error message; The process failed: {0}", e.ToString());
                        throw;
                    }
                }

                /*** Add raw data to footer ****/
                sw.WriteLine($"\r\n\r\nEND DOSE TABLE\r\nSCAN HEADER\r\nFACILITY INFORMATION\r\nInstitution\tCWCCC\t\r\nDelivery System\t{tmtmach}\t\r\nDelivery System Manufacturer\t\t\r\nDelivery System Model #\t\t\r\nDelivery System Serial #\tH195586\t\r\nField Detector Model #\tCC13 #96910 Orange\t\r\nField Detector Serial #\t96910\t\r\nReference Detector Model #\tCC13 #96830 Orange\t\r\nReference Detector Serial #\t96830\t\r\nSNC EQUIPMENT\r\nApplication Programming Interface\t4.0.1.8\t\r\nHardware Device Interface\t4.0.0.354\t\r\nSunSCAN 3D Model #\tNot Available\t\r\nSunSCAN 3D Serial #\t64368013\t\r\nSunSCAN 3D Firmware\t1.0.4.2\t\r\nDrive Factor Diameter (pulse/mm)\t0\t\r\nDrive Factor Vertical (pulse/mm)\t0\t\r\nDrive Factor Ring (pulse/mm)\t0\t\r\nElectrometer Model #\tNot Available\t\r\nElectrometer Serial #\t64368013\t\r\nElectrometer Firmware\t1.0.1.3\t\r\nLeveling Platform Model #\tNot Available\t\r\nLeveling Platform Serial #\tNot Available\t\r\nLift Table Model #\tNot Available\t\r\nLift Table Serial #\tNot Available\t\r\nReservoir Model #\tNot Available\t\r\nReservoir Serial #\tNot Available\t\r\nBootloader Firmware\tNot Available\t\r\nError Status Register\tNot Available\t\r\nFormatted Application Version\tNot Available\t\r\nFormatted Bootblock Version\tNot Available\t\r\nFormatted PMD OS Firmware Version\tNot Available\t\r\nHardware Version\tNot Available\t\r\nDELIVERY SYSTEM\r\nGantry Angle (degrees)\t0\t\r\nCollimator Angle (degrees)\t0.00\t\r\nCollimation Type\tJaws and MLC\t\r\nWedge Type\tOpen Field\t\r\nWedge Angle (degrees)\t0.00\t\r\nWedge Direction\t\t\r\nField Size X (cm)\t{sumx_mlc}\t\r\nField Size Y (cm)\t{sumy_mlc}\t\r\nField Shape\tSquare\t\r\nMeasurement Unit\tcm\t\r\nCollimator Position Jaws X1 (cm) {xfs}\t\t\r\nCollimator Position Jaws X2 (cm) {xfs}\t\t\r\nCollimator Position Jaws Y1 (cm) {yfs}\t\t\r\nCollimator Position Jaws Y2 (cm) {yfs}\t\t\r\nCollimator Position MLC X1 (cm) {xfs_mlc}\t\t\r\nCollimator Position MLC X2 (cm) {mlc2}\t\t\r\nCollimator Position MLC Y1 (cm) {yfs_mlc}\t\t\r\nCollimator Position MLC Y2 (cm) {mlc2}\t\t\r\nSETUP PARAMETERS\r\nRing Center (cm)\t0\t\r\nAngle Offset (degrees)\t0\t\r\nHysteresis Minus (cm)\t0\t\r\nMEASUREMENT DETAILS\r\nComments\t\t\r\nScan Id\t3384\t\r\nScan Date\t02/17/2024 03:57\t\r\nScan Type\tDepth Scan\t\r\nScan Medium\tWater\t\r\nSource to Surface Distance (cm)\t{ssd}\t\r\nIon Chamber Equivalent Model\t\t\r\nScan Source\tSunSCAN 3D\t\r\nSunSCAN\tTrue\t\r\nMeasurement Mode\tContinuous\t\r\nScan Speed (cm/s)\t0.05 cm/second\t\r\nStyle\t7\t\r\nEMF Spacing (cm)\t0.050\t\r\nOptimized Rotation\tTrue\t\r\nDiameter Drive Scan Direction\tFalse\t\r\nAdditional Scan Range (cm)\t5.00\t\r\nIntegrated Measurement\tFalse\t\r\nEffective Point of Measurement (cm)\t0.15\t\r\nDetector Bias Voltage (V)\t304.34\t\r\nReference Detector Bias Voltage (V)\t302.80\t\r\nField Background Rate (counts/update)\t0.005\t\r\nReference Background Rate (counts/update)\t0.00207035175879397\t\r\nNormalization Value (Field/Reference)\t0.802148539760211\t\r\nPulse Normalized\tFalse\t\r\nMeasurement Current\tDynamic\t\r\nOverscan Amount\t\t\r\nOffset Detector Holder\tFalse\t\r\n\r\n\r\nBEGIN RAW DATA\t\t\t\t\r\n\tSequence\tTheta\tDiameter\tX (cm)\tY (cm)\tZ (cm)\tDelta Time (s)\tReference Cumulative Counts\tField Cumulative Counts\tCorrected Ratio\tCumulative Pulses\tVoltage (V)\tElectrometer Timer (s)\tMotion Timer (s)\tPressure\tInternal Temperature\tExternal Temperature\tPlus 5 Sensor\tReference: Pulse Size (counts)\tReference: Timed Measurement\tReference: Cumulative Buckets\tReference: Cumulative Counts (counts)\tReference: Cumulative Counts Corrected (counts)\tReference: Number of Updates when Measuring Charge\tReference: Number of Updates when Measuring Bucket Charge\tReference: Rail\tField: Pulse Size (counts)\tField: Timed Measurement\tField: Cumulative Buckets\tField: Cumulative Counts (counts)\tField: Cumulative Counts Corrected (counts)\tField: Number of Updates when Measuring Charge\tField: Number of Updates when Measuring Bucket Charge\tField: Rail\tTPR Sensor Calibrated\tTPR Sensor Raw\tTPR Buildup\tReference Voltage (V)\tActual X (cm)\tActual Y (cm)\tActual Z (cm)\tApplied Pulse Count\tRails Hit\r\n\t1\t0\t18.051\t18.051\t0\t2.8\t0.1\t81935\t467\t0.00563875845749496\t12\t304.2573\t0.1\t0.0989\t0\t0\t-273.15\t0\t81961\t0\t0\t81935\t81935\t1000\t0\t0\t497\t0\t0\t467\t467\t1000\t0\t0\t0\t0\t5.394\t303.0365\t18.051\t0\t2.8\t0\tFalse\r\n\t2\t0\t18.051\t18.051\t0\t2.8\t0.05\t116054\t655\t0.00543701832617846\t17\t304.3488\t0.15\t0.145\t0\t0\t-273.15\t0\t116102\t0\t0\t116054\t116054\t1500\t0\t0\t691\t0\t0\t655\t655\t1500\t0\t0\t0\t0\t5.394\t302.8992\t18.051\t0\t2.8\t0\tFalse");
                sw.WriteLine("END RAW DATA\r\nEND SCAN\t");
                sw.WriteLine($"Created By: {_currentUserId}");
                sw.Close();


                Console.WriteLine($"Success, PDDs/etc.. exported to: {fullPath}");
                Console.WriteLine($"Note that PDDs are given with respect to the reference point (e.g. 0 cm = reference point depth)");
            }
        }

        static void ExportAllProfileWithSameSettings()
        {
            Console.WriteLine($"Exporting profiles for all points contained within the plan");
            int countr = 0;
            //redefined in terms of plan reference points, instead of patient ref points b.c. for some reason it's become null here...
            
            //if (planSelec.ReferencePoints.ElementAt(countr) == null) { Console.WriteLine($"[Found the null]"); }
           // else { Console.WriteLine($"[was not the null.]"); }
                foreach (var alldem in planSelec.ReferencePoints)
                { //PUT STUFF HERE, s.t. condition is met:
                    Console.WriteLine($"[{countr}]");

                    if (planSelec.ReferencePoints.ElementAt(countr).HasLocation(planSelec) == true && alldem != null) //a variable is null - trying to troubleshoot which one. Can't be this one bc it's listing it.
                    {



                        // Console.WriteLine($"[{countr}].{alldem.Name} - {patient.ReferencePoints.ElementAt(countr)}");

                        // var validrefpts = 
                        refSelec = planSelec.ReferencePoints.ElementAt(countr); //ref point selected - based on counting through valid ref points contained within the plan.
                        reffy = refSelec.GetReferencePointLocation(planSelec);

                        VVector startCrosslineMUL = new VVector((reffy.x - symdist), reffy.y, reffy.z);
                        VVector stopCrosslineMUL = new VVector((reffy.x + symdist), reffy.y, reffy.z);
                        /*** EXTRACT PROFILEs for each ref point: ****/
                        proff = planSelec.Dose.GetDoseProfile(startCrosslineMUL, stopCrosslineMUL, preallocatedBuffer: usedis);
                        /*** Start and stop points for inline ****/
                        VVector startInlineMUL = new VVector(reffy.x, reffy.y, (reffy.z + symdist));
                        VVector stopInlineMUL = new VVector(reffy.x, reffy.y, (reffy.z - symdist));
                        /*** get profile based on start and stop and double array ****/
                        proffI = planSelec.Dose.GetDoseProfile(startInlineMUL, stopInlineMUL, preallocatedBuffer: usedis);

                        /*** display profile points ****/
                        conttta = 0;
                        maxyv = 0;
                        mvI = 0;
                        foreach (var v in proffI)
                        {
                            Console.WriteLine($"{proffI[conttta].Value}");
                            if ((double)proffI[conttta].Value > maxyv)
                            {
                                mvI = proffI[conttta].Value;
                                maxyv = mvI;

                            }
                            conttta++;
                        }

                        Console.WriteLine($" ");
                        Console.WriteLine($" ");

                        /*** Display crossplane profile points ****/
                        Console.WriteLine($"Crossplane profiles (assuming HFS setup) [Gy]:");
                        contta = 0;
                        maxyvv = 0;
                        mvv = 0;
                        foreach (var vv in proff)
                        {
                            Console.WriteLine($"{(proff[contta].Value)}"); //  /normaval
                            if (proff[contta].Value > maxyvv)
                            {
                                mvv = proff[contta].Value;
                                maxyvv = mvv;

                            }
                            contta++;
                        }

                    /*** Normalisation ****/
                    /*** create string array to select normalisation value ****/
                    /*
                    //[0] - None (units will maybe be in cGy? or in %, depending on plan settings? not sure). 
                    //[1] - Maximum value anywhere along the extraced PDD/Profile
                    //[3] - Value at the selected reference point
                    //[4] - Custom value - will be next prompted to input a value to normalise to. */

                    //* Commented out the below, as normalisation is set to same settings for all in this subroutine.
                    //Console.WriteLine($"What normalisation value would you like to use?");
                    //string[] cwla = { "None", "Max value along profile", "Value at the reference point", "Custom Value - if selected will be prompted to input." };
                    //var cwlaList = new List<string>(cwla);
                    //int pdproSa = GetUserListSelection(cwlaList);
                    
                    ///*** Display selected value ****/
                    //var normavalname = cwla[pdproSa]; //Option selected
                    //Console.WriteLine($"You have selected, {pdproSa}. - {normavalname}");
                    ////
                    ///*//

                    //OK, now
                    //NORMALISE:
                    normaval = 1;
                        normavalINPLANE = 1;
                    /*** If custom normalisation ****/
                    if (pdproSa == 3)
                    {
                        /*** Have forced displayed units to whatever the system has as the absolute units. For our centre this is Gy (but others may have cGy). ****/
                        Console.WriteLine($"You have selected custom normalisation. Input the normalisation value in units of UNITs: Gy."); //
                        int normya = -1;
                        if (!int.TryParse(Console.ReadLine(), out normya))
                        {
                            /*** Ensure value selected is numeric otherwise error ****/
                            Console.WriteLine("Invalid Selection"); return;
                        }
                        /*** Set custom normalisation values ****/
                        normaval = normya;
                        normavalINPLANE = normaval;
                    }
                    /*** If nor mornalisation value set ****/
                    else if (pdproSa == 0)
                    {
                        Console.WriteLine($"You have selected No normalisation. The profile units will remain in Gray.");
                        /*** Set norm values to 100 ****/
                        normaval = 100;
                        normavalINPLANE = normaval;
                    }
                    /*** if reference point set ****/
                    else if (pdproSa == 2)
                    {
                        decimal dd = (proff.Count / 2);
                        int ah1 = (int)Math.Round(dd, 0);


                        Console.WriteLine($"You have selected to normalise to the value at the reference point. Normalisation values will be:");
                        Console.WriteLine($" {proff[ah1].Value} , units of: {proff.Unit}");

                        /*** Set value from reference point ****/
                        normaval = (double)proff[ah1].Value;
                        normavalINPLANE = normaval;
                    }
                    /*** Set the max point as norm point ****/
                    else if (pdproSa == 1)
                    {
                        Console.WriteLine($"You have selected to normalise to the max value along the profile. Normalisation values will be:");
                        Console.WriteLine($"Crossplane {mvv} ,units of: {proff.Unit}");
                        Console.WriteLine($"Inplane {mvI} ,units of: {proffI.Unit}");

                        /*** Set norm values from max point ****/
                        normaval = (double)mvv;
                        normavalINPLANE = (double)mvI;
                    }
                    else { //do nothing.
                           }

                        /*** Apply normaization ****/
                        contta = 0;
                        foreach (var v in proff)
                        {
                            Console.WriteLine($"{((proff[contta].Value * 100) / normaval)}"); //  /normaval

                            contta++;
                        }

                        /*** Display profile points as normalised values ****/
                        Console.WriteLine($"Inplane profiles are (normalised) - expressed as a %:");
                        conttta = 0;
                        foreach (var v in proffI)
                        {
                            Console.WriteLine($"{((proffI[conttta].Value * 100) / normavalINPLANE)}");

                            conttta++;
                        }

                        /*** Set File name ****/
                        string fileName = $"{planSelec}_{refSelec}_{countr}_TPSexport_Crossline.snctxt";
                        
                        fileName = SanitizeFileName(fileName);
                        /*** Set Folder name ****/
                        string fullPath = destinationFolder + fileName;
                        /*** Create directory to save file ****/
                        if (!Directory.Exists(destinationFolder)) Directory.CreateDirectory(destinationFolder);

                        /*** Write data to file ****/
                        StreamWriter sw = new StreamWriter(fullPath);
                        sw.WriteLine("Tab-Delimited Scan Output");
                        sw.WriteLine("FILE HEADER");
                        sw.WriteLine($"File Name\t{fileName}");
                        sw.WriteLine("File Date\t02/20/2077 00:01");
                        sw.WriteLine($"File Export Version\t4.0.1.8\t\r\nFile Version\t5.0\t\r\nFile Scan Count\t1\t\r\nBEGIN SCAN\r\nSummary Comments\t\t\r\nSummary Beam Type\t{modality}\t\r\nBeam Type\t{modality}\t\r\nEnergy (MV / MeV)\t{type}\t\r\nRate Type\t{rtype}\t\r\nSummary Energy (MV/MeV)\t{type}\t\r\nSummary FieldSize X (cm)\t{sumx_mlc}\t\r\nSummary FieldSize Y (cm)\t{sumy_mlc}\t\r\nSummary Wedge Type\tOpen Field\t\r\nSummary Wedge Angle (degrees)\t0.00\t\r\nSummary Scan Type\tCrossline");
                        sw.WriteLine($"Is PDI to PDD\t{tf_pdi}\t\r\n\r\nBEGIN DOSE TABLE\t\t\r\nAction\tSmooth\t\r\n\tX (cm)\tY (cm)\tZ (cm)\tRelative Dose (%)");


                        /*** Enumerate of points and write to file ****/
                        /*** NOTE MUST CONVERT TO WATER-TANK CO-ORDINATES. //So going from DICOM to SUNDOSE WT. ****/
                        int co = 0;
                        try
                        {
                            VVector startCrosslineexp = new VVector((reffy.x - symdist), reffy.y, reffy.z); //required to redefine doseprofile prior to export.
                            VVector stopCrosslineexp = new VVector((reffy.x + symdist), reffy.y, reffy.z);
                            //usedis = new double[(int)Math.Ceiling((startCrosslineexp - stopCrosslineexp).Length / stepsize)];
                            proff = planSelec.Dose.GetDoseProfile(startCrosslineexp, stopCrosslineexp, preallocatedBuffer: usedis);

                        foreach (var qwerty in proff)
                            {
                            /*** Water tank co-ords are x = crossline, y = inline, z = depth, units are 'cm' so must also convert from mm. ****/
                            /*** whereas DICOM co-ordinates (in HFS) are x = crossline, y = depth, z = inline. ****/
                            bool res = Double.IsNaN(proff[co].Value); //If it's NaN give a true value. Will invert it in the next line to determine if we should save the value to file.
                            

                            if (!res)
                            {
                                sw.WriteLine($"\t{Math.Round(proff[co].Position.x / 10, 4)}\t{Math.Round(proff[co].Position.z / 10, 4)}\t{Math.Round(reffy.y, 4)}\t{Math.Round(proff[co].Value * 100 / normaval, 4)}"); //Keep tab formatting same.
                            }
                                co++;
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("No Crosslines written due to error. Most likely none were create. Full error message; The process failed: {0}", e.ToString());
                            throw;
                        }

                        /*** Add raw data to file footer ****/
                        sw.WriteLine($"\r\n\r\nEND DOSE TABLE\r\nSCAN HEADER\r\nFACILITY INFORMATION\r\nInstitution\tCWCCC\t\r\nDelivery System\t{tmtmach}\t\r\nDelivery System Manufacturer\t\t\r\nDelivery System Model #\t\t\r\nDelivery System Serial #\tH195586\t\r\nField Detector Model #\tCC13 #96910 Orange\t\r\nField Detector Serial #\t96910\t\r\nReference Detector Model #\tCC13 #96830 Orange\t\r\nReference Detector Serial #\t96830\t\r\nSNC EQUIPMENT\r\nApplication Programming Interface\t4.0.1.8\t\r\nHardware Device Interface\t4.0.0.354\t\r\nSunSCAN 3D Model #\tNot Available\t\r\nSunSCAN 3D Serial #\t64368013\t\r\nSunSCAN 3D Firmware\t1.0.4.2\t\r\nDrive Factor Diameter (pulse/mm)\t0\t\r\nDrive Factor Vertical (pulse/mm)\t0\t\r\nDrive Factor Ring (pulse/mm)\t0\t\r\nElectrometer Model #\tNot Available\t\r\nElectrometer Serial #\t64368013\t\r\nElectrometer Firmware\t1.0.1.3\t\r\nLeveling Platform Model #\tNot Available\t\r\nLeveling Platform Serial #\tNot Available\t\r\nLift Table Model #\tNot Available\t\r\nLift Table Serial #\tNot Available\t\r\nReservoir Model #\tNot Available\t\r\nReservoir Serial #\tNot Available\t\r\nBootloader Firmware\tNot Available\t\r\nError Status Register\tNot Available\t\r\nFormatted Application Version\tNot Available\t\r\nFormatted Bootblock Version\tNot Available\t\r\nFormatted PMD OS Firmware Version\tNot Available\t\r\nHardware Version\tNot Available\t\r\nDELIVERY SYSTEM\r\nGantry Angle (degrees)\t0\t\r\nCollimator Angle (degrees)\t0.00\t\r\nCollimation Type\tJaws and MLC\t\r\nWedge Type\tOpen Field\t\r\nWedge Angle (degrees)\t0.00\t\r\nWedge Direction\t\t\r\nField Size X (cm)\t{sumx_mlc}\t\r\nField Size Y (cm)\t{sumy_mlc}\t\r\nField Shape\tSquare\t\r\nMeasurement Unit\tcm\t\r\nCollimator Position Jaws X1 (cm) {xfs}\t\t\r\nCollimator Position Jaws X2 (cm) {xfs}\t\t\r\nCollimator Position Jaws Y1 (cm) {yfs}\t\t\r\nCollimator Position Jaws Y2 (cm) {yfs}\t\t\r\nCollimator Position MLC X1 (cm) {xfs_mlc}\t\t\r\nCollimator Position MLC X2 (cm) {mlc2}\t\t\r\nCollimator Position MLC Y1 (cm) {yfs_mlc}\t\t\r\nCollimator Position MLC Y2 (cm) {mlc2}\t\t\r\nSETUP PARAMETERS\r\nRing Center (cm)\t25.218\t\r\nAngle Offset (degrees)\t4.53\t\r\nHysteresis Minus (cm)\t0\t\r\nMEASUREMENT DETAILS\r\nComments\t\t\r\nScan Id\t3384\t\r\nScan Date\t02/17/2024 03:57\t\r\nScan Type\tCrossline\t\r\nScan Medium\tWater\t\r\nSource to Surface Distance (cm)\t{ssd}\t\r\nIon Chamber Equivalent Model\t\t\r\nScan Source\tSunSCAN 3D\t\r\nSunSCAN\tTrue\t\r\nMeasurement Mode\tContinuous\t\r\nScan Speed (cm/s)\t0.05 cm/second\t\r\nStyle\t7\t\r\nEMF Spacing (cm)\t0.050\t\r\nOptimized Rotation\tTrue\t\r\nDiameter Drive Scan Direction\tFalse\t\r\nAdditional Scan Range (cm)\t5.00\t\r\nIntegrated Measurement\tFalse\t\r\nEffective Point of Measurement (cm)\t0.15\t\r\nDetector Bias Voltage (V)\t304.34\t\r\nReference Detector Bias Voltage (V)\t302.80\t\r\nField Background Rate (counts/update)\t0.005\t\r\nReference Background Rate (counts/update)\t0.00207035175879397\t\r\nNormalization Value (Field/Reference)\t0.802148539760211\t\r\nPulse Normalized\tFalse\t\r\nMeasurement Current\tDynamic\t\r\nOverscan Amount\t\t\r\nOffset Detector Holder\tFalse\t\r\n\r\n\r\nBEGIN RAW DATA\t\t\t\t\r\n\tSequence\tTheta\tDiameter\tX (cm)\tY (cm)\tZ (cm)\tDelta Time (s)\tReference Cumulative Counts\tField Cumulative Counts\tCorrected Ratio\tCumulative Pulses\tVoltage (V)\tElectrometer Timer (s)\tMotion Timer (s)\tPressure\tInternal Temperature\tExternal Temperature\tPlus 5 Sensor\tReference: Pulse Size (counts)\tReference: Timed Measurement\tReference: Cumulative Buckets\tReference: Cumulative Counts (counts)\tReference: Cumulative Counts Corrected (counts)\tReference: Number of Updates when Measuring Charge\tReference: Number of Updates when Measuring Bucket Charge\tReference: Rail\tField: Pulse Size (counts)\tField: Timed Measurement\tField: Cumulative Buckets\tField: Cumulative Counts (counts)\tField: Cumulative Counts Corrected (counts)\tField: Number of Updates when Measuring Charge\tField: Number of Updates when Measuring Bucket Charge\tField: Rail\tTPR Sensor Calibrated\tTPR Sensor Raw\tTPR Buildup\tReference Voltage (V)\tActual X (cm)\tActual Y (cm)\tActual Z (cm)\tApplied Pulse Count\tRails Hit\r\n\t1\t0\t18.051\t18.051\t0\t2.8\t0.1\t81935\t467\t0.00563875845749496\t12\t304.2573\t0.1\t0.0989\t0\t0\t-273.15\t0\t81961\t0\t0\t81935\t81935\t1000\t0\t0\t497\t0\t0\t467\t467\t1000\t0\t0\t0\t0\t5.394\t303.0365\t18.051\t0\t2.8\t0\tFalse\r\n\t2\t0\t18.051\t18.051\t0\t2.8\t0.05\t116054\t655\t0.00543701832617846\t17\t304.3488\t0.15\t0.145\t0\t0\t-273.15\t0\t116102\t0\t0\t116054\t116054\t1500\t0\t0\t691\t0\t0\t655\t655\t1500\t0\t0\t0\t0\t5.394\t302.8992\t18.051\t0\t2.8\t0\tFalse");
                        sw.WriteLine("END RAW DATA\r\nEND SCAN\t");
                        sw.WriteLine($"Created By: {_currentUserId}");
                        sw.Close();

                        Console.WriteLine($"Success, profiles/etc.. exported to: {fullPath}");

                        /*** Write Inline to file ****/
                        /*** Set file name ****/
                        string fileName2 = $"{planSelec}_{refSelec}_{countr}_TPSexport_Inline.snctxt";
                        fileName2 = SanitizeFileName(fileName2);
                        /*** Set Folder name ****/
                        string fullPath2 = destinationFolder + fileName2;
                        /*** Create directory to save file ****/
                        if (!Directory.Exists(destinationFolder)) Directory.CreateDirectory(destinationFolder);

                        /*** write data to file ****/
                        StreamWriter sw2 = new StreamWriter(fullPath2);
                        sw2.WriteLine("Tab-Delimited Scan Output");
                        sw2.WriteLine("FILE HEADER");
                        sw2.WriteLine($"File Name\t{fileName2}");
                        sw2.WriteLine("File Date\t02/20/2077 00:02");
                        sw2.WriteLine($"File Export Version\t4.0.1.8\t\r\nFile Version\t5.0\t\r\nFile Scan Count\t1\t\r\nBEGIN SCAN\r\nSummary Comments\t\t\r\nSummary Beam Type\t{modality}\t\r\nBeam Type\t{modality}\t\r\nEnergy (MV / MeV)\t{type}\t\r\nRate Type\t{rtype}\t\r\nSummary Energy (MV/MeV)\t{type}\t\r\nSummary FieldSize X (cm)\t{sumx_mlc}\t\r\nSummary FieldSize Y (cm)\t{sumy_mlc}\t\r\nSummary Wedge Type\tOpen Field\t\r\nSummary Wedge Angle (degrees)\t0.00\t\r\nSummary Scan Type\tInline");
                        sw2.WriteLine($"Is PDI to PDD\t{tf_pdi}\t\r\n\r\nBEGIN DOSE TABLE\t\t\r\nAction\tSmooth\t\r\n\tX (cm)\tY (cm)\tZ (cm)\tRelative Dose (%)");

                        /*** Enumerate over inline values and write to file ****/
                        /*** NOTE MUST CONVERT TO WATER-TANK CO-ORDINATES. //So going from DICOM to SUNDOSE WT. ****/
                        co = 0;
                        try
                        {
                             VVector startInlineexp = new VVector(reffy.x, reffy.y, (reffy.z + symdist)); //redefine prior to export. Unsure why needed to do these but fixes the bug -JH.
                             VVector stopInlineexp = new VVector(reffy.x, reffy.y, (reffy.z - symdist));
                            /*** Get Profile using start and stop points and double array ****/
                            proffI = planSelec.Dose.GetDoseProfile(startInlineexp, stopInlineexp, preallocatedBuffer: usedis);

                        foreach (var qwerty in proffI) //enumerate over cross-plane.//NOTE MUST CONVERT TO WATER-TANK CO-ORDINATES. //So going from DICOM to SUNDOSE WT.
                            {
                            /*** Water tank co-ords are x = crossline, y = inline, z = depth, units are 'cm' so must also convert from mm. ****/
                            /*** whereas DICOM co-ordinates (in HFS) are x = crossline, y = depth, z = inline. ****/
                            bool res = Double.IsNaN(proffI[co].Value); //If it's NaN give a true value. Will invert it in the next line to determine if we should save the value to file.

                            if (!res)
                            {
                                sw2.WriteLine($"\t{Math.Round(proffI[co].Position.x / 10, 4)}\t{Math.Round(proffI[co].Position.z / 10, 4)}\t{Math.Round(reffy.y, 1)}\t{Math.Round(proffI[co].Value * 100 / normavalINPLANE, 4)}"); //Get rid of these values - replace with actual ones. Keep tab formatting same.
                            }
                                co++;
                            }
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("No Crosslines written due to error. Most likely none were create. Full error message; The process failed: {0}", e.ToString());
                            throw;
                        }

                        /*** Add raw data to footer ****/
                        sw2.WriteLine($"\r\n\r\nEND DOSE TABLE\r\nSCAN HEADER\r\nFACILITY INFORMATION\r\nInstitution\tCWCCC\t\r\nDelivery System\t{tmtmach}\t\r\nDelivery System Manufacturer\t\t\r\nDelivery System Model #\t\t\r\nDelivery System Serial #\tH195586\t\r\nField Detector Model #\tCC13 #96910 Orange\t\r\nField Detector Serial #\t96910\t\r\nReference Detector Model #\tCC13 #96830 Orange\t\r\nReference Detector Serial #\t96830\t\r\nSNC EQUIPMENT\r\nApplication Programming Interface\t4.0.1.8\t\r\nHardware Device Interface\t4.0.0.354\t\r\nSunSCAN 3D Model #\tNot Available\t\r\nSunSCAN 3D Serial #\t64368013\t\r\nSunSCAN 3D Firmware\t1.0.4.2\t\r\nDrive Factor Diameter (pulse/mm)\t0\t\r\nDrive Factor Vertical (pulse/mm)\t0\t\r\nDrive Factor Ring (pulse/mm)\t0\t\r\nElectrometer Model #\tNot Available\t\r\nElectrometer Serial #\t64368013\t\r\nElectrometer Firmware\t1.0.1.3\t\r\nLeveling Platform Model #\tNot Available\t\r\nLeveling Platform Serial #\tNot Available\t\r\nLift Table Model #\tNot Available\t\r\nLift Table Serial #\tNot Available\t\r\nReservoir Model #\tNot Available\t\r\nReservoir Serial #\tNot Available\t\r\nBootloader Firmware\tNot Available\t\r\nError Status Register\tNot Available\t\r\nFormatted Application Version\tNot Available\t\r\nFormatted Bootblock Version\tNot Available\t\r\nFormatted PMD OS Firmware Version\tNot Available\t\r\nHardware Version\tNot Available\t\r\nDELIVERY SYSTEM\r\nGantry Angle (degrees)\t0\t\r\nCollimator Angle (degrees)\t0.00\t\r\nCollimation Type\tJaws and MLC\t\r\nWedge Type\tOpen Field\t\r\nWedge Angle (degrees)\t0.00\t\r\nWedge Direction\t\t\r\nField Size X (cm)\t{sumx_mlc}\t\r\nField Size Y (cm)\t{sumy_mlc}\t\r\nField Shape\tSquare\t\r\nMeasurement Unit\tcm\t\r\nCollimator Position Jaws X1 (cm) {xfs}\t\t\r\nCollimator Position Jaws X2 (cm) {xfs}\t\t\r\nCollimator Position Jaws Y1 (cm) {yfs}\t\t\r\nCollimator Position Jaws Y2 (cm) {yfs}\t\t\r\nCollimator Position MLC X1 (cm) {xfs_mlc}\t\t\r\nCollimator Position MLC X2 (cm) {mlc2}\t\t\r\nCollimator Position MLC Y1 (cm) {yfs_mlc}\t\t\r\nCollimator Position MLC Y2 (cm) {mlc2}\t\t\r\nSETUP PARAMETERS\r\nRing Center (cm)\t0\t\r\nAngle Offset (degrees)\t0\t\r\nHysteresis Minus (cm)\t0\t\r\nMEASUREMENT DETAILS\r\nComments\t\t\r\nScan Id\t3384\t\r\nScan Date\t02/17/2024 03:57\t\r\nScan Type\tInline\t\r\nScan Medium\tWater\t\r\nSource to Surface Distance (cm)\t{ssd}\t\r\nIon Chamber Equivalent Model\t\t\r\nScan Source\tSunSCAN 3D\t\r\nSunSCAN\tTrue\t\r\nMeasurement Mode\tContinuous\t\r\nScan Speed (cm/s)\t0.05 cm/second\t\r\nStyle\t7\t\r\nEMF Spacing (cm)\t0.050\t\r\nOptimized Rotation\tTrue\t\r\nDiameter Drive Scan Direction\tFalse\t\r\nAdditional Scan Range (cm)\t5.00\t\r\nIntegrated Measurement\tFalse\t\r\nEffective Point of Measurement (cm)\t0.15\t\r\nDetector Bias Voltage (V)\t304.34\t\r\nReference Detector Bias Voltage (V)\t302.80\t\r\nField Background Rate (counts/update)\t0.005\t\r\nReference Background Rate (counts/update)\t0.00207035175879397\t\r\nNormalization Value (Field/Reference)\t0.802148539760211\t\r\nPulse Normalized\tFalse\t\r\nMeasurement Current\tDynamic\t\r\nOverscan Amount\t\t\r\nOffset Detector Holder\tFalse\t\r\n\r\n\r\nBEGIN RAW DATA\t\t\t\t\r\n\tSequence\tTheta\tDiameter\tX (cm)\tY (cm)\tZ (cm)\tDelta Time (s)\tReference Cumulative Counts\tField Cumulative Counts\tCorrected Ratio\tCumulative Pulses\tVoltage (V)\tElectrometer Timer (s)\tMotion Timer (s)\tPressure\tInternal Temperature\tExternal Temperature\tPlus 5 Sensor\tReference: Pulse Size (counts)\tReference: Timed Measurement\tReference: Cumulative Buckets\tReference: Cumulative Counts (counts)\tReference: Cumulative Counts Corrected (counts)\tReference: Number of Updates when Measuring Charge\tReference: Number of Updates when Measuring Bucket Charge\tReference: Rail\tField: Pulse Size (counts)\tField: Timed Measurement\tField: Cumulative Buckets\tField: Cumulative Counts (counts)\tField: Cumulative Counts Corrected (counts)\tField: Number of Updates when Measuring Charge\tField: Number of Updates when Measuring Bucket Charge\tField: Rail\tTPR Sensor Calibrated\tTPR Sensor Raw\tTPR Buildup\tReference Voltage (V)\tActual X (cm)\tActual Y (cm)\tActual Z (cm)\tApplied Pulse Count\tRails Hit\r\n\t1\t0\t18.051\t18.051\t0\t2.8\t0.1\t81935\t467\t0.00563875845749496\t12\t304.2573\t0.1\t0.0989\t0\t0\t-273.15\t0\t81961\t0\t0\t81935\t81935\t1000\t0\t0\t497\t0\t0\t467\t467\t1000\t0\t0\t0\t0\t5.394\t303.0365\t18.051\t0\t2.8\t0\tFalse\r\n\t2\t0\t18.051\t18.051\t0\t2.8\t0.05\t116054\t655\t0.00543701832617846\t17\t304.3488\t0.15\t0.145\t0\t0\t-273.15\t0\t116102\t0\t0\t116054\t116054\t1500\t0\t0\t691\t0\t0\t655\t655\t1500\t0\t0\t0\t0\t5.394\t302.8992\t18.051\t0\t2.8\t0\tFalse");
                        sw2.WriteLine("END RAW DATA\r\nEND SCAN\t");
                        sw2.WriteLine($"Created By: {_currentUserId}");
                        sw2.Close();

                        countr++;
                    }
                    else
                    {
                        Console.WriteLine($"Skipping: [{countr}].{alldem.Name} - {planSelec.ReferencePoints.ElementAt(countr)} - Because it is INVALID - is not contained in this plan.");
                        countr++;
                    }

                }
            
        }

        public static string SanitizeFileName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_'); // Replace invalid characters with underscore
            }
            return name;
        }

        static int GetUserListSelection(List<string> list)
        {
            int countt = 0;

            foreach (var cct in list)
            {
                /*** Display list of items and IDS for user to select from ****/
                Console.WriteLine($"[{countt}]. {cct}");
                countt++;
            }
            Console.WriteLine("Select a Item From the List");
            int ppSelect = -1; 
            if (!int.TryParse(Console.ReadLine(), out ppSelect))
            {
                /*** Selection by number otherwise return error ****/
                Console.WriteLine("Invalid Selection"); return ppSelect;

            }
            return ppSelect;
        }
    }
}

