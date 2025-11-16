using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using VMS.TPS.Common.Model.API;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Windows.Navigation;
using VMS.TPS.Common.Model.Types;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.EMMA;
using System.Numerics;
using Microsoft.Office.Interop.Excel;
using System.Threading;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Linq.Expressions;

namespace BeginPlanCheck2
{
    public class CheckingFunctions
    {
        public static string Getdayte()
        {
            DateTime localDate = DateTime.Now;
            //string culture = "en-GB";
            //var cultura = new CultureInfo("en-GB");
            string values = localDate.ToString("MM/dd/yyyy"); //localDate.ToString(cultura);
            //string values = localDate.ToString();
            values = values.Substring(0, 10); //first 10 characters are the dd/mm/yyyy; date in correct format.
            return values;
        }


        //{ "Thomas Eade", "Patrick Horsley", "Kandeepan Thuraisingam", "Johnson Lam", "Simon Tang", "Gillian Lamoury", "George Hruby", "Michael Back", "Dasantha Jayamanne", "Gabrielle Metz", "Susan Carroll"}

        //Definition for method to infer staff-initials from user name. Not currently used.
        public static string InferUserN(string _currentUserId)
        {
            string userinitials = string.Empty;

            if (_currentUserId == "00000000") // 
            {
                userinitials = "JH";

            }
            //elseif(_currentUserId == "")
            //{
            //    userinitials = "";
            //}

            else { userinitials = _currentUserId; }

            return userinitials;
        }

        //Do one better, and use user name, to derive intials. //Added custom hard coding for Eugene (Shiaw J. T), based on stafflink number.
        public static string initialS(string _usern, string _currentUserId)
        {
            string ui = string.Empty;


            if (_currentUserId != "00000000" | _currentUserId != "00000" | _currentUserId != @"0000000" | _currentUserId != @"000000") //Messy.
            {

                StringBuilder sb = new StringBuilder();

                string un = _usern;

                foreach (char jj in un)
                {

                    //Console.WriteLine(jj);
                    if (char.IsUpper(jj))
                    {

                        sb.Append(jj, 1);


                    }


                }


                ui = sb.ToString();

                Console.WriteLine(ui);



            }
            else {

                ui = "ET"; //Then userID matches 0000000 which is 
            }
            return ui;



        }

        //Not in use.
        public static string OLDinitialS(string _usern)
        {

            StringBuilder sb = new StringBuilder();

            string un = _usern;

            foreach (char jj in un)
            {


                if (char.IsUpper(jj))
                {

                    sb.Append(jj, 1);


                }


            }


            string ui = sb.ToString();

            Console.WriteLine(ui);

            return ui;
        }
        //Not in use.

        public static string[] GetPlanApprvlHistry(PlanSetup plan) //tested. The combinedString2 variable will list all the approval history, can string search this for
                                                                   //treatmentapproved; or reviewed status. Then the enumerable value of this, will co-incide with the name of the approver/reviewer in combinedString2. 
                                                                   //Can then string test this against allowed tmtapprovers/reviewers (list of ROs).
        {
            //string reviewer = plan.PlanningApproverDisplayName.ToString();
            // string other = plan.TreatmentApproverDisplayName.ToString();

            string[] other2 = new string[plan.ApprovalHistory.Count()];
            string[] other3 = new string[plan.ApprovalHistory.Count()];

            int iii = 0;
            foreach (var s in plan.ApprovalHistory)
            {
                other2[iii] = s.ApprovalStatus.ToString(); //plan.ApprovalHistory
                other3[iii] = s.UserDisplayName.ToString();
                iii++;
            }


            string combinedString = string.Join(",", other2); // Status
            string combinedString2 = string.Join(",", other3); // Person whom altered.

            string[] outputs = { combinedString, combinedString2 };

            return outputs;
        }

        public static bool PlanReviewedOrApproved(PlanSetup plan)
        {
            bool result = false;

            string[] other2 = new string[plan.ApprovalHistory.Count()];
            string[] other3 = new string[plan.ApprovalHistory.Count()];

            int iii = 0;
            foreach (var s in plan.ApprovalHistory)
            {
                other2[iii] = s.ApprovalStatus.ToString(); //plan.ApprovalHistory
                other3[iii] = s.UserDisplayName.ToString();
                iii++;
            }


            string[] cbs2 = other2; // Status
            string[] cbs1 = other3; // Person whom altered.

            string[] authorisedreviers = { "Thomas Eade", "Patrick Horsley", "Kandeepan Thuraisingam", "Johnson Lam", "Simon Tang", "Gillian Lamoury", "George Hruby", "Michael Back", "Dasantha Jaymanne", "Gabrielle Metz", "Susan Carroll", "Sarah Bergamin", "Andrew Kneebone", "Joseph Chan", "Andrew Kneebone" };
            string[] allowedstati = { "Planning Approved", "Reviewed", "PlanningApproved" };//{ "Treatment Approved", "Reviewed", "TreatmentApproved" };
            string[] disallowedstati = { "Unapproved", "UnApproved", "Rejected", "Retired" }; //Yes - case sensitive.

            int incre = 0;
            int savincre = 0;
            //Check for treatment approved or reviewed in approval status
            //Check for Unapproved, Rejected, Retired stati. Save the increment at which these occur. Then when testing for approval status, ensure these occur after. 
            //(ordering starts at 0,1,...; for LAST modification.
            //Need to ensure that a plan has NOT been tmt approved/reviewed THEN subsequently unapproved/rejected/retired.
            //Check for treatment approved or reviewed in approval status
            foreach (var s in cbs1)
            {
                if (disallowedstati.Contains(cbs2[incre]))
                {

                    savincre = incre; //Will return an integer for the LAST modification (most recent) with a disallowed stati. Can then check later that this hasn't occured AFTER an allowed stati.
                    Console.WriteLine(incre.ToString());
                    break;
                }

                incre++;
            }

            incre = 0; //Reset.





            foreach (var s in cbs1)
            {
                if (authorisedreviers.Contains(s) && allowedstati.Contains(cbs2[incre]) && (incre < savincre))
                {
                    result = true; //Than plan is either reviewed or planning approved. By one of the authorised users within the list.
                    break;
                }
                incre++;
            }


            return result;
        }


        public static string GetNameOfPlanReviewerOrApprover(PlanSetup plan)
        {

            {
                //string reviewer = plan.PlanningApproverDisplayName.ToString();
                // string other = plan.TreatmentApproverDisplayName.ToString();

                string[] other2 = new string[plan.ApprovalHistory.Count()];
                string[] other3 = new string[plan.ApprovalHistory.Count()];

                int iii = 0;
                foreach (var s in plan.ApprovalHistory)
                {
                    other2[iii] = s.ApprovalStatus.ToString(); //plan.ApprovalHistory
                    other3[iii] = s.UserDisplayName.ToString();
                    iii++;
                }


                string combinedString = string.Join(",", other2); // Status
                string combinedString2 = string.Join(",", other3); // Person whom altered.

                //string[] outputs = { combinedString, combinedString2 };

                //////  

                //bool result = false;

                string reviewer = string.Empty;

                string[] cbs2 = combinedString.Split(',');
                string[] cbs1 = combinedString2.Split(',');

                string[] authorisedreviers = { "Thomas Eade", "Patrick Horsley", "Kandeepan Thuraisingam", "Johnson Lam", "Simon Tang", "Gillian Lamoury", "George Hruby", "Michael Back", "Dasantha Jaymanne", "Gabrielle Metz", "Susan Carroll", "Andrew Kneebone" };
                string[] allowedstati = {"Planning Approved", "Reviewed", "PlanningApproved"}; // { "Treatment Approved", "Reviewed" };

                int incre = 0;
                //Check for treatment approved or reviewed in approval status
                foreach (var s in cbs1)
                {
                    if (authorisedreviers.Contains(s) && allowedstati.Contains(cbs2[incre]))
                    {
                        //result = true; //Than plan is either reviewed or planning approved. By one of the authorised users within the list.
                        reviewer = s;
                        break;
                    }
                    incre++;
                }


                return reviewer;
            }
        }

        public static string[] GetPersciptionStatus(Course course)
        {
            string rx_status = string.Empty;
            string rx_modifier = string.Empty;
            string[] ap = { string.Empty, string.Empty, string.Empty };

            bool authorisedperson = false;
            try
            {

                rx_status = course.TreatmentPhases.First().Prescriptions.First().Status;
                rx_modifier = course.TreatmentPhases.First().Prescriptions.First().HistoryUserDisplayName.ToString();




                string[] authorisedreviers = { "Thomas Eade", "Patrick Horsley", "Kandeepan Thuraisingam", "Johnson Lam", "Simon Tang", "Gillian Lamoury", "George Hruby", "Michael Back", "Dasantha Jaymanne", "Gabrielle Metz", "Susan Carroll", "Sarah Bergamin", "Andrew Kneebone", "Joseph Chan" };
                string allowedstati = "Approved";


                //Check for Rx Approved
                if (authorisedreviers.Contains(rx_modifier) && allowedstati.Contains(rx_status))
                {
                    //result = true; //Than Rx is approved. And the last modifier Was one of the authorised users within the list. Presumably only last modifiers can approve!
                    authorisedperson = true;

                }

                ap[0] = rx_status; ap[1] = rx_modifier; ap[2] = authorisedperson.ToString();
            }
            catch
            {
                Console.WriteLine("No prescription set. Hence No authorised reviewers, or other error.");
            }
            return ap;
        }
   

        public static double GetDoseGird(PlanSetup plan) //tested. 
        {
            double DG_x = plan.Dose.XRes;
            double DG_y = plan.Dose.YRes;

            
            //double DG_z = plan.Dose.ZRes; %Dose-grid resolution in z direction depends upon slice-width.

            double DG = 99999; //Have an obviously error value for an unassigned dose-grid.

            if (DG_x == DG_y)
            {
                //Than dose-grid is (in X-Y directions) isotropic and has value (in cm) of:
                 DG = DG_x/10;
            }
            else
            {
                 DG = 999; //Dose-grid is NON-isotropic, return error value.
            }
            return DG;
        }

        public static double GetHyperArcDoseGird(PlanSetup plan)
        {
            double HA_DG_ = 9999999;
            string HA_DG = string.Empty;
            

            string[] resultString = { string.Empty, string.Empty };

            try
            {
                HA_DG = plan.PhotonCalculationOptions.ToArray()[2].ToString();

                resultString = Regex.Split(HA_DG, @"[^0-9\.]+");//Regex.Match(HA_DG, @"\d+").Value;
                                                                //string doubleArray = Regex.Split(HA_DG, @"[^0-9\.]+").Where(c => c != "." && c.Trim() != "");

                //Console.WriteLine(resultString[1]);

                try
                {
                    HA_DG_ = double.Parse(resultString[1]);
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }

                //HA_DG_
            }
            catch {
                Console.WriteLine("Error in dose-grid calculation check. Check there is even a dose-grid set!");
            }

            //CalculationGridSizeInCMForSRSAndHyperArc;
            return HA_DG_;
        }


        //Now check the plan value for the dose grid is less than or equal to the desired value.
        public static bool CheckDoseGirdSize(double desired, double planvalue) 
        {
            bool islessthanorequalto = false;

          if (planvalue <= desired)
            {
                islessthanorequalto = true;
            }
          
            return islessthanorequalto;
        }

        //Gets the physical material table.
        //NOT USED
        public static string GetPhysicalMatTable(PlanSetup plan) //NOT WORKING.
        {
            //string physicalmaterialtable = plan.StructureSet.Structures.FirstOrDefault().StructureCode.ToString();
            string physicalmaterialtable = plan.StructureSet.UID.ToString(); // gives UID of SS. Could query database for physical material table.
            //Alternatively assign HU to every single structure using the DEFAULT HU (which should be xb....yadda yadda); then check they match. FOR ALL structures? Seems convoluted.

            //string getthePMT = plan.StructureSet.Structures.First().GetAssignedHU()


            return physicalmaterialtable;
        }


        //Gets the Volume Dose Calculation Algorithm.
        public static string GetVolDoseCalcAlgo(PlanSetup plan) //WORKING.
        {
            string volumedosecalculationalgorithm = plan.PhotonCalculationModel.ToString();

            return volumedosecalculationalgorithm;
        }

        //String compare to test if same.
        //public static bool 

        //Gets the Dose Optimiser Algorithm (PO_16100).
        public static string GetOptimiserCalcAlgo(PlanSetup plan) //Working.
        {
            string Optcalculationalgorithm = plan.GetCalculationModel(VMS.TPS.Common.Model.Types.CalculationType.PhotonVMATOptimization).ToString();
            //string Optcalculationalgorithm = plan.GetCalculationModel(VMS.TPS.Common.Model.Types.CalculationType.PhotonVMATOptimization);
            //CalculationType.PhotonVMATOptimization

            return Optcalculationalgorithm;
        }

        //Gets the Portal Dose Calculation Algorithm (AAA for halcyons').
       // public static string GetPortalCalcAlgo(PlanSetup plan) //NOT WORKING.
       // {
       //     string Portalcalculationalgorithm = plan.PhotonCalculationOptions.
            

        //    return Portalcalculationalgorithm;
       // }

        //Gets Hetreogeneity correction value from the dose calc options.
        public static string GetHetreogeneity(PlanSetup plan) //WORKING.
        {
            string Hcorr = string.Empty;

            try
            {
                //string[] Hcorr = plan.PhotonCalculationOptions.ToArray().
                 Hcorr = plan.PhotonCalculationOptions.ToArray()[5].ToString();
                //string Optcalculationalgorithm = plan.GetCalculationModel(VMS.TPS.Common.Model.Types.CalculationType.PhotonVMATOptimization);
            }
            catch 
            { 
             Hcorr = string.Empty;
            }
            return Hcorr;
        }

        //// Check the photon angle optimiser is PO_1600;
        // public static string GetPO_opt(PlanSetup plan)
        // {
        //    string PO = string.Empty;

        //PO = plan.PhotonCalculationOptions.ToArray()[1].ToString(); //CalculationType.PhotonOptimization.GetType;
        //PO = dose. 

        //   return PO;

        // }

        //Gets Autofeathering value, (Should be PO_16100 has been set as optimisation). Assuming it's a Halcyon plan.
        public static string GetAutoFeather(PlanSetup plan) //WORKING.
        {
            string af = string.Empty;
            try
            {
                //string[] Hcorr = plan.PhotonCalculationOptions.ToArray().
                af = plan.GetCalculationOptions(GetOptimiserCalcAlgo(plan)).ToArray()[0].ToString();
            }
            catch
            {
                Console.WriteLine("Autofeathering check failed due to is empty or other error.");
            }
            
            return af;
        }

        //Get beam MU (3 decimal places or better.)
        public static double GetBeamMU(Beam beam)
        {
            double MUs = 0.000; 
            
            MUs = beam.Meterset.Value;

            return MUs;
        }

        //Search the beam calculation logs, for use in IMRT, when searching for the lost MU factor and the Max MU for Carriage Group.
        public static string GetBeamlogs(Beam beam, string searchstring) //Call for each beam, with searchstring of "LostMUfactor" or whatever. Also repeat process for "MaxMU for CarriageGroup".
        {
            string calclogs = string.Empty; 

            foreach(var s in beam.CalculationLogs)
            {
                var cakelogs = s.MessageLines;
                foreach (var i in cakelogs)
                {
                   // Console.WriteLine(i);
                    if (i.Contains(searchstring))
                        calclogs = i; //Really? You want to define 'i' as a string variable? But then what is sqrt(-1)?
                }
            }
            
            return calclogs;
        }


     
        public static bool beamAlign(PlanSetup plan)
        {
            var bms = plan.Beams.Where(q => !q.IsSetupField); 
            int cmax = bms.Count();//Get number of beams that aren't setup fields. //plan.Beams.Count(q => !q.IsSetupField);
            int counter = 0;
            bool beamalig = false;
            double[] b_x = new double[cmax];
            double[] b_y = new double[cmax];
            double[] b_z = new double[cmax];
                      
            
            foreach (var b in bms)
            {
                if (!b.IsSetupField) //if it's not a setup field, grab the x,y,z co-ords for each tmt. beam.
                {
                    b_x[counter] = b.IsocenterPosition.x;
                    b_y[counter] = b.IsocenterPosition.y;
                    b_z[counter] = b.IsocenterPosition.z;

                    Console.WriteLine($"The co-ords are: {b_x[counter].ToString()},{b_y[counter].ToString()},{b_z[counter].ToString()}"); //Note that co-ords are in Eclipse coordinates.
                }
                 counter++;
            }
            //Now compare the x,y,z co-ords for each tmt beam; s.t.:
                        
            return beamalig = b_x.Skip(1).All(x => x == b_x.First()) && b_y.Skip(1).All(y => y == b_y.First()) && b_z.Skip(1).All(z => z == b_z.First());
                        

        }

        public static bool AllBeamsSameTmtMachine(PlanSetup plan, string expectedtmtMachine)
        //e.g. expectedtmtMachine = "OHAL"
        {
            bool alltmtmachsame = false;
            var machs = plan.Beams.Where(x => x.TreatmentUnit.Id.Contains(expectedtmtMachine));
            int numel = machs.Count();
            return alltmtmachsame = numel == plan.Beams.Count(); // Number of elements with the expected treatment machine will equal the total number of beams in the plan if they are all the expectedtmtmachine.
        }

        public static bool AllBeamsDoseDynamicOrVMAT(PlanSetup plan)
        {
            bool allmlc = false;
            var bms = plan.Beams.Where(q => !q.IsSetupField);                     

            return allmlc = bms.All(x => x.MLCPlanType.ToString().Contains("DoseDynamic") | x.MLCPlanType.ToString().Contains("VMAT"));
        }


        ///Write to a different sheet.
        public static void WriteToCellOnSheet(string fileName, int sheetnumber, string row, int col, string value)
            {
      
                //     //  WorksheetPart worksheetPart = GetWorksheetPartByName(spreadsheet, "sheet1");

                //     spreadsheet.WorkbookPart
                //     WorksheetPart worksheetA = spreadsheet.WorkbookPart();
                // }
         Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = excelApp.Workbooks.Open(fileName, true);
        //Excel.Workbook workbook = excelApp.WorkbookOpen(fileName);
        //Excel.Worksheet worksheet1 = ((Excel.Worksheet)Application.ActiveWorkbook.Worksheets[1]);
        Excel.Worksheet worksheet = (workbook.Application.ActiveWorkbook.Worksheets[sheetnumber]); //Should give the active worksheet, given by sheetnumber; does not have 0 index. starts at 1.

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


        //Write a name-value pair array of values, to a different sheet:
        public static void WriteLotsOfValuesToCellsOnSheet(string fileName, int sheetnumber, string[] rows, int[] cols, string[] values)
        {

          
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(fileName, true);
          
            Excel.Worksheet worksheet = (workbook.Application.ActiveWorkbook.Worksheets[sheetnumber]); //Should give the active worksheet, given by sheetnumber; does not have 0 index. starts at 1.

            int incre = 0;
            //now write too it:
            foreach (string cell in rows)
            {   

                // worksheet.Cell(row, col).Value = value;
                string colstring = cols[incre].ToString();
                string rss = String.Join(string.Empty, cell, colstring); //Aren't in write format for function import.

                Excel.Range ra = worksheet.Range[rss]; //get the range.
                ra.Value2 = values[incre]; //Now write it to the desired.

                incre++;
            }

            workbook.Save();
            //Savin' & closing time.
            workbook.Close();

         
        }

       

        //Function to write check-boxes, for cases of Dubbo-plan-checks:

        //E.g.: 
        //WriteTickBox_Dubbo_ToCell(fn, "A", 1, "false", ""); //Make it "TRUE" for tick.
        //WriteTickBox_Dubbo_ToCell(fn, "C", 20, "TRUE", "\u2071"); //Make it "TRUE" for tick.

        public static void WriteTickBox_Dubbo_ToCell(string fileName, string row, int col, string value, string commentcaption) //Make it "TRUE" for tick. "False for anything else." 
                                                                                                                  //Commentcaption will be normal font size, but won't save nor display properly if the cell sizing is smaller than the font. Also, unable to edit this caption - without removing checkbox.
        {

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(fileName, true);

            Excel.Worksheet worksheet = (workbook.Application.ActiveWorkbook.Worksheets[1]); //Should give first active worksheet.

            //now write too it:
            // worksheet.Cell(row, col).Value = value;
            string colstring = col.ToString();
            string rss = String.Join(string.Empty, row, colstring); //Aren't in write format for function import.

            Excel.Range cell = worksheet.Range[rss]; //get the range.
                                                     //ra.Value2 = value; //Now write it to the desired.

            OLEObjects objs = worksheet.OLEObjects();
            OLEObject obj = objs.Add("Forms.CheckBox.1", System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, cell.Left + 1, cell.Top + 1, cell.Width - 2, cell.Height - 2);
            obj.Object.Caption = commentcaption; //Caption here.
            if (value == "TRUE")
            {
                obj.Object.value = true;
            }
            else
            {
                obj.Object.value = false;
            }

            workbook.Save();

            workbook.Close();

        }


        //Function to find first empty row of a spreadsheet:
        public static int GetEmptyCellRow(string filename, int sheetnumber)
        {
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(filename, true);

            Excel.Worksheet xlWorkSheet = (workbook.Application.ActiveWorkbook.Worksheets[sheetnumber]); //Should give the active worksheet, given by sheetnumber; does not have 0 index. starts at 1.
            var xlRange = (Excel.Range)xlWorkSheet.Cells[xlWorkSheet.Rows.Count, 1];
            int lastRow = (int)xlRange.get_End(Excel.XlDirection.xlUp).Row;
            int newRow = lastRow + 1;
            
            workbook.Save();
            Thread.Sleep(500);
            workbook.Close();
            Thread.Sleep(500); //wait for 1 sec total. O/w workbook not closed in time for next writing.
            return newRow;


        }






        //Function to write name-value array pairs to lots of check-box cells:
        //Function to write check-boxes, for cases of Dubbo-plan-checks:

        //E.g.: 
        //WriteTickBox_Dubbo_ToCell(fn, "A", 1, "false", ""); //Make it "TRUE" for tick.
        //WriteTickBox_Dubbo_ToCell(fn, "C", 20, "TRUE", "\u2071"); //Make it "TRUE" for tick.

        //NOT used, can write check-boxes to dubbo checklist. Refer to draft version in T-box, as is more polished.

        public static void WriteLotsOfTickBoxs_Dubbo_ToCells(string fileName, string[] rows, int[] cols, string[] values, string[] commentcaptions) //Make it "TRUE" for tick. "False for anything else." 
                                                                                                                  //Commentcaption will be normal font size, but won't save nor display properly if the cell sizing is smaller than the font. Also, unable to edit this caption - without removing checkbox.
        {

            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(fileName, true);

            Excel.Worksheet worksheet = (workbook.Application.ActiveWorkbook.Worksheets[1]); //Should give first active worksheet.

            //now write too it:

            int incre = 0;
            //now write too it:
            foreach (string ceL in rows)
            {

                
                string colstring = cols[incre].ToString(); //already string but w/e.
                string row = rows[incre].ToString(); 
                string value = values[incre].ToString();
                string commentcaption = commentcaptions[incre].ToString();

                string rss = String.Join(string.Empty, row, colstring); //Aren't in write format for function import.
                Excel.Range cell = worksheet.Range[rss]; //get the range.
                //ra.Value2 = values[incre]; //Now write it to the desired.

                OLEObjects objs = worksheet.OLEObjects();
                OLEObject obj = objs.Add("Forms.CheckBox.1", System.Reflection.Missing.Value, System.Reflection.Missing.Value, false, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value, cell.Left + 32, cell.Top + 1, cell.Width - 2, cell.Height - 2);
                obj.Object.Caption = commentcaption; //Caption here.
                if (value == "TRUE")
                {
                    obj.Object.value = true;
                }
                else
                {
                    obj.Object.value = false;
                }

                incre++;
            }                         

            workbook.Save();

            workbook.Close();

        }




    }
}

