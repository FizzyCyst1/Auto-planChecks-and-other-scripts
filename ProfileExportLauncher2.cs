using System.Windows;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using System.Diagnostics;
using System.IO;
using System;

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context)
        {
            try
            {
                /**** Throw exception if the patient object is null ****/
                if (context.Patient == null) throw new Exception("Please open patient before running script");
                /**** Throw exception if no structureSets are present ****/
                //if (context.Patient.StructureSets.Count() == 0) throw new Exception("No structure sets are attached to this patient");
                /**** Set the launch path of the executable ****/
                string launcherPath = @"\\ORDCARI-MFS901\va_data$\PublishedScripts\StandAlone\BeamProfiles\";
                /**** remove domain name from user ID ****/
                string currentUser = context.CurrentUser.Id.Trim().Replace(@"nswhealth\", "");
                /**** Set the exe to run script ****/
                string esapiStandaloneExecutable = @"SunDoseProfileExport2.exe";
                /**** Check the Course, Plan and StructrueSet ID for spaces ****/
                if (context.Course.Id.Contains(" ")) throw new Exception("Course ID contains a space");
                if (context.ExternalPlanSetup != null)
                {
                    if (context.ExternalPlanSetup.Id.Contains(" ")) throw new Exception("Plan ID contains a space");
                }
                if (context.StructureSet.Id.Contains(" ")) throw new Exception("StructureSet ID contains a space");
                /**** Set the argument string by join the components of the Script Context based off whether plan is added ****/
                string arguments = context.PlanSetup == null
                                    ? string.Format("{0};{1};{2};{3}", context.Patient.Id, context.Course.Id, context.StructureSet.Id, currentUser)
                                    : string.Format("{0};{1};{2};{3};{4}", context.Patient.Id, context.Course.Id, context.PlanSetup.Id, context.StructureSet.Id, currentUser);
                Process.Start(Path.Combine(launcherPath, esapiStandaloneExecutable), arguments);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to start application due to following error: " + ex);
            }
        }

        public string GetSourceFilePath([CallerFilePath] string sourceFilePath = "")
        {
            return sourceFilePath;
        }
    }
}
