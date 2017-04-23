using InstallerService.Models;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using static InstallerService.InstallerService;

namespace AnyWinformProject
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            // Use the following when embeding Dlls/Assemblies as Resources
            // Rember to right click Reference and set to Copy Local = false    
            // Also, right click a copy of Regerence Dll and set to Embeded Resource
            // DLLService.LoadDLLsFromEmbeddedResources(new List<string>() { "ProgramName.Folders.NameOfDll.dll" }  );

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Subcribe to installer events.
            InstallerEventPipeline += InstallerEventHandler;

            // Handle Any Uninstall Request.
            if (args.Length != 0)
            {
                if (args[0].ToString().ToLower().Contains(@"/uninstallprompt") == true)
                {
                    UninstallProgram(InstallerSettings());
                }
            }
            else if (InstallOrContinueUpdating(InstallerSettings()) == false)
            {
                Application.Run(new Form1());
            }
        }

        private static InstallerSettingsContainer InstallerSettings()
        {
            return new InstallerSettingsContainer()
            {
                // Tip-Get GUID by right clicking Project in VS.
                GUIDText = "b86d0142-e97e-4af4-923c-c38ce9d0590b",
                ReleasesURL = @"https://api.github.com/repos/ttugates/Installer/releases",
                DisplayName = "Winform Template",
                Publisher = "MDS",
                URLInfoAbout = "https://github.com/ttugates",
                Contact = "michaelstramel@gmail.com",
            };
        }

        private static void InstallerEventHandler(object sender, InstallerMessageEventArgs e)
        {
            // Set Verbosity of Messages here.
            Verbosity verbisityToShow = Verbosity.Normal;

            if ((int)e.howVerbose <= (int)verbisityToShow)
            {
                MessageBox.Show(e.message, e.title);
            }
        }
    }
}
