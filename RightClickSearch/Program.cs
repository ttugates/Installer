using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace RightClickSearch
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            LoadDLLsFromEmbeddedResources();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Handle Uninstall Request First.
            if (args.Length != 0)
            {
                if (args[0].ToString().ToLower().Contains(@"/uninstallprompt") == true)
                {
                    InstallerHelper.UninstallProgram(InstallerSettings());
                }
            }
                        
            else if (InstallerHelper.InstallOrContinueUpdating(InstallerSettings()) == false)
            {
                // Run Check for updates on fire-and-forget thread.
                Task.Run(() => InstallerHelper.CheckForUpdates());

                if (args.Length != 0)
                {                   
                    Application.Run(new SearchForm(args[0]));
                }
                else
                {
                    Application.Run(new SearchForm());
                }
            }
        }

        private static void LoadDLLsFromEmbeddedResources()
        {
            // Load DLLs as Embeded Resources.
            string resourceToLoad = "RightClickSearch.Assemblies.Newtonsoft.Json.dll";
            EmbeddedAssemblyHandler.Load(resourceToLoad, resourceToLoad);
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(FetchAssembliesFromMemory);
        }

        private static Assembly FetchAssembliesFromMemory(object sender, ResolveEventArgs args)
        {                  
            return EmbeddedAssemblyHandler.Get(args.Name);
        }

        private static InstallerSettingsContainer InstallerSettings()
        {
            return new InstallerSettingsContainer()
            {
                // Tip-Get GUID by right clicking Project in VS.
                guidText = "8e1bb901-6955-4512-8da4-530f2643b28d",
                assemblyName = System.Diagnostics.Process.GetCurrentProcess().ProcessName,
                ReleasesURL = @"https://api.github.com/repos/ttugates/Installer/releases",
                DisplayName = "PowerSearch By MDS",
                Publisher = "MDS",
                DisplayIcon = "", //???
                URLInfoAbout = "www.google.com",
                Contact = "michaelstramel@gmail.com",
            };
        }
        
    }    
}


// TODO - Get rid of, "This program might not have uninstalled correctly"
// http://stackoverflow.com/questions/898220/how-to-prevent-this-program-might-not-have-installed-correctly-messages-on-vis