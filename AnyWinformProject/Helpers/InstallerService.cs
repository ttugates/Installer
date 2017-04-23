/* InstallerService has a dependency on:
 *  JSONService which is System.Text.Json
 *  InstallerSettingsContainer.cs
 *  JSONGitModels.cs
 *  
 *  When building, go to Porperties of Consuming App, Assembly Information and Set File Version.
 *  When publishing to GitHub Releases, set the Tag Version to match the File Version.  
 */

using InstallerService.Models;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading.Tasks;

namespace InstallerService
{
    public static class InstallerService
    {
        #region Gloabls

        private static Assembly entryAssy = Assembly.GetEntryAssembly();
        private static string _executingAssyPathComplete = entryAssy.Location;
        private static string _executingAssyPathOnly = Path.GetDirectoryName(entryAssy.Location);
        private static string _executingAssyVersion = entryAssy.GetName().Version.ToString();
        private static string _executingAssyNameOnly = Path.GetFileNameWithoutExtension(entryAssy.Location);
        private static string _executingAssyNameAndExtension = Path.GetFileName(entryAssy.Location);
        private static string _installPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), _executingAssyNameOnly);
        private static string _installPathComplete = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), _executingAssyNameOnly) + "\\" + _executingAssyNameAndExtension;
        private static string _pendingInstallPath = Path.Combine(_installPath, "PendingInstall");
        private static string _pendingInstallPathComplete = _pendingInstallPath + "\\" + _executingAssyNameAndExtension;
        private static string _pid = Process.GetCurrentProcess().Id.ToString();
        private static string _keyName = @"HKEY_CURRENT_USER\SOFTWARE\" + _executingAssyNameOnly;
        private static string _isInstalledValueName = "IsInstalled";
        private static string _installerPIDValueName = "InstalledPID";
        private static string _currentVersion = "CurrentVersion";
        private static string _pendingUpdateValueName = "PendingUpdate";
        private static string _installerLocation = "InstallerPathFull";
        private static string _displayIcon = Path.Combine(_installPath, _executingAssyNameAndExtension);

        private static string _guidText;
        private static string _releasesURL;
        private static string _displayName;
        private static string _publisher;
        private static string _uRLInfoAbout;
        private static string _contact;


        #endregion Gloabls

        #region Public Methods

        public static bool InstallOrContinueUpdating(InstallerSettingsContainer settings)
        {
            InstallerEventPipeline(null, new InstallerMessageEventArgs(_pid, "InstallOrContinueUpdating Entered", Verbosity.Detailed));
            LoadSettings(settings);

            string isInstalled = (string)(Registry.GetValue(_keyName, _isInstalledValueName, "False"));
            string updatesAvail = (string)(Registry.GetValue(_keyName, _pendingUpdateValueName, "False"));

            // ** TODO - Figure this out..  I have seen where when the registry settings do not exist,
            // Registry.GetValue returns null instead of the defined "False" set as the 3rd Parameter.
            //  Oddly, I can use Regedit and manually make one of the keys.  And Both will show as False.
            //  The following is a crudge until I resolve the cause.  Likely to do with x32 vs x64 etc.
            if (isInstalled == null)
            {
                isInstalled = "False";
            }
            if (updatesAvail == null)
            {
                updatesAvail = "False";
            }

            // If First Install
            if (isInstalled == "False" && updatesAvail == "False")
            {
                InstallerEventPipeline(null, new InstallerMessageEventArgs(_pid, "Is first install", Verbosity.Detailed));
                InstallProgram();
                return true;
            }
            // Else if Pending Update
            else if (isInstalled == "False" && updatesAvail == "True")
            {
                InstallerEventPipeline(null, new InstallerMessageEventArgs(_pid, "Is Pending Update.", Verbosity.Detailed));
                LoadPendingUpdateInstaller();
                return true;
            }
            // Else Installing Update
            else if (isInstalled == "True" && updatesAvail == "True")
            {
                InstallerEventPipeline(null, new InstallerMessageEventArgs(_pid, "Installing Update", Verbosity.Detailed));
                FinalizePendingUpdateInstall();
                return true;
            }
            // Installed no Pending Updates
            else
            {
                // Check if this is first run after install, if so, close installer process.
                KillLastProcess();

                HandleInstallerFile();

                // Remove Pending Install Folder if exists.
                if (Directory.Exists(_pendingInstallPath) == true)
                {
                    Directory.Delete(_pendingInstallPath, true);
                }

                // Run Check for updates on fire-and-forget thread.
                Task.Run(() => CheckForUpdates());

                return false;
            }
        }

        public static async Task CheckForUpdates(bool autoUpdate = true)
        {
            InstallerEventPipeline(null, new InstallerMessageEventArgs(_pid, "Begin Check for Updates", Verbosity.Detailed));

            string json;

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent", "Anything");
                json = await client.GetStringAsync(_releasesURL);
            }

            JsonParser parser = new JsonParser();
            List<GitReleases> gitReleases = parser.Parse<List<GitReleases>>(json);
            Version currentReleaseVersion = new Version(gitReleases.OrderByDescending(x => x.tag_name).Select(x => x.tag_name).FirstOrDefault());
            Version runningVersion = new Version(_executingAssyVersion);

            if (currentReleaseVersion > runningVersion)
            {
                if (autoUpdate == true)
                {
                    string browser_download_url = gitReleases.OrderByDescending(x => x.tag_name).Select(x => x.assets).FirstOrDefault().Select(x => x.browser_download_url).FirstOrDefault();
                    string fileNaame = gitReleases.OrderByDescending(x => x.tag_name).Select(x => x.assets).FirstOrDefault().Select(x => x.name).FirstOrDefault();
                    await InitiatePendingUpdate(browser_download_url, fileNaame);
                }
            }
            InstallerEventPipeline(null, new InstallerMessageEventArgs(_pid, "End Check for Updates", Verbosity.Detailed));
        }

        public static void UninstallProgram(InstallerSettingsContainer settings)
        {
            LoadSettings(settings);

            // Close all other running Processes
            KillAllOtherProcesses();

            // Remove Registry Entries
            try
            {
                using (RegistryKey localKey32 = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry32).OpenSubKey(@"SOFTWARE", true))
                {
                    localKey32.DeleteSubKeyTree(_executingAssyNameOnly);
                }
            }
            catch { }

            try
            {
                using (RegistryKey parent = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", true))
                {
                    parent.DeleteSubKeyTree(_executingAssyNameOnly);
                }
            }
            catch { }

            // Remove Desktop Shortcut
            DeleteDesktopShortcut();

            // TODO - Delete Program
            // Need to solve issue of EXE cant delete itself.
            //  Perhaps - Using Task Scheduler....
            //  Or Registry -

            InstallerEventPipeline(null, new InstallerMessageEventArgs("Uninstall Complete", _displayName + " was successfully uninstalled." 
                + Environment.NewLine + Environment.NewLine + "Please disregard any following message asking if Uninstall Was Successful." + Environment.NewLine
                + "An upcomming release of this installer will address this issue. :)", Verbosity.Normal));
            Environment.Exit(1);
        }

        #endregion Public Methods

        #region Private Methods

        private static async Task InitiatePendingUpdate(string browser_download_url, string fileName)
        {
            InstallerEventPipeline(null, new InstallerMessageEventArgs("Updates Found", "Updates Found, downloading.", Verbosity.Normal));
            Stream stream;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var contentBytes = await client.GetByteArrayAsync(browser_download_url);
                    stream = new MemoryStream(contentBytes);
                }

                Directory.CreateDirectory(_pendingInstallPath);

                string downloadFileLocation = Path.Combine(_pendingInstallPath, fileName);

                using (FileStream fs = new FileStream(downloadFileLocation, FileMode.Create))
                {
                    stream.CopyTo(fs);
                }

                stream.Close();

                Registry.SetValue(_keyName, _isInstalledValueName, false);
                Registry.SetValue(_keyName, _pendingUpdateValueName, true);
                InstallerEventPipeline(null, new InstallerMessageEventArgs("Finished Downloading", "Finished Dowloading and Staged for Install.", Verbosity.Normal));
            }
            catch { }
        }

        private static void LoadPendingUpdateInstaller()
        {
            try
            {
                int pid = Process.GetCurrentProcess().Id;

                // Set Registry to FinalizePendingUpdateInstall
                Registry.SetValue(_keyName, _pendingUpdateValueName, true);
                Registry.SetValue(_keyName, _isInstalledValueName, true);
                Registry.SetValue(_keyName, _installerPIDValueName, pid);
                Registry.SetValue(_keyName, _currentVersion, _executingAssyVersion);

                // Start new PendingInstall process.
                Process.Start(_pendingInstallPathComplete);
            }
            catch (Exception ex)
            {
                InstallerEventPipeline(null, new InstallerMessageEventArgs("Error", "Error Updating:" + Environment.NewLine + ex.Message, Verbosity.DetailedWithErrors));
            }
        }

        private static void FinalizePendingUpdateInstall()
        {
            try
            {
                KillLastProcess();

                int pid = Process.GetCurrentProcess().Id;

                Directory.CreateDirectory(_installPath);

                // Wait for calling instance to close / release lock on application.exe
                int i = 0;
                while (IsFileLocked(_installPathComplete) == true || i > 45)
                {
                    System.Threading.Thread.Sleep(333);
                    i++;
                }

                // Copy to final destination.
                File.Copy(_pendingInstallPathComplete, _installPathComplete, true);

                // Set Registry as Installed.
                Registry.SetValue(_keyName, _isInstalledValueName, true);
                Registry.SetValue(_keyName, _pendingUpdateValueName, false);
                Registry.SetValue(_keyName, _installerPIDValueName, pid);

                // Update Add/Remove Programs Uninstall Info
                CreateUninstaller();

                // Start new installed process.
                Process.Start(_installPathComplete);
            }
            catch (Exception ex)
            {
                InstallerEventPipeline(null, new InstallerMessageEventArgs("Error", "Error Updating:" + Environment.NewLine + ex.Message, Verbosity.DetailedWithErrors));
            }
        }

        private static void InstallProgram()
        {
            try
            {
                int pid = Process.GetCurrentProcess().Id;

                Directory.CreateDirectory(_installPath);

                int i = 0;
                while (IsFileLocked(_installPathComplete) == true || i > 60)
                {
                    System.Threading.Thread.Sleep(250);
                    i++;
                }

                File.Copy(_executingAssyPathComplete, _installPathComplete, true);

                Registry.SetValue(_keyName, _isInstalledValueName, true);
                Registry.SetValue(_keyName, _currentVersion, _executingAssyVersion);
                Registry.SetValue(_keyName, _installerPIDValueName, pid);
                Registry.SetValue(_keyName, _installerLocation, _executingAssyPathComplete);

                CreateUninstaller();

                // Create Desktop Shortcut
                CreateDesktopShortcut(_installPathComplete, _executingAssyNameOnly);

                InstallerEventPipeline(null, new InstallerMessageEventArgs(_executingAssyNameOnly + " is installed. ", "Successfully Installed." + Environment.NewLine + Environment.NewLine +
                    "Uninstall via Add/Remove Programs." + Environment.NewLine + Environment.NewLine +
                    "A shortcut has been placed on your desktop." + Environment.NewLine + Environment.NewLine + 
                    "Your installer has been placed in your Downloads folder.", Verbosity.Normal));

                if (Process.Start(_installPathComplete) == null)
                {
                    int lastError = Marshal.GetLastWin32Error();
                }
            }
            catch (Exception ex)
            {
                InstallerEventPipeline(null, new InstallerMessageEventArgs("Error", "Error Updating:" + Environment.NewLine + ex.Message, Verbosity.DetailedWithErrors));
            }
        }

        private static void CreateUninstaller()
        {
            string ApplicationVersion = _executingAssyVersion;
            string DisplayVersion = ApplicationVersion;
            string InstallDate = DateTime.Now.ToString("yyyyMMdd");
            string UninstallString = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + _executingAssyNameOnly + "\\" + _executingAssyNameOnly + ".exe " + @" /uninstallprompt";

            using (RegistryKey parent = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", true))
            {
                if (parent == null)
                {
                    throw new Exception("Uninstall registry key not found.");
                }
                try
                {
                    RegistryKey key = null;
                    
                    try
                    {
                        key = parent.OpenSubKey(_executingAssyNameOnly, true) ?? parent.CreateSubKey(_executingAssyNameOnly);

                        if (key == null)
                        {
                            throw new Exception(String.Format("Unable to create uninstaller."));
                        }

                        
                        try
                        {
                            long installSize = Directory.GetFiles(_installPath, "*", SearchOption.AllDirectories).Sum(t => (new FileInfo(t).Length));
                            key.SetValue("EstimatedSize", installSize/1024, RegistryValueKind.DWord);
                        }
                        catch { }

                        key.SetValue("DisplayName", _displayName);
                        key.SetValue("ApplicationVersion", ApplicationVersion);
                        key.SetValue("Publisher", _publisher);
                        key.SetValue("DisplayVersion", DisplayVersion);
                        key.SetValue("InstallDate", InstallDate);
                        key.SetValue("UninstallString", UninstallString);
                        key.SetValue("DisplayIcon", _installPathComplete);
                        key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                        key.SetValue("NoRepair", 1, RegistryValueKind.DWord);


                        /*
                        InstallLocation (string)    - Installation directory ($INSTDIR)
                        DisplayIcon (string)        - Path, filename and index of of the icon that will be displayed next to your application name
                        Publisher (string)          - (Company) name of the publisher
                        ModifyPath (string)         - Path and filename of the application modify program
                        InstallSource (string)      - Location where the application was installed from
                        ProductID (string)          - Product ID of the application
                        Readme (string)             - Path (File or URL) to readme information
                        RegOwner (string)           - Registered owner of the application
                        RegCompany (string)         - Registered company of the application
                        HelpLink (string)           - Link to the support website
                        HelpTelephone (string)      - Telephone number for support
                        URLUpdateInfo (string)      - Link to the website for application updates
                        URLInfoAbout (string)       - Link to the application home page
                        DisplayVersion (string)     - Displayed version of the application
                        VersionMajor (DWORD)        - Major version number of the application
                        VersionMinor (DWORD)        - Minor version number of the application
                        NoModify (DWORD)            - 1 if uninstaller has no option to modify the installed application
                        NoRepair (DWORD)            - 1 if the uninstaller has no option to repair the installation
                        SystemComponent (DWORD)     - Set 1 to prevents display of the application in the Programs List of the Add/Remove Programs in the Control Panel.
                        EstimatedSize (DWORD)       - The size of the installed files (in KB)
                        Comments (string)           - A comment describing the installer package
                        */
                    }
                    finally
                    {
                        if (key != null)
                        {
                            key.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    InstallerEventPipeline(null, new InstallerMessageEventArgs("Error creating Uninstaller", "An error occurred writing uninstall information to the registry.  The service is fully installed but can only be uninstalled manually through the command line." + Environment.NewLine + ex.Message, Verbosity.DetailedWithErrors));
                }
            }
        }

        private static bool IsFileLocked(string fileName)
        {
            FileInfo file = new FileInfo(fileName);
            FileStream stream = null;

            if (File.Exists(fileName) == false)
            {
                return false;
            }

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }

        private static void KillLastProcess()
        {
            int? pidOfInstaller = Registry.GetValue(_keyName, _installerPIDValueName, null) as int?;
            if (pidOfInstaller != null)
            {
                try
                {
                    Process p = Process.GetProcessById((int)pidOfInstaller);
                    if (p.ProcessName.ToLower().Contains((_executingAssyNameOnly).ToLower()) == true)
                    {
                        p.Kill();
                    }
                }
                catch
                {
                }

                try
                {   // http://www.rhyous.com/2011/01/24/how-read-the-64-bit-registry-from-a-32-bit-application-or-vice-versa/
                    // TODO - If make this a 64 bit app, this will need to be dynamic.
                    RegistryKey localKey32 = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry32);
                    localKey32 = localKey32.OpenSubKey(@"SOFTWARE\" + _executingAssyNameOnly, true);
                    localKey32.DeleteValue(_installerPIDValueName);
                }
                catch (Exception ex)
                {
                    InstallerEventPipeline(null, new InstallerMessageEventArgs("Error", "Error Deleting PID from registry:" + Environment.NewLine + ex.Message, Verbosity.DetailedWithErrors));
                }

            }
        }

        private static void KillAllOtherProcesses()
        {
            var allProcceses = Process.GetProcesses();
            int currentPID = Process.GetCurrentProcess().Id;
            foreach (var process in allProcceses)
            {
                try
                {
                    if (process.ProcessName == _executingAssyNameOnly)
                    {
                        if (process.Id != currentPID)
                        {
                            process.Kill();
                        }
                    }
                }
                catch { }
            }
        }

        private static void CreateDesktopShortcut(string appName, string executableName)
        {
            string[] Lines = {
                @"set WshShell = WScript.CreateObject(""WScript.Shell"")",
                @"strDesktop = WshShell.SpecialFolders(""Desktop"")",
                @"set oShellLink = WshShell.CreateShortcut(strDesktop & ""\" + _executingAssyNameOnly + @".lnk"")",
                @"oShellLink.TargetPath = ""%appdata%\" + _executingAssyNameOnly + @"\" + _executingAssyNameAndExtension + "\"",
                @"oShellLink.WindowStyle = 1",
                @"oShellLink.Description = """ + _executingAssyNameOnly + "\"",
                @"oShellLink.WorkingDirectory = strDesktop",
                @"oShellLink.Save()"
            };
            File.WriteAllLines("createShortcut.vbs", Lines);
            Process P = Process.Start("createShortcut.vbs");
            P.WaitForExit(int.MaxValue);
            File.Delete("createShortcut.vbs");
        }

        private static void DeleteDesktopShortcut()
        {
            try
            {
                var shortcutPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                shortcutPath = Path.Combine(shortcutPath, _executingAssyNameOnly + ".lnk");

                if (File.Exists(shortcutPath) == true)
                {
                    File.Delete(shortcutPath);
                }
            }
            catch { }
        }

        private static void LoadSettings(InstallerSettingsContainer settings)
        {
            _guidText =     settings.GUIDText;
            _releasesURL =  settings.ReleasesURL;
            _displayName =  settings.DisplayName;
            _publisher =    settings.Publisher;
            _uRLInfoAbout = settings.URLInfoAbout;
            _contact =      settings.Contact;
        }

        private static void VerifyRunAsInstalled()
        {
            // Handle clicking of old installer instead of shortcut or installed exe.




            // If this is not running from installation folder.
            if (_executingAssyPathOnly != _installPath)
            {
                Version thisExeVersion = Version.Parse(_executingAssyVersion);
                Version installedVersion = Version.Parse((string)Registry.GetValue(_keyName, _currentVersion, "0.0.0.0"));
                // If this is not a newer version of installed exe.
                if (installedVersion >= thisExeVersion)
                {
                    InstallerEventPipeline(null, new InstallerMessageEventArgs(_executingAssyNameOnly + ", Version: " + installedVersion.ToString(), _executingAssyNameOnly + ", Version: " + installedVersion.ToString() + Environment.NewLine + "is already installed." + Environment.NewLine + Environment.NewLine +
                        "Use Add/Remove Programs to uninstall." + Environment.NewLine + Environment.NewLine +
                        "Or use the shortcut on your Desktop to run the software.", Verbosity.Normal));

                    Environment.Exit(1);
                }
            }
        }

        private static void HandleInstallerFile()
        {
            InstallerEventPipeline(null, new InstallerMessageEventArgs("Entering Handle Installer", "HandleInstallerFile() called", Verbosity.Detailed));
            try
            {
                // Does registry value indicating Installer needs to be moved to Downloads exist?  
                string source = (string)Registry.GetValue(_keyName, _installerLocation, null);                
                if (source != null)
                {
                    string destination = Path.Combine(GetDownloadsFolderPath(), _displayName + " Installer");
                    
                    InstallerEventPipeline(null, new InstallerMessageEventArgs("Attempting to move installer to:", destination, Verbosity.Detailed));

                    try
                    {
                        Directory.CreateDirectory(destination);
                    }
                    catch
                    {
                        InstallerEventPipeline(null, new InstallerMessageEventArgs("Error", "Unable to create Downloads Directory.", Verbosity.DetailedWithErrors));
                    }
                    try
                    {
                        destination = Path.Combine(destination, Path.GetFileName(source));
                        if (File.Exists(destination) == true)
                        {
                            File.Delete(destination);
                        }
                        File.Move(source, destination);
                        
                        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(_keyName, true))
                        {
                            if (key != null)
                            { 
                                key.DeleteValue(_installerLocation);
                            }
                        }
                        Registry.SetValue(_keyName, _installerLocation, null);
                        
                    }
                    catch (Exception ex)
                    {
                        InstallerEventPipeline(null, new InstallerMessageEventArgs("Error", "Unable to move Installer to Downloads Folder." + Environment.NewLine + ex.Message + Environment.NewLine + ex.StackTrace, Verbosity.DetailedWithErrors));
                    }

                }
            }
            catch (Exception ex)
            {
                InstallerEventPipeline(null, new InstallerMessageEventArgs("Error", "Error in HandleInstallerFile()." + Environment.NewLine + ex.Message, Verbosity.DetailedWithErrors));
            }
        }

        private static string GetDownloadsFolderPath()
        {
            // http://stackoverflow.com/questions/10667012/getting-downloads-folder-in-c
            string destination = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "{374DE290-123F-4565-9164-39C4925E467B}", String.Empty).ToString();
            return destination;
        }

        #endregion Private Methods

        #region Events

        public static event EventHandler<InstallerMessageEventArgs> InstallerEventPipeline;

        public class InstallerMessageEventArgs : EventArgs
        {
            public InstallerMessageEventArgs(string _title, string _message, Verbosity _howVerbose)
            {
                message = _message;
                title = _title;
                howVerbose = _howVerbose;
            }

            public string message { get; set; }
            public string title { get; set; }
            public Verbosity howVerbose { get; set; }
        }

        public enum Verbosity
        {
            Normal = 0,
            Detailed = 1,
            DetailedWithErrors = 2,            
        }

        #endregion Events
    }
}
