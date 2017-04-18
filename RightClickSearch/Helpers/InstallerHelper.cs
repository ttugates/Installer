using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace RightClickSearch
{
    public static class InstallerHelper
    {
        #region Gloabls
        private static string _assemblyName;
        private static string _guidText;
        private static string _releasesURL;
        private static string _displayName;
        private static string _publisher;
        private static string _displayIcon;
        private static string _uRLInfoAbout;
        private static string _contact;
        private static string _keyName;
        private static string _isInstalledValueName;
        private static string _installerPIDValueName;
        private static string _pendingUpdateValueName;
        private static string _installFolder;
        private static string _pendingInstallFolder;
        private static string _currentVersion;
        #endregion

        #region Public Methods
        public static bool InstallOrContinueUpdating(InstallerSettingsContainer settings)
        {
            LoadSettings(settings);
            // VerifyRunAsInstalled();

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
                InstallProgram();
                return true;
            }
            // Else if Pending Update
            else if (isInstalled == "False" && updatesAvail == "True")
            {
                LoadPendingUpdateInstaller();
                return true;
            }
            // Else Installing Update
            else if (isInstalled == "True" && updatesAvail == "True")
            {
                FinalizePendingUpdateInstall();
                return true;
            }
            // Installed no Pending Updates
            else
            {
                // Check if this is first run after install, if so, close installer process.
                KillLastProcess();

                // Remove Pending Install Folder if exists.
                if (Directory.Exists(_pendingInstallFolder) == true)
                {
                    Directory.Delete(_pendingInstallFolder, true);
                }
                            
                return false;
            }

        }

        public static async Task CheckForUpdates(bool autoUpdate = true)
        {
            string json;

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("User-Agent", "Anything");
                json = await client.GetStringAsync(_releasesURL);
            }

            List<GitReleases> gitReleases = JsonConvert.DeserializeObject<List<GitReleases>>(json);
            Version currentReleaseVersion = new Version(gitReleases.OrderByDescending(x => x.tag_name).Select(x => x.tag_name).FirstOrDefault());
            Version runningVersion = new Version(Application.ProductVersion);

            if (currentReleaseVersion > runningVersion)
            {
                if (autoUpdate == true)
                {
                    string browser_download_url = gitReleases.OrderByDescending(x => x.tag_name).Select(x => x.assets).FirstOrDefault().Select(x => x.browser_download_url).FirstOrDefault();
                    string fileNaame = gitReleases.OrderByDescending(x => x.tag_name).Select(x => x.assets).FirstOrDefault().Select(x => x.name).FirstOrDefault();
                    await InitiatePendingUpdate(browser_download_url, fileNaame);
                }
            }
        }
        
        public static void UninstallProgram(InstallerSettingsContainer settings)
        {
            LoadSettings(settings);

            DialogResult dialogResult = MessageBox.Show("Uninstall " + _assemblyName + "?", "Uninstall?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                // Close all other running Processes
                KillAllOtherProcesses();

                // Remove Registry Entries
                try
                {
                    using (RegistryKey localKey32 = RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, RegistryView.Registry32).OpenSubKey(@"SOFTWARE", true))
                    {
                        localKey32.DeleteSubKeyTree(_assemblyName);
                    }
                }
                catch { }


                try
                {
                    using (RegistryKey parent = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", true))
                    {
                        parent.DeleteSubKeyTree(_assemblyName);
                    }
                }
                catch { }

                // Remove Desktop Shortcut
                DeleteDesktopShortcut();

                // TODO - Delete Program
                // Need to solve issue of EXE cant delete itself.
                //  Perhaps - Using Task Scheduler....
                //  Or Registry - 


                MessageBox.Show("Uninstall Successful.");
                Application.Exit();
            }
            else if (dialogResult == DialogResult.No)
            {
                MessageBox.Show("Uninstall Aborted.");
                Application.Exit();
            }

           
        }
        #endregion

        #region Private Methods
        private static async Task InitiatePendingUpdate(string browser_download_url, string fileName)
        {
            Stream stream;
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    var contentBytes = await client.GetByteArrayAsync(browser_download_url);
                    stream = new MemoryStream(contentBytes);
                }

                Directory.CreateDirectory(_pendingInstallFolder);

                string downloadFileLocation = Path.Combine(_pendingInstallFolder, fileName);

                using (FileStream fs = new FileStream(downloadFileLocation, FileMode.Create))
                {
                    stream.CopyTo(fs);
                }

                stream.Close();

                Registry.SetValue(_keyName, _isInstalledValueName, false);
                Registry.SetValue(_keyName, _pendingUpdateValueName, true);
            }
            catch { }
        }

        private static void LoadPendingUpdateInstaller()
        {
            try
            {
                string exePath = Application.ExecutablePath;
                var pendingPath = _pendingInstallFolder + "\\" + Path.GetFileName(exePath);
                int pid = Process.GetCurrentProcess().Id;

                // Set Registry to FinalizePendingUpdateInstall
                Registry.SetValue(_keyName, _pendingUpdateValueName, true);
                Registry.SetValue(_keyName, _isInstalledValueName, true);
                Registry.SetValue(_keyName, _installerPIDValueName, pid);
                Registry.SetValue(_keyName, _currentVersion, Application.ProductVersion);

                // Start new PendingInstall process.
                Process.Start(pendingPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Updating:" + Environment.NewLine + ex.Message);
            }
        }

        private static void FinalizePendingUpdateInstall()
        {
            try
            {
                KillLastProcess();

                // Get PendingInstaller Path & PID
                string exePath = Application.ExecutablePath;
                int pid = Process.GetCurrentProcess().Id;
                string destPath = _installFolder;

                Directory.CreateDirectory(destPath);
                destPath += "\\" + Path.GetFileName(exePath);

                // Wait for calling instance to close / release lock on application.exe
                int i = 0;
                while (IsFileLocked(destPath) == true || i > 45)
                {
                    System.Threading.Thread.Sleep(333);
                    i++;
                }

                // Copy to final destination.
                File.Copy(exePath, destPath, true);

                // Set Registry as Installed.
                Registry.SetValue(_keyName, _isInstalledValueName, true);
                Registry.SetValue(_keyName, _pendingUpdateValueName, false);
                Registry.SetValue(_keyName, _installerPIDValueName, pid);

                // Update Add/Remove Programs Uninstall Info
                CreateUninstaller(destPath);

                // Start new installed process.
                Process.Start(destPath);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Installing:" + Environment.NewLine + ex.Message);
            }

        }

        private static void InstallProgram()
        {
            try
            {
                string exePath = Application.ExecutablePath;
                int pid = Process.GetCurrentProcess().Id;

                string destPath = _installFolder;

                Directory.CreateDirectory(destPath);

                var fullAppName = Path.GetFileName(exePath);
                destPath += "\\" + fullAppName;

                int i = 0;
                while (IsFileLocked(destPath) == true || i > 60)
                {
                    System.Threading.Thread.Sleep(250);
                    i++;
                }

                File.Copy(exePath, destPath, true);

                Registry.SetValue(_keyName, _isInstalledValueName, true);
                Registry.SetValue(_keyName, _currentVersion, Application.ProductVersion);                
                Registry.SetValue(_keyName, _installerPIDValueName, pid);

                CreateUninstaller(destPath);

                // Create Desktop Shortcut
                CreateDesktopShortcut(_assemblyName, fullAppName);

                MessageBox.Show("Successfully Installed." + Environment.NewLine + Environment.NewLine +
                    "Uninstall via Add/Remove Programs." + Environment.NewLine + Environment.NewLine +
                    "A shortcut has been placed on your desktop." + Environment.NewLine
                    );

                // Start new installed process.
                Process.Start(destPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Installing:" + Environment.NewLine + ex.Message);
            }
        }

        private static void CreateUninstaller(string location)
        {
            string ApplicationVersion = Application.ProductVersion;
            string DisplayVersion = ApplicationVersion;
            string InstallDate = DateTime.Now.ToString("yyyyMMdd");
            string UninstallString = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + _assemblyName + "\\" + _assemblyName + ".exe " + @" /uninstallprompt";

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
                        key = parent.OpenSubKey(_assemblyName, true) ?? parent.CreateSubKey(_assemblyName);

                        if (key == null)
                        {
                            throw new Exception(String.Format("Unable to create uninstaller."));
                        }

                        key.SetValue("DisplayName", _displayName);
                        key.SetValue("ApplicationVersion", ApplicationVersion);
                        key.SetValue("Publisher", _publisher);
                        key.SetValue("DisplayVersion", DisplayVersion);
                        key.SetValue("InstallDate", InstallDate);
                        key.SetValue("UninstallString", UninstallString);
                        // key.SetValue("URLInfoAbout", URLInfoAbout);
                        // key.SetValue("Contact", Contact);
                        // key.SetValue("DisplayIcon", DisplayIcon);
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
                    throw new Exception(
                        "An error occurred writing uninstall information to the registry.  The service is fully installed but can only be uninstalled manually through the command line.",
                        ex);
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
                    if (p.ProcessName.ToLower().Contains((Application.ProductName).ToLower()) == true)
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
                    localKey32 = localKey32.OpenSubKey(@"SOFTWARE\" + _assemblyName, true);
                    localKey32.DeleteValue(_installerPIDValueName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error Deleting PID from registry:" + Environment.NewLine + ex.Message);
                }
            }

        }

        private static void KillAllOtherProcesses()
        {
            var allProcceses = Process.GetProcesses();
            var currentProcessName = Application.ProductName;
            var currentProcessVersion = Application.ProductVersion;
            int currentPID = Process.GetCurrentProcess().Id;
            foreach (var process in allProcceses)
            {
                try
                {
                    if (process.ProcessName == currentProcessName)
                    {
                        if (process.Id != currentPID && process.MainModule.FileVersionInfo.ProductVersion == currentProcessVersion)
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
                @"set oShellLink = WshShell.CreateShortcut(strDesktop & ""\" + appName + @".lnk"")",
                @"oShellLink.TargetPath = ""%appdata%\" + appName + @"\" + executableName + "\"",
                @"oShellLink.WindowStyle = 1",
                @"oShellLink.Description = """ + appName + "\"",
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
                shortcutPath = Path.Combine(shortcutPath, _assemblyName + ".lnk");

                if (File.Exists(shortcutPath) == true)
                {
                    File.Delete(shortcutPath);
                }
            }
            catch { }

        }

        private static void LoadSettings(InstallerSettingsContainer settings)
        {
            _assemblyName = settings.assemblyName;
            _guidText = settings.guidText;
            _releasesURL = settings.ReleasesURL;
            _displayName = settings.DisplayName;
            _publisher = settings.Publisher;
            _displayIcon = settings.DisplayIcon;
            _uRLInfoAbout = settings.URLInfoAbout;
            _contact = settings.Contact;

            _keyName = @"HKEY_CURRENT_USER\SOFTWARE\" + settings.assemblyName;
            _isInstalledValueName = "IsInstalled";
            _installerPIDValueName = "InstalledPID";
            _currentVersion = "CurrentVersion";
            _pendingUpdateValueName = "PendingUpdate";
            _installFolder = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\" + settings.assemblyName; 
            _pendingInstallFolder = _installFolder + "\\PendingInstall";
    }

        private static void VerifyRunAsInstalled()
        {
            // Handle clicking of old installer instead of shortcut or installed exe.          

            // If this is not running from installation folder.
            if (Application.StartupPath != _installFolder)
            {
                Version thisExeVersion = Version.Parse( Application.ProductVersion);
                Version installedVersion = Version.Parse((string)Registry.GetValue(_keyName, _currentVersion, "0.0.0.0"));
                // If this is not a newer version of installed exe.
                if (installedVersion >= thisExeVersion)
                {
                    MessageBox.Show(_assemblyName + ", Version: " + installedVersion.ToString() + Environment.NewLine + "is already installed."  + Environment.NewLine + Environment.NewLine +
                        "Use Add/Remove Programs to uninstall." + Environment.NewLine + Environment.NewLine +
                        "Or use the shortcut on your Desktop to run the software.", _assemblyName + ", Version: " + installedVersion.ToString() + " is already installed.");

                    Environment.Exit(1);
                }
            }

        }
        #endregion

    }

    public class InstallerSettingsContainer
    {
        public string assemblyName { get; set; }
        public string guidText { get; set; }
        public string ReleasesURL { get; set; }
        public string DisplayName { get; set; }
        public string Publisher { get; set; }
        public string DisplayIcon { get; set; }
        public string URLInfoAbout { get; set; }
        public string Contact { get; set; }        
    }
}


