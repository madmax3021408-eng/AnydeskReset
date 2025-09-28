using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AnydeskReset
{
    public partial class Form1 : Form
    {
        private string backupDir;
        private string userConfPath;
        private string backupConfPath;
        private string anydeskExePath;
        private bool backupCreated = false;

        public Form1()
        {
            InitializeComponent();
            InitializePaths();
        }

        private void InitializePaths()
        {
            backupDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "anydesk_backup");
            userConfPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AnyDesk", "user.conf");
            backupConfPath = Path.Combine(backupDir, "user.conf");
            anydeskExePath = FindAnyDeskInstaller();
        }

        private string FindAnyDeskInstaller()
        {
            try
            {
                // Point to your program's Resources folder inside the build output
                string resourcesPath = Path.Combine(Application.StartupPath, "Resources");

                if (!Directory.Exists(resourcesPath))
                {
                    LogMessage("[ERROR] Resources folder not found.");
                    return null;
                }

                // Search for AnyDesk installer in Resources
                var files = Directory.GetFiles(resourcesPath, "AnyDesk*.exe");
                if (files.Length > 0)
                {
                    FileInfo mostRecent = null;
                    foreach (string file in files)
                    {
                        FileInfo fi = new FileInfo(file);
                        if (mostRecent == null || fi.LastWriteTime > mostRecent.LastWriteTime)
                            mostRecent = fi;
                    }
                    return mostRecent.FullName;
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Searching for installer: {ex.Message}");
            }

            return null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Clear previous logs
            richTextBox1.Clear();

            // Check admin privileges first
            if (!IsRunningAsAdministrator())
            {
                LogMessage("[ERROR] Administrator privileges required. Please re-run as Administrator.");
                return;
            }

            LogMessage("=== AnyDesk Reset & Reinstall Tool ===");
            LogMessage("Starting process...");

            try
            {
                // Execute the complete process in sequence
                Step1_CheckInstaller();
                Step2_BackupUserConf();
                Step3_KillingAnyDesk();
                Step4_RemoveAnydesk();
                Step5_InstallAnyDesk();
                Step6_KillingAnyDesk(); // Stop again after installation
                Step7_RestoreUserConf();
                Step8_RunAnydesk();
                Step9_Cleanup();

                LogMessage("[SUCCESS] All actions completed. AnyDesk is ready to use.");
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Process failed: {ex.Message}");
            }
        }

        // Helper method to log messages to RichTextBox
        private void LogMessage(string message)
        {
            if (richTextBox1.InvokeRequired)
            {
                richTextBox1.Invoke(new Action<string>(LogMessage), message);
            }
            else
            {
                richTextBox1.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                richTextBox1.ScrollToCaret();
            }
        }

        private bool IsRunningAsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        // STEP 1: Check for AnyDesk installer
        private void Step1_CheckInstaller()
        {
            if (string.IsNullOrEmpty(anydeskExePath))
            {
                LogMessage("[WARNING] AnyDesk.exe not found in Downloads folder.");
                LogMessage("Please download from: https://anydesk.com/en/downloads");

                DialogResult result = MessageBox.Show(
                    "AnyDesk installer not found in Downloads folder.\n\n" +
                    "Without the installer, automatic reinstallation will be skipped.\n\n" +
                    "Continue with removal only?",
                    "Installer Not Found",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result != DialogResult.Yes)
                {
                    throw new OperationCanceledException("User cancelled due to missing installer.");
                }
                LogMessage("User chose to continue without installer.");
            }
            else
            {
                LogMessage($"[OK] Installer found: {Path.GetFileName(anydeskExePath)}");
            }
        }

        // STEP 2: Backup user configuration
        private void Step2_BackupUserConf()
        {
            LogMessage("[1/8] Backing up user configuration...");

            try
            {
                if (!File.Exists(userConfPath))
                {
                    LogMessage("[INFO] user.conf not found - skipping backup");
                    backupCreated = false;
                    return;
                }

                // Always create backup directory if it doesn't exist
                if (!Directory.Exists(backupDir))
                {
                    Directory.CreateDirectory(backupDir);
                    LogMessage("Backup directory created");
                }

                // Always perform the backup
                File.Copy(userConfPath, backupConfPath, true);
                backupCreated = true;
                LogMessage("[OK] Backup created successfully");
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Backup failed: {ex.Message}");
                backupCreated = false;
            }
        }

        // STEP 3: Kill AnyDesk processes
        private void Step3_KillingAnyDesk()
        {
            LogMessage("[2/8] Stopping AnyDesk processes...");
            StopAnyDeskProcesses();
        }

        // STEP 4: Remove AnyDesk files and traces
        private void Step4_RemoveAnydesk()
        {
            LogMessage("[3/8] Removing AnyDesk traces...");

            List<string> pathsToRemove = new List<string>
            {
                @"C:\ProgramData\AnyDesk",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AnyDesk"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AnyDesk"),
                @"C:\Program Files (x86)\AnyDesk"
            };

            bool anyDeleted = false;
            foreach (string path in pathsToRemove)
            {
                try
                {
                    if (Directory.Exists(path))
                    {
                        Directory.Delete(path, true);
                        LogMessage($"Deleted: {path}");
                        anyDeleted = true;
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"[ERROR] Failed to delete {path}: {ex.Message}");
                }
            }

            if (anyDeleted)
            {
                LogMessage("[OK] AnyDesk traces removed");
            }
            else
            {
                LogMessage("[INFO] No AnyDesk traces found to remove");
            }
        }

        // STEP 5: Install AnyDesk
        private void Step5_InstallAnyDesk()
        {
            if (string.IsNullOrEmpty(anydeskExePath))
            {
                LogMessage("[INFO] Skipping installation - no installer found");
                return;
            }

            LogMessage("[4/8] Installing AnyDesk...");
            LogMessage($"Installer: {anydeskExePath}");

            try
            {
                // 1. Start the installer
                var installerProcess = Process.Start(new ProcessStartInfo
                {
                    FileName = anydeskExePath,
                    UseShellExecute = true
                });

                // 2. Wait a moment so installer window shows
                System.Threading.Thread.Sleep(1000);

                // 3. Show guide popup ONCE in front of everything
                using (var installForm = new Form())
                {
                    installForm.Text = "AnyDesk Installation Guide";
                    installForm.Size = new Size(800, 1000);
                    installForm.StartPosition = FormStartPosition.CenterScreen;

                    // Trick: show it once as TopMost, then disable it
                    installForm.Load += (s, e) =>
                    {
                        installForm.TopMost = true;
                        installForm.BringToFront();
                        installForm.Activate();

                        // Turn off "always on top" right after it's displayed
                        installForm.BeginInvoke(new Action(() => installForm.TopMost = false));
                    };

                    var label = new Label()
                    {
                        Text = "Follow these steps to complete installation. When finished, click Continue.",
                        Dock = DockStyle.Top,
                        Height = 60,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = new Font("Segoe UI", 10, FontStyle.Bold)
                    };

                    var step1Label = new Label()
                    {
                        Text = "Step 1",
                        Dock = DockStyle.Top,
                        Height = 40,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = new Font("Segoe UI", 10, FontStyle.Bold)
                    };

                    var picture1 = new PictureBox()
                    {
                        Dock = DockStyle.Top,
                        Height = 250,
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Image = Image.FromFile(Path.Combine(Application.StartupPath, "Resources", "anydesk_guide1.png"))
                    };

                    var step2Label = new Label()
                    {
                        Text = "Step 2",
                        Dock = DockStyle.Top,
                        Height = 40,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Font = new Font("Segoe UI", 10, FontStyle.Bold)
                    };

                    var picture2 = new PictureBox()
                    {
                        Dock = DockStyle.Top,
                        Height = 340,
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Image = Image.FromFile(Path.Combine(Application.StartupPath, "Resources", "anydesk_guide2.png"))
                    };

                    var button = new Button()
                    {
                        Text = "Continue",
                        Dock = DockStyle.Bottom,
                        Height = 40
                    };
                    button.Click += (s, e) => installForm.Close();

                    installForm.Controls.Add(button);
                    installForm.Controls.Add(picture2);
                    installForm.Controls.Add(step2Label);
                    installForm.Controls.Add(picture1);
                    installForm.Controls.Add(step1Label);
                    installForm.Controls.Add(label);

                    installForm.ShowDialog(this);
                }

                LogMessage("Installation completed by user");
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Installation failed: {ex.Message}");
            }
        }



        internal static class NativeMethods
        {
            public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
            public const UInt32 SWP_NOMOVE = 0x0002;
            public const UInt32 SWP_NOSIZE = 0x0001;
            public const UInt32 SWP_NOACTIVATE = 0x0010;

            [System.Runtime.InteropServices.DllImport("user32.dll")]
            public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
                int X, int Y, int cx, int cy, uint uFlags);
        }





        // STEP 6: Kill AnyDesk again after installation
        private void Step6_KillingAnyDesk()
        {
            LogMessage("[5/8] Stopping AnyDesk after installation...");
            StopAnyDeskProcesses();
        }

        // STEP 7: Restore user configuration
        private void Step7_RestoreUserConf()
        {
            if (!backupCreated)
            {
                LogMessage("[INFO] Skipping restoration - no backup available");
                return;
            }

            LogMessage("[6/8] Restoring user configuration...");

            try
            {
                if (File.Exists(backupConfPath))
                {
                    string userConfDir = Path.GetDirectoryName(userConfPath);
                    if (!Directory.Exists(userConfDir))
                    {
                        Directory.CreateDirectory(userConfDir);
                    }

                    File.Copy(backupConfPath, userConfPath, true);
                    LogMessage("[OK] user.conf restored successfully");
                }
                else
                {
                    LogMessage("[WARNING] Backup file not found - skipping restoration");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Restoration failed: {ex.Message}");
            }
        }

        // STEP 8: Run AnyDesk
        private void Step8_RunAnydesk()
        {
            LogMessage("[7/8] Launching AnyDesk...");

            string defaultPath = @"C:\Program Files (x86)\AnyDesk\AnyDesk.exe";

            if (File.Exists(defaultPath))
            {
                try
                {
                    Process.Start(defaultPath);
                    LogMessage("[OK] AnyDesk started successfully");
                }
                catch (Exception ex)
                {
                    LogMessage($"[ERROR] Failed to start AnyDesk: {ex.Message}");
                }
            }
            else
            {
                LogMessage("[ERROR] AnyDesk executable not found");
            }
        }

        // STEP 9: Cleanup backup files
        private void Step9_Cleanup()
        {
            LogMessage("[8/8] Cleaning up...");

            try
            {
                if (Directory.Exists(backupDir))
                {
                    Directory.Delete(backupDir, true);
                    LogMessage("Backup files cleaned up");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Cleanup failed: {ex.Message}");
            }
        }

        // Helper method to stop AnyDesk processes
        private void StopAnyDeskProcesses()
        {
            try
            {
                LogMessage("Stopping AnyDesk services...");

                // Stop AnyDesk service using PowerShell
                ExecutePowerShellCommand("Try { Stop-Service -Name 'AnyDesk' -Force -ErrorAction Stop } Catch { }");

                // Kill AnyDesk processes
                string[] processNames = { "AnyDesk", "AnyDesk_Service" };
                bool processesStopped = false;

                foreach (string processName in processNames)
                {
                    try
                    {
                        foreach (Process process in Process.GetProcessesByName(processName))
                        {
                            process.Kill();
                            process.WaitForExit(5000);
                            LogMessage($"Stopped: {processName}");
                            processesStopped = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"[INFO] No {processName} process found or already stopped");
                    }
                }

                // Additional service stop attempt
                ExecuteCommand("sc", "stop AnyDesk");

                if (processesStopped)
                {
                    LogMessage("[OK] All AnyDesk processes stopped");
                }
                else
                {
                    LogMessage("[INFO] No AnyDesk processes were running");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[ERROR] Stopping processes failed: {ex.Message}");
            }
        }

        private void ExecutePowerShellCommand(string command)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "powershell",
                    Arguments = $"-Command \"{command}\"",
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true
                };
                Process.Start(psi)?.WaitForExit(5000);
            }
            catch (Exception ex)
            {
                LogMessage($"[INFO] PowerShell command failed: {ex.Message}");
            }
        }

        private void ExecuteCommand(string command, string arguments)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = command,
                    Arguments = arguments,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    CreateNoWindow = true
                };
                Process.Start(psi)?.WaitForExit(5000);
            }
            catch (Exception ex)
            {
                LogMessage($"[INFO] Command failed: {ex.Message}");
            }
        }
    }
}