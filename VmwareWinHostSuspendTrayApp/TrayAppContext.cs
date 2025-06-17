using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.Json;
using System.Drawing;

namespace LogikVmwareWinHostSuspendTrayApp
{
    public class TrayAppContext : ApplicationContext
    {
        private NotifyIcon trayIcon;
        private List<string> vmPaths;
        private string configPath = Path.Combine(Application.StartupPath, "appsettings.json");
        private string logPath = Path.Combine(Application.StartupPath, "log.txt");
        private HiddenForm hiddenForm;

        public TrayAppContext()
        {
            LoadConfig();
            RegisterStartup();
            hiddenForm = new HiddenForm(SuspendVMs, Log);
            trayIcon = new NotifyIcon()
            {
                Icon = SystemIcons.Application,
                Visible = true,
                Text = "VMware Host Suspend Helper",
                ContextMenuStrip = new ContextMenuStrip()
                {
                    Items = {
                        new ToolStripMenuItem("Suspension manuelle", null, (s, e) => SuspendVMs()),
                        new ToolStripMenuItem("Ouvrir journal", null, (s, e) => OpenLog()),
                        new ToolStripMenuItem("Paramètres", null, (s, e) => OpenSettings()),
                        new ToolStripMenuItem("Quitter", null, (s, e) => ExitApp()),
                        new ToolStripMenuItem("Quitter && désactiver le démarrage automatique", null, (s, e) => UnregisterAndExit())
                    }
                }

            };

            Log("Application démarrée.");
        }

        private void UnregisterAndExit()
        {
            string appName = "VmwareWinHostSuspendTrayApp";
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                if (key.GetValue(appName) != null)
                {
                    key.DeleteValue(appName);
                    Log("Démarrage automatique désactivé via le menu.");
                }
            }
            ExitApp();
        }


        private void SuspendVMs()
        {
            string vmrunPath = @"C:\\Program Files (x86)\\VMware\\VMware Workstation\\vmrun.exe";
            foreach (var vmx in vmPaths)
            {
                try
                {
                    Process.Start(vmrunPath, $"suspend \"{vmx}\" soft");
                    Log($"Suspension de la VM : {vmx}");
                }
                catch (Exception ex)
                {
                    Log($"Erreur lors de la suspension : {ex.Message}");
                }
            }
        }

        private void RegisterStartup()
        {
            string appName = "VmwareWinHostSuspendTrayApp";
            string exePath = Application.ExecutablePath;
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                if ((string)key.GetValue(appName) != exePath)
                {
                    key.SetValue(appName, exePath);
                    Log("Ajout de l'application au démarrage automatique.");
                }
            }
        }

        private void LoadConfig()
        {
            try
            {
                if (File.Exists(configPath))
                {
                    var json = File.ReadAllText(configPath);
                    var doc = JsonSerializer.Deserialize<Dictionary<string, List<string>>>(json);
                    vmPaths = doc["VMPaths"];
                }
                else
                {
                    vmPaths = new List<string>();
                }
            }
            catch
            {
                vmPaths = new List<string>();
            }
        }

        private void SaveConfig()
        {
            var doc = new Dictionary<string, List<string>> { { "VMPaths", vmPaths } };
            var json = JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(configPath, json);
        }

        private void OpenLog()
        {
            if (File.Exists(logPath))
                Process.Start("notepad.exe", logPath);
        }

        private void OpenSettings()
        {
            using (var ofd = new OpenFileDialog())
            {
                ofd.Multiselect = true;
                ofd.Filter = "VMware VMX files (*.vmx)|*.vmx";
                ofd.Title = "Sélectionner une ou plusieurs machines virtuelles";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    vmPaths = ofd.FileNames.ToList();
                    SaveConfig();
                    Log($"Nouvelles VMs sélectionnées : {string.Join(", ", vmPaths)}");
                }
            }
        }

        private void ExitApp()
        {
            Log("Fermeture de l'application.");
            trayIcon.Visible = false;
            hiddenForm.Close();
            Application.Exit();
        }

        private void Log(string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            File.AppendAllText(logPath, $"{timestamp} - {message}{Environment.NewLine}");
        }

        // Hidden form to capture WM_POWERBROADCAST messages
        private class HiddenForm : Form
        {
            private readonly Action suspendAction;
            private readonly Action<string> logger;

            public HiddenForm(Action suspendAction, Action<string> logger)
            {
                this.suspendAction = suspendAction;
                this.logger = logger;
                this.ShowInTaskbar = false;
                this.WindowState = FormWindowState.Minimized;
                this.Visible = false;
            }

            protected override void WndProc(ref Message m)
            {
                const int WM_POWERBROADCAST = 0x0218;
                const int PBT_APMSUSPEND = 0x0004;

                if (m.Msg == WM_POWERBROADCAST && m.WParam.ToInt32() == PBT_APMSUSPEND)
                {
                    logger("WM_POWERBROADCAST reçu : mise en veille imminente.");
                    suspendAction();
                }

                base.WndProc(ref m);
            }
        }
    }
}
