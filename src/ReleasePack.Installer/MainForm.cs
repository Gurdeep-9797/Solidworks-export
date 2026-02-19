using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ReleasePack.Installer
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            DetectSw();
        }

        private void DetectSw()
        {
            try 
            {
                string path = SolidWorksDetector.FindSolidWorksExecutable();
                if (!string.IsNullOrEmpty(path))
                {
                    _txtSwPath.Text = path;
                    _lblStatus.Text = "SolidWorks detected.";
                }
                else
                {
                    _lblStatus.Text = "SolidWorks not found. Please select manually.";
                }
            }
            catch {}
        }
        
        private void BtnBrowse_Click(object sender, EventArgs e)
        {
             using (var dlg = new OpenFileDialog { Filter = "SLDWORKS.exe|SLDWORKS.exe" })
             {
                 if (dlg.ShowDialog() == DialogResult.OK) _txtSwPath.Text = dlg.FileName;
             }
        }

        private void BtnInstall_Click(object sender, EventArgs e)
        {
            _btnInstall.Enabled = false;
            _lblStatus.Text = "Installing...";
            Application.DoEvents();

            try
            {
                // 1. Prepare Paths
                string installDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "ReleasePack");
                if (!Directory.Exists(installDir)) Directory.CreateDirectory(installDir);

                string sourceDir = Path.GetDirectoryName(Application.ExecutablePath);

                // Check for ReleasePack.AddIn.dll nearby (Release build)
                string mainDllName = "ReleasePack.AddIn.dll";
                string sourceDllPath = Path.Combine(sourceDir, mainDllName);

                // Fallback for dev environment path
                if (!File.Exists(sourceDllPath))
                {
                    string devPath = Path.GetFullPath(Path.Combine(sourceDir, @"..\..\..\ReleasePack.AddIn\bin\Release\net48", mainDllName));
                    if (File.Exists(devPath))
                    {
                         sourceDir = Path.GetDirectoryName(devPath);
                         sourceDllPath = devPath;
                    }
                }

                if (!File.Exists(sourceDllPath))
                {
                    MessageBox.Show($"Could not find source files (ReleasePack.AddIn.dll).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    _btnInstall.Enabled = true;
                    return;
                }

                // 2. Copy ALL DLLs from source to destination
                string[] files = Directory.GetFiles(sourceDir, "*.dll");
                foreach (string file in files)
                {
                    string fileName = Path.GetFileName(file);
                    string destFile = Path.Combine(installDir, fileName);
                    File.Copy(file, destFile, true);
                }

                // Also copy the Installer exe itself for uninstallation/future use
                try 
                {
                    File.Copy(Application.ExecutablePath, Path.Combine(installDir, "ReleasePack.Installer.exe"), true);
                }
                catch { /* Ignore self-copy errors if running from same dir */ }

                _lblStatus.Text = "Files copied. Registering...";
                Application.DoEvents();

                // 3. Register the MAIN DLL in the DESTINATION folder
                // This ensures dependencies (copied above) are found.
                string targetDll = Path.Combine(installDir, mainDllName);
                string result = RegistrationHelper.RegisterAddIn(targetDll);
                
                if (result == "Success")
                {
                    _lblStatus.Text = "Success!";
                    MessageBox.Show("Installation Complete!\nRestart SolidWorks.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    _lblStatus.Text = "Registration Failed.";
                    MessageBox.Show(result, "Registration Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Installation failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _lblStatus.Text = "Error.";
            }
            finally
            {
                _btnInstall.Enabled = true;
            }
        }
    }
}
