using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using ReleasePack.AddIn.UI;

namespace ReleasePack.AddIn
{
    /// <summary>
    /// Main COM Add-In entry point for SolidWorks Release Pack.
    /// Registers a toolbar button and a TaskPane panel for the export UI.
    /// </summary>
    [Guid("E8A3F1B0-7C4D-4E5A-9B2F-3D6A8C1E0F42")]
    [ComVisible(true)]
    [ProgId("ReleasePack.AddIn")]
    public class SwAddin : ISwAddin
    {
        public static SwAddin Instance { get; private set; }
        public ISldWorks SwApp => _swApp;

        private ISldWorks _swApp;
        private int _addinCookie;
        private ITaskpaneView _taskpaneView;
        private ReleasePackTaskPane _taskPane;

        private const int CMD_GROUP_ID = 42;
        private const int CMD_RELEASE_PACK = 1;

        #region ISwAddin Implementation

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            Instance = this;
            _swApp = (ISldWorks)ThisSW;
            _addinCookie = Cookie;

            _swApp.SetAddinCallbackInfo(0, this, _addinCookie);

            CreateCommandManager();
            CreateTaskpane();

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandManager();

            if (_taskpaneView != null)
            {
                _taskpaneView.DeleteView();
                Marshal.ReleaseComObject(_taskpaneView);
                _taskpaneView = null;
            }

            _taskPane = null;
            _swApp = null;
            Instance = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            return true;
        }

        #endregion

        #region UI Setup

        private void CreateCommandManager()
        {
            try
            {
                ICommandManager cmdMgr = _swApp.GetCommandManager(_addinCookie);

                // Get the assembly directory
                string assemblyDir = System.IO.Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
                string iconDir = System.IO.Path.Combine(assemblyDir, "UI", "Icons");

                string[] iconList = new string[] {
                    System.IO.Path.Combine(iconDir, "icon_20.png"),
                    System.IO.Path.Combine(iconDir, "icon_32.png"),
                    System.IO.Path.Combine(iconDir, "icon_40.png"),
                    System.IO.Path.Combine(iconDir, "icon_64.png"),
                    System.IO.Path.Combine(iconDir, "icon_96.png"),
                    System.IO.Path.Combine(iconDir, "icon_128.png")
                };

                CommandGroup cmdGroup = cmdMgr.CreateCommandGroup(CMD_GROUP_ID,
                    "Release Pack", "One-click release pack generator", "", -1);

                if (cmdGroup != null)
                {
                    cmdGroup.LargeMainIcon = System.IO.Path.Combine(iconDir, "icon_32.png");
                    cmdGroup.SmallMainIcon = System.IO.Path.Combine(iconDir, "icon_16.png");
                    // cmdGroup.LargeIconList = iconList; // Commented out due to interop version mismatch
                    // cmdGroup.SmallIconList = iconList;

                    int index = cmdGroup.AddCommandItem2(
                        "Generate Release Pack", -1,
                        "Open Release Pack Generator", "Click to open Release Pack panel",
                        0, "OnReleasePackClick", "", CMD_GROUP_ID, CMD_RELEASE_PACK);

                    cmdGroup.HasToolbar = true;
                    cmdGroup.HasMenu = true;
                    cmdGroup.Activate();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("CreateCommandManager failed: " + ex.Message);
            }
        }

        private void RemoveCommandManager()
        {
            try
            {
                ICommandManager cmdMgr = _swApp.GetCommandManager(_addinCookie);
                cmdMgr.RemoveCommandGroup(CMD_GROUP_ID);
            }
            catch { }
        }

        private void CreateTaskpane()
        {
            try
            {
                _taskpaneView = _swApp.CreateTaskpaneView2(
                    string.Empty, "Release Pack Generator");

                if (_taskpaneView != null)
                {
                    _taskPane = (ReleasePackTaskPane)_taskpaneView.AddControl(
                        typeof(ReleasePackTaskPane).FullName, "");

                    if (_taskPane != null)
                    {
                        _taskPane.Initialize(_swApp);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("CreateTaskpane failed: " + ex.Message);
            }
        }

        #endregion

        #region Callbacks

        public void OnReleasePackClick()
        {
            // Show the taskpane if hidden
            if (_taskpaneView != null)
            {
                _taskpaneView.ShowView();
            }
        }

        #endregion

        #region COM Registration

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            try
            {
                var hklm = Microsoft.Win32.Registry.LocalMachine;
                var hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = @"SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}";

                var addinKey = hklm.CreateSubKey(keyname);
                addinKey.SetValue(null, 0);
                addinKey.SetValue("Description", "One-click industrial release pack generator with smart drawings, DXF, STEP, PDF, BOM");
                addinKey.SetValue("Title", "SolidWorks Release Pack");

                keyname = @"Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}";
                addinKey = hkcu.CreateSubKey(keyname);
                addinKey.SetValue(null, 1); // Load on startup
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error registering Release Pack Add-In: " + ex.Message);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                var hklm = Microsoft.Win32.Registry.LocalMachine;
                var hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = @"SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname, false);

                keyname = @"Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname, false);
            }
            catch { }
        }

        #endregion
    }
}
