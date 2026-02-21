using System;
using System.IO;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.AddIn.UI
{
    /// <summary>
    /// Builds and manages the native SolidWorks CommandManager Tab (Ribbon)
    /// for the V2 Release Pack Subsystem.
    /// Acts as a centralized module similar to "Weldments" or "Sheet Metal".
    /// </summary>
    public class CommandTabBuilder
    {
        private ISldWorks _swApp;
        private int _addinCookie;
        private ICommandManager _cmdMgr;
        private ICommandGroup _cmdGroup;
        
        // Command indices returned by AddCommandItem2
        private int _idxGenerate;
        private int _idxBreakdown;
        private int _idxSettings;
        
        // Command IDs
        public const int CMD_GROUP_ID = 42;
        public const int CMD_GENERATE = 1;
        public const int CMD_ASSEMBLY_BREAKDOWN = 2;
        public const int CMD_SETTINGS = 3;
        
        private string _iconsDir;

        public CommandTabBuilder(ISldWorks swApp, int addinCookie)
        {
            _swApp = swApp;
            _addinCookie = addinCookie;
            
            // Resolve icon path dynamically
            string assemblyDir = Path.GetDirectoryName(new Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            _iconsDir = Path.Combine(assemblyDir, "UI", "Icons");
        }

        public void Build()
        {
            try
            {
                _cmdMgr = _swApp.GetCommandManager(_addinCookie);
                if (_cmdMgr == null) return;

                // 1. Create the CommandGroup (houses the actual executable buttons)
                CreateCommandGroup();

                // 2. Create the CommandTab (the visual ribbon UI) in relevant document types
                CreateCommandTab((int)swDocumentTypes_e.swDocPART);
                CreateCommandTab((int)swDocumentTypes_e.swDocASSEMBLY);
                CreateCommandTab((int)swDocumentTypes_e.swDocDRAWING);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine("Failed to build Command Tab: " + ex.Message);
            }
        }

        public void Remove()
        {
            try
            {
                if (_cmdMgr == null)
                    _cmdMgr = _swApp.GetCommandManager(_addinCookie);
                    
                _cmdMgr?.RemoveCommandGroup(CMD_GROUP_ID);
                ClearCommandTabs();
            }
            catch { }
        }

        private void CreateCommandGroup()
        {
            int err = 0;
            
            // 1=Main List, 2=Toolbar, 3=Menu
            _cmdGroup = _cmdMgr.CreateCommandGroup2(
                CMD_GROUP_ID,
                "Release Pack Subsystem",
                "Advanced V2 Drafting and Export Subsystem",
                "", -1, true, ref err);

            if (_cmdGroup != null)
            {
                // Assign Icons (Fallback to built-in if missing)
                string icon16 = Path.Combine(_iconsDir, "icon_16.png");
                string icon32 = Path.Combine(_iconsDir, "icon_32.png");
                
                if (File.Exists(icon16)) _cmdGroup.SmallMainIcon = icon16;
                if (File.Exists(icon32)) _cmdGroup.LargeMainIcon = icon32;

                // Button 1: Master Generation
                _idxGenerate = _cmdGroup.AddCommandItem2(
                    "Generate Drafts & Exports",
                    -1,
                    "Run V2 Subsystem Engine",
                    "Executes intelligent dimensioning, layering, and export generation",
                    0, 
                    "OnBtnGenerateClick",
                    "OnBtnGenerateEnable",
                    CMD_GROUP_ID, CMD_GENERATE);

                // Button 2: Assembly Tree Breakdown
                _idxBreakdown = _cmdGroup.AddCommandItem2(
                    "Assembly Breakdown",
                    -1,
                    "Analyze Assembly Topology",
                    "Generates visual and data tree of multi-level assemblies",
                    0, 
                    "OnBtnBreakdownClick",
                    "OnBtnBreakdownEnable",
                    CMD_GROUP_ID, CMD_ASSEMBLY_BREAKDOWN);
                    
                // Button 3: Settings
                _idxSettings = _cmdGroup.AddCommandItem2(
                    "System Settings",
                    -1,
                    "Configure Template and Standard behaviors",
                    "Toggle ISO/ANSI and Layer definitions",
                    0, 
                    "OnBtnSettingsClick",
                    "OnBtnSettingsEnable",
                    CMD_GROUP_ID, CMD_SETTINGS);

                _cmdGroup.HasToolbar = true;
                _cmdGroup.HasMenu = true;
                _cmdGroup.Activate();
            }
        }

        private void CreateCommandTab(int docType)
        {
            try
            {
                // Attempt to get existing tab
                ICommandTab cmdTab = _cmdMgr.GetCommandTab(docType, "Release Pack");
                
                // If it doesn't exist, create it
                if (cmdTab == null)
                {
                    cmdTab = _cmdMgr.AddCommandTab(docType, "Release Pack");
                    
                    if (cmdTab != null)
                    {
                        CommandTabBox cmdBox = cmdTab.AddCommandTabBox();
                        
                        // We map the commands from our group to the ribbon using global Command IDs
                        int[] cmdIDs = new int[] 
                        { 
                            _cmdGroup.get_CommandID(_idxGenerate), 
                            _cmdGroup.get_CommandID(_idxBreakdown), 
                            _cmdGroup.get_CommandID(_idxSettings) 
                        };
                        
                        int[] textTypes = new int[] 
                        { 
                            (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal,
                            (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_NoText,
                            (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow
                        };
                        
                        cmdBox.AddCommands(cmdIDs, textTypes);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"Failed to create Command Tab for docType {docType}: " + ex.Message);
            }
        }
        
        private void ClearCommandTabs()
        {
            try
            {
                var docTypes = new[] { (int)swDocumentTypes_e.swDocPART, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swDocumentTypes_e.swDocDRAWING };
                foreach (int docType in docTypes)
                {
                    ICommandTab tab = _cmdMgr.GetCommandTab(docType, "Release Pack");
                    if (tab != null) _cmdMgr.RemoveCommandTab((CommandTab)tab);
                }
            }
            catch { }
        }
    }
}
