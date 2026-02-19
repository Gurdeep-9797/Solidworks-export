using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Handles creation of Bill of Materials (BOM) tables and balloons for assembly drawings.
    /// </summary>
    public class BOMEngine
    {
        private readonly ISldWorks _swApp;
        private readonly IProgressCallback _progress;

        public BOMEngine(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
        }

        /// <summary>
        /// Inserts a standard BOM table and applies auto-balloons if the drawing references an assembly.
        /// </summary>
        public void ProcessAssembly(DrawingDoc drawing, View targetView, string templatePath = "")
        {
            if (drawing == null || targetView == null) return;

            // 1. Check if view is an assembly
            ModelDoc2 model = (ModelDoc2)targetView.ReferencedDocument;
            if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                return; // Not an assembly, skip BOM
            }

            _progress?.LogMessage("Assembly detected → Generating BOM & Balloons...");

            // 2. Insert BOM Table
            TableAnnotation table = InsertBOM(drawing, targetView, templatePath);

            // 3. Auto-Balloon
            if (table != null)
            {
                ApplyAutoBalloons(drawing, targetView);
                
                // 4. Align BOM to corner (Bottom-Right or Top-Right)
                // Note: SolidWorks anchors usually handle this, but we can enforce position if needed.
            }
        }

        private TableAnnotation InsertBOM(DrawingDoc drawing, View view, string userTemplatePath)
        {
            try
            {
                // Find template
                string template = userTemplatePath;
                if (string.IsNullOrEmpty(template) || !File.Exists(template))
                {
                    template = _swApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsBOMTemplates);
                    // If directory, we need a file. Let's try to find a standard one.
                    if (!File.Exists(template))
                    {
                         // Try standard SW installation path fallback
                         string swDir = _swApp.GetExecutablePath(); // .../SLDWORKS.exe
                         string dataDir = Path.Combine(Path.GetDirectoryName(swDir), @"lang\english\bom-standard.sldbomtbt");
                         if (File.Exists(dataDir)) template = dataDir;
                    }
                }

                if (string.IsNullOrEmpty(template))
                {
                    _progress?.LogWarning("Could not find a BOM template. Skipping table.");
                    return null;
                }

                // Insert BOM Table
                // x, y are anchor points relative to sheet. 
                // We'll let the anchor point decide, or place it at (0,0) and move it.
                // InsertBomTable4(TemplateName, X, Y, BomType, Configuration)
                
                ModelDoc2 doc = (ModelDoc2)drawing;
                bool selected = doc.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0,0,0, false, 0, null, 0);
                
                if (!selected) return null;

                TableAnnotation table = (TableAnnotation)view.InsertBomTable4(
                    false, 
                    0.0, 0.0, 
                    (int)swBOMConfigurationAnchorType_e.swBOMConfigurationAnchor_BottomRight, 
                    (int)swBomType_e.swBomType_PartsOnly, 
                    "", 
                    template, 
                    false, 
                    (int)swNumberingType_e.swNumberingType_Detailed, 
                    true 
                );

                if (table != null)
                {
                    _progress?.LogMessage("BOM Table inserted.");
                }
                else
                {
                    _progress?.LogWarning("Failed to insert BOM Table.");
                }
                
                return table;
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"BOM insertion error: {ex.Message}");
                return null;
            }
        }

        private void ApplyAutoBalloons(DrawingDoc drawing, View view)
        {
            try
            {
                 ModelDoc2 doc = (ModelDoc2)drawing;
                 doc.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0,0,0, false, 0, null, 0);

                // AutoBalloon3 signature from error: (int, bool, int, int, int, string, int, string, string)
                // 1. Layout: 1 (Circular)
                // 2. IgnoreMultiple: true
                // 3. ReverseDirection: -1
                // 4. LeaderStyle: 1
                // 5. LeaderEdit: 1 (or bool -> int?)
                // 6. UpperText: "Item Number" (Maybe arg 6 is Upper Text?)
                // 7. LowerTextContent: 0
                // 8. LowerText: ""
                // 9. CustomText?
                
                drawing.AutoBalloon3(1, true, -1, 1, 1, "", 0, "", "");
                
                _progress?.LogMessage("Auto-Balloons applied.");
            }
            catch (Exception ex)
            {
                 _progress?.LogWarning($"AutoBalloon failed: {ex.Message}");
            }
        }
    }
}
