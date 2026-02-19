using System;
using System.Collections.Generic;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Generates assembly-specific drawings with exploded views, 
    /// auto-balloons, and BOM tables.
    /// </summary>
    public class AssemblyDrawingGenerator
    {
        private readonly ISldWorks _swApp;
        private readonly IProgressCallback _progress;

        public AssemblyDrawingGenerator(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
        }

        /// <summary>
        /// Generate an assembly drawing with exploded ISO view, balloons, and BOM.
        /// </summary>
        public string Generate(ModelNode assemblyNode, ExportOptions options, string outputFolder)
        {
            _progress?.LogMessage($"── Generating assembly drawing for: {assemblyNode.FileName} ──");

            ModelDoc2 modelDoc = OpenOrGetModel(assemblyNode.FilePath);
            if (modelDoc == null)
            {
                _progress?.LogError($"Could not open assembly: {assemblyNode.FilePath}");
                return null;
            }

            try
            {
                // Get template and create drawing
                string templatePath = GetDrawingTemplatePath(options);

                DrawingDoc drawing = (DrawingDoc)_swApp.NewDocument(
                    templatePath,
                    (int)swDwgPaperSizes_e.swDwgPaperA3size,
                    0.420, 0.297);

                if (drawing == null)
                {
                    _progress?.LogError("Failed to create assembly drawing document.");
                    return null;
                }

                ModelDoc2 drawingDoc = (ModelDoc2)drawing;

                // 1. Add isometric view (exploded if available)
                string viewName = GetExplodedViewName(modelDoc) ?? "*Isometric";
                View isoView = (View)drawing.CreateDrawViewFromModelView3(
                    assemblyNode.FilePath, viewName,
                    0.21, 0.17, 0); // Center of A3 sheet

                if (isoView != null)
                {
                    isoView.ScaleRatio = new double[] { 0.3, 1.0 }; // Fit assembly
                    _progress?.LogMessage($"Assembly ISO view placed (view: {viewName})");

                    // Show exploded state if available
                    if (viewName != "*Isometric")
                    {
                        try { isoView.ShowSheetMetalBendNotes = false; } catch { }
                    }
                }

                // 2. Add a front view for reference
                View frontView = (View)drawing.CreateDrawViewFromModelView3(
                    assemblyNode.FilePath, "*Front",
                    0.13, 0.17, 0);

                if (frontView != null)
                {
                    frontView.ScaleRatio = new double[] { 0.25, 1.0 };
                    _progress?.LogMessage("Assembly front view placed.");
                }

                // 3. Insert auto-balloons on the ISO view
                InsertAutoBalloons(drawing, isoView);

                // 4. Insert BOM table
                InsertBomTable(drawing, isoView, assemblyNode);

                // 5. Populate title block
                PopulateTitleBlock(drawingDoc, assemblyNode);

                // 6. Save
                string drawingPath = DrawingGenerator.GetOutputPath(assemblyNode, outputFolder, ".slddrw");
                Directory.CreateDirectory(Path.GetDirectoryName(drawingPath));

                int errors = 0, warnings = 0;
                drawingDoc.Extension.SaveAs2(
                    drawingPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null, "", false, ref errors, ref warnings);

                _progress?.LogMessage($"Assembly drawing saved: {drawingPath}");
                return drawingPath;
            }
            catch (Exception ex)
            {
                _progress?.LogError($"Assembly drawing generation failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Insert auto-balloons on a drawing view.
        /// Version-aware: tries AutoBalloon4 (v29+/SW 2021), then AutoBalloon3 (all versions).
        /// </summary>
        private void InsertAutoBalloons(DrawingDoc drawing, View view)
        {
            if (view == null) return;

            try
            {
                drawing.ActivateView(view.Name);

                ModelDoc2 doc = (ModelDoc2)drawing;
                doc.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                object balloonsObj = null;
                int apiVersion = SwVersionHelper.GetMajorVersion(_swApp);

                // Try AutoBalloon4 (SW 2021+ / v29+)
                if (apiVersion >= 29)
                {
                    try
                    {
                        balloonsObj = drawing.AutoBalloon4(
                            0, true,
                            (int)swBalloonStyle_e.swBS_Circular,
                            (int)swBalloonFit_e.swBF_2Chars,
                            0, "", 0, "", "", true);
                    }
                    catch { _progress?.LogMessage("AutoBalloon4 not available, trying fallback..."); }
                }

                // Fallback: AutoBalloon3 (available since ~v26 / SW 2018)
                if (balloonsObj == null)
                {
                    try
                    {
                        balloonsObj = drawing.AutoBalloon3(
                            0, true,
                            (int)swBalloonStyle_e.swBS_Circular,
                            (int)swBalloonFit_e.swBF_2Chars,
                            0, "", 0, "", "");
                    }
                    catch { _progress?.LogMessage("AutoBalloon3 fallback also failed."); }
                }

                if (balloonsObj != null)
                    _progress?.LogMessage("Auto-balloons inserted.");
                else
                    _progress?.LogWarning("Auto-balloons returned null. Manual balloon placement may be needed.");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Auto-balloon insertion failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Insert a BOM table linked to the drawing view.
        /// Version-aware: tries InsertBomTable4 (v28+/SW 2020), then InsertBomTable3 (all versions).
        /// </summary>
        private void InsertBomTable(DrawingDoc drawing, View view, ModelNode assemblyNode)
        {
            if (view == null) return;

            try
            {
                drawing.ActivateView(view.Name);
                string bomTemplate = FindBomTemplate();
                string config = assemblyNode.ConfigurationName ?? "";

                BomTableAnnotation bomTable = null;
                int apiVersion = SwVersionHelper.GetMajorVersion(_swApp);

                // Try InsertBomTable4 (SW 2020+ / v28+)
                if (apiVersion >= 28)
                {
                    try
                    {
                        bomTable = (BomTableAnnotation)view.InsertBomTable4(
                            false, 0.40, 0.29,
                            (int)swBomType_e.swBomType_TopLevelOnly,
                            0, config, bomTemplate, false,
                            (int)swNumberingType_e.swNumberingType_Detailed, false);
                    }
                    catch { _progress?.LogMessage("InsertBomTable4 not available, trying fallback..."); }
                }

                // Fallback: InsertBomTable3 (available since ~v26 / SW 2018)
                if (bomTable == null)
                {
                    try
                    {
                        bomTable = (BomTableAnnotation)view.InsertBomTable3(
                            false, 0.40, 0.29,
                            (int)swBomType_e.swBomType_TopLevelOnly,
                            0, config, bomTemplate, false);
                    }
                    catch { _progress?.LogMessage("InsertBomTable3 fallback also failed."); }
                }

                if (bomTable != null)
                    _progress?.LogMessage("BOM table inserted into assembly drawing.");
                else
                    _progress?.LogWarning("BOM table insertion returned null.");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"BOM table insertion failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Try to find an exploded view configuration in the assembly.
        /// </summary>
        private string GetExplodedViewName(ModelDoc2 modelDoc)
        {
            try
            {
                ConfigurationManager configMgr = modelDoc.ConfigurationManager;
                string[] configNames = (string[])modelDoc.GetConfigurationNames();

                if (configNames == null) return null;

                foreach (string name in configNames)
                {
                    if (name.ToLowerInvariant().Contains("explod"))
                    {
                        _progress?.LogMessage($"Found exploded configuration: {name}");
                        return "*Isometric"; // View from the exploded config
                    }
                }
            }
            catch { }

            return null;
        }

        private void PopulateTitleBlock(ModelDoc2 drawingDoc, ModelNode node)
        {
            try
            {
                CustomPropertyManager propMgr =
                    (CustomPropertyManager)drawingDoc.Extension.CustomPropertyManager[""];
                if (propMgr == null) return;

                propMgr.Add3("PartNumber", (int)swCustomInfoType_e.swCustomInfoText,
                    node.PartNumber, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                propMgr.Add3("Description", (int)swCustomInfoType_e.swCustomInfoText,
                    node.Description, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                propMgr.Add3("Revision", (int)swCustomInfoType_e.swCustomInfoText,
                    node.Revision, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                propMgr.Add3("DrawnBy", (int)swCustomInfoType_e.swCustomInfoText,
                    System.Environment.UserName, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                propMgr.Add3("DrawnDate", (int)swCustomInfoType_e.swCustomInfoText,
                    DateTime.Now.ToString("yyyy-MM-dd"), (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Title block population warning: {ex.Message}");
            }
        }

        /// <summary>
        /// Get the drawing template path using the shared multi-source resolver.
        /// </summary>
        private string GetDrawingTemplatePath(ExportOptions options)
        {
            return SwVersionHelper.FindDrawingTemplate(_swApp, options);
        }

        private string FindBomTemplate()
        {
            try
            {
                string swPath = _swApp.GetExecutablePath();
                string installDir = Path.GetDirectoryName(swPath);

                string[] searchPaths = new[]
                {
                    Path.Combine(installDir, @"lang\english\bom-standard.sldbomtbt"),
                    Path.Combine(installDir, @"lang\english\bom-all.sldbomtbt"),
                };

                foreach (string p in searchPaths)
                {
                    if (File.Exists(p)) return p;
                }
            }
            catch { }

            return "";
        }

        private ModelDoc2 OpenOrGetModel(string filePath)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.GetOpenDocumentByName(filePath);
            if (doc != null) return doc;

            int errors = 0, warnings = 0;
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            int docType = ext == ".sldasm" ? (int)swDocumentTypes_e.swDocASSEMBLY
                        : (int)swDocumentTypes_e.swDocPART;

            return (ModelDoc2)_swApp.OpenDoc6(filePath, docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "", ref errors, ref warnings);
        }
    }
}
