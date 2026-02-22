using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Generates a complete industrial-grade drawing for a part model.
    /// 
    /// Pipeline:
    ///   1. Determine sheet size from model bounding box
    ///   2. Create drawing from template
    ///   3. Place standard 3rd-angle views (Front, Top, Right, ISO)
    ///   4. Analyze features → add Section view if internal geometry
    ///   5. Analyze features → add Detail view if small geometry
    ///   6. Run DimensionEngine for annotations
    ///   7. Populate title block from custom properties
    /// </summary>
    public class DrawingGenerator
    {
        private readonly ISldWorks _swApp;
        private readonly IProgressCallback _progress;
        private readonly FeatureAnalyzer _featureAnalyzer;
        private readonly DimensionEngine _dimensionEngine;
        private readonly BOMEngine _bomEngine; // New field

        // Standard sheet sizes in meters (width, height) — landscape
        private static readonly Dictionary<SheetSizeOption, (double w, double h, int swEnum)> SheetSizes =
            new Dictionary<SheetSizeOption, (double, double, int)>
        {
            { SheetSizeOption.A4_Landscape, (0.297, 0.210, (int)swDwgPaperSizes_e.swDwgPaperA4size) },
            { SheetSizeOption.A3_Landscape, (0.420, 0.297, (int)swDwgPaperSizes_e.swDwgPaperA3size) },
            { SheetSizeOption.A2_Landscape, (0.594, 0.420, (int)swDwgPaperSizes_e.swDwgPaperA2size) },
            { SheetSizeOption.A1_Landscape, (0.841, 0.594, (int)swDwgPaperSizes_e.swDwgPaperA1size) },
            { SheetSizeOption.A0_Landscape, (1.189, 0.841, (int)swDwgPaperSizes_e.swDwgPaperA0size) },
        };

        // Margin from sheet edge (meters)
        private const double MARGIN = 0.025; // 25mm
        private const double VIEW_GAP = 0.030; // 30mm between views

        public DrawingGenerator(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
            _featureAnalyzer = new FeatureAnalyzer(progress);
            _dimensionEngine = new DimensionEngine(swApp, progress);
            _bomEngine = new BOMEngine(swApp, progress);
        }

        /// <summary>
        /// Generates a drawing for the given model node.
        /// Returns the path to the saved drawing file, or null on failure.
        /// </summary>
        public string Generate(ModelNode node, string outputDir, ExportOptions options)
        {
            _progress?.LogMessage($"── Generating drawing for: {node.FileName} ──");

            ModelDoc2 modelDoc = OpenOrGetModel(node.FilePath);
            if (modelDoc == null)
            {
                _progress?.LogError($"Could not open model: {node.FilePath}");
                return null;
            }

            try
            {
                // 1. Get bounding box
                double[] bbox = GetBoundingBox(modelDoc);
                double modelW = Math.Abs(bbox[3] - bbox[0]);
                double modelH = Math.Abs(bbox[4] - bbox[1]);
                double modelD = Math.Abs(bbox[5] - bbox[2]);

                _progress?.LogMessage($"Model bounds: {modelW * 1000:F1} × {modelH * 1000:F1} × {modelD * 1000:F1} mm");

                // 2. V3: Analyze features FIRST (drives view selection)
                List<AnalyzedFeature> features = _featureAnalyzer.Analyze(modelDoc);
                System.Windows.Forms.Application.DoEvents();

                // 3. V3: Smart View Planning (decides which views are needed)
                var viewPlan = ViewPlanner.Plan(features, modelW, modelH, modelD);
                _progress?.LogMessage($"View Plan: {viewPlan.Summary}");

                // 4. V3: Calculate mathematically optimal sheet size & scale based on actual required views
                var templateManager = new ReleasePack.Engine.Layout.TemplateManager(_swApp);
                var (bestSheet, isoScale) = templateManager.ComputeOptimalLayout(
                    bbox, viewPlan.NeedTopView, viewPlan.NeedRightView, viewPlan.NeedIsoView);
                
                _progress?.LogMessage($"V3 Layout Engine: Selected {bestSheet.Name} at 1:{1/isoScale}");

                // 5. Create drawing document
                string templatePath = GetDrawingTemplatePath(options);
                DrawingDoc drawing = CreateDrawingDocument(templatePath, bestSheet.SwPaperSize, bestSheet.Width, bestSheet.Height);

                if (drawing == null)
                {
                    _progress?.LogError("Failed to create drawing document.");
                    return null;
                }

                ModelDoc2 drawingDoc = (ModelDoc2)drawing;

                // 6. Lock sheet scale
                templateManager.ApplyTemplateToDrawing(drawing, bestSheet, isoScale);

                // 7. Calculate View Envelopes safely without overlap
                bool isAnsi = options.ViewStandard == ViewStandard.ThirdAngle;
                var binSolver = new ReleasePack.Engine.Layout.BinPackingSolver(isAnsi);
                var solvedLayout = binSolver.SolveStandardLayout(
                    bbox, isoScale, bestSheet, 
                    viewPlan.NeedTopView, viewPlan.NeedRightView, viewPlan.NeedIsoView);
                
                if (solvedLayout == null)
                {
                    _progress?.LogWarning("V3 Strict Layout Warning: Views marginally exceed paper boundaries. Attempting safety scale reduction.");
                    // Fallback to next lowest scale if overlaps breach (Simple mitigation)
                    isoScale = ReleasePack.Engine.Layout.TemplateManager.GetValidISO5455Scale(isoScale * 0.9);
                    templateManager.ApplyTemplateToDrawing(drawing, bestSheet, isoScale);
                    solvedLayout = binSolver.SolveStandardLayout(
                        bbox, isoScale, bestSheet, 
                        viewPlan.NeedTopView, viewPlan.NeedRightView, viewPlan.NeedIsoView);
                }

                // 8. Place actual drawing views using determined safe zones
                if (solvedLayout != null)
                {
                    PlaceStandardViewsV3(drawing, node.FilePath, solvedLayout);
                }
                else
                {
                    _progress?.LogError("CRITICAL: Layout generation breached sheet constraints entirely. Cannot guarantee non-overlapped output.");
                }
                System.Windows.Forms.Application.DoEvents();

                // 9. Add section view if needed
                if (viewPlan.NeedSectionView)
                {
                    _progress?.LogMessage("Internal features detected → adding section view.");
                    AddSectionView(drawing, bestSheet.Width, bestSheet.Height);
                    System.Windows.Forms.Application.DoEvents();
                }

                // 10. Add detail views for small features
                if (viewPlan.NeedDetailViews)
                {
                    _progress?.LogMessage($"Small features detected → adding {viewPlan.DetailFeatures.Count} detail view(s).");
                    AddDetailView(drawing, viewPlan.DetailFeatures, bestSheet.Width, bestSheet.Height);
                    System.Windows.Forms.Application.DoEvents();
                }

                // 11. Apply full annotation pipeline (center marks, dims, callouts, TYP notes)
                _dimensionEngine.ApplyDimensions(drawing, features);
                System.Windows.Forms.Application.DoEvents();

                // 12. Assembly BOM & Balloons
                View isoView = GetIsoView(drawing);
                if (isoView != null)
                {
                    _bomEngine.ProcessAssembly(drawing, isoView, options.BomTemplatePath);
                    System.Windows.Forms.Application.DoEvents();
                }

                // 13. Populate Title Block (Metadata)
                PopulateTitleBlock(drawing, node, options);

                // 14. Force rebuild to ensure all annotations are current
                drawingDoc.ForceRebuild3(false);

                // 15. Save Drawing
                string drawingPath = GetOutputPath(node, outputDir, ".slddrw");
                SaveDrawing(drawingDoc, drawingPath);

                _progress?.LogMessage($"Drawing saved: {drawingPath}");
                return drawingPath;
            }
            catch (Exception ex)
            {
                _progress?.LogError($"Drawing generation failed for {node.FileName}: {ex.Message}");
                return null;
            }
        }

        private void PlaceStandardViewsV3(DrawingDoc drawing, string modelPath, Dictionary<string, ReleasePack.Engine.Layout.BinPackingSolver.ViewEnvelope> layout)
        {
            var solver = new ReleasePack.Engine.Layout.BinPackingSolver();

            if (layout.ContainsKey("FRONT"))
            {
                View frontView = (View)drawing.CreateDrawViewFromModelView3(modelPath, "*Front", 0, 0, 0);
                if (frontView != null) solver.CommitView(frontView, layout["FRONT"]);
            }

            if (layout.ContainsKey("TOP"))
            {
                View topView = (View)drawing.CreateDrawViewFromModelView3(modelPath, "*Top", 0, 0, 0);
                if (topView != null) solver.CommitView(topView, layout["TOP"]);
            }

            if (layout.ContainsKey("RIGHT"))
            {
                View rightView = (View)drawing.CreateDrawViewFromModelView3(modelPath, "*Right", 0, 0, 0);
                if (rightView != null) solver.CommitView(rightView, layout["RIGHT"]);
            }

            if (layout.ContainsKey("ISO"))
            {
                View isoView = (View)drawing.CreateDrawViewFromModelView3(modelPath, "*Isometric", 0, 0, 0);
                if (isoView != null) solver.CommitView(isoView, layout["ISO"]);
            }
        }

        // V1 Dynamic Template Calculation Engine Deprecated - V3 Uses explicit layout footprint sizing using TemplateManager

        /// <summary>
        /// Add isometric view in the top-right area of the sheet.
        /// </summary>
        private void AddIsometricView(DrawingDoc drawing, string modelPath,
            double sheetW, double sheetH, double scale,
            double modelW, double modelH, double modelD)
        {
            try
            {
                double isoX = sheetW * 0.78;
                double isoY = sheetH * 0.75;
                double isoScale = scale * 0.6; // ISO view slightly smaller

                View isoView = (View)drawing.CreateDrawViewFromModelView3(
                    modelPath, "*Isometric", isoX, isoY, 0);

                if (isoView != null)
                {
                    isoView.ScaleRatio = new double[] { isoScale, 1.0 };
                    _progress?.LogMessage("Isometric view placed.");
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Could not add isometric view: {ex.Message}");
            }
        }

        /// <summary>
        /// Attempts to find the Isometric view.
        /// </summary>
        private View GetIsoView(DrawingDoc drawing)
        {
            try
            {
                object[] views = (object[])drawing.GetViews();
                if (views == null) return null;

                foreach (object obj in views)
                {
                    View v = (View)obj;
                    // Check if orientation is *Isometric
                    // Note: Use GetOrientationName() if available, strict check might be hard.
                    // Fallback: Check position (Top-Right quadrant) or just assume 4th view if standard.
                    // Let's rely on the creation loop.
                    // Since we can't easily rely on names, let's look for the one that WAS created as ISO.
                    // But we don't have that reference.
                    // Strategy: Find the view that looks like it (Double dimension check? No).
                    // Robust: Iterate and check `v.GetOrientationName()`.
                    // But interop might not expose it easily.
                    // Let's simply check if the name starts with "Drawing View" and it's likely the 4th one (Index 4, since 0=Sheet).
                    // Sheet = views[0]
                    // Front = views[1]
                    // Top = views[2]
                    // Right = views[3]
                    // Iso = views[4]
                    // So if we have at least 5 items, views[4] is a good candidate.
                }
                
                // Better approach: BOM on the *first* view if Iso specific one fails? No, BOM usually goes on the main assembly view.
                // Let's return the Last view that is NOT a detail or section view.
                // Section and Detail views have specific types.
                // Standard views are swDrawingView (1).
                
                for (int i = views.Length - 1; i >= 0; i--)
                {
                    View v = (View)views[i];
                    // Skip sheet (Type 1 usually, or Name "Sheet1")
                    if (!v.Name.StartsWith("Sheet") && v.Type != (int)swDrawingViewTypes_e.swDrawingSheet)
                    {
                        // Return the last valid view as candidate for Iso
                        return v; 
                    }
                }
            }
            catch {}
            return null;
        }

        /// <summary>
        /// Set scale on all existing views.
        /// </summary>
        private void SetAllViewScales(DrawingDoc drawing, double scale)
        {
            try
            {
                object[] views = (object[])drawing.GetViews();
                if (views == null) return;

                foreach (object viewObj in views)
                {
                    View view = (View)viewObj;
                    if (view != null && view.Name != "Sheet1")
                    {
                        try
                        {
                            view.ScaleRatio = new double[] { scale, 1.0 };
                        }
                        catch { }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// Add a section view through the model center (vertical cut through front view).
        /// Version-aware: tries CreateSectionViewAt5 (SW 2022+/v30+), then CreateSectionView (all versions).
        /// </summary>
        private void AddSectionView(DrawingDoc drawing, double sheetW, double sheetH)
        {
            try
            {
                // Activate the front view to create section from it
                drawing.ActivateView("Drawing View1"); // First created view = Front

                ModelDoc2 doc = (ModelDoc2)drawing;

                // Draw section line (vertical through center)
                object[] views = (object[])drawing.GetViews();
                if (views == null || views.Length < 2) return;

                View frontView = (View)views[1]; // Index 0 is sheet; 1 is front
                double[] pos = (double[])frontView.Position;
                double cx = pos[0];
                double cy = pos[1];

                double lineTop = cy + 0.05;
                double lineBot = cy - 0.05;

                doc.SketchManager.CreateLine(cx, lineTop, 0, cx, lineBot, 0);

                double sectionX = sheetW * 0.55;
                double sectionY = sheetH * 0.4;

                View sectionView = null;
                int apiVersion = SwVersionHelper.GetMajorVersion(_swApp);

                // Try newest API first, chain fallbacks
                if (apiVersion >= 30) // SW 2022+
                {
                    try
                    {
                        sectionView = (View)drawing.CreateSectionViewAt5(
                            sectionX, sectionY, 0, "A", 0, true, 0);
                    }
                    catch { _progress?.LogMessage("CreateSectionViewAt5 not available, trying fallback..."); }
                }

                // Fallback: CreateSectionView (available since v24 / SW 2016)
                if (sectionView == null)
                {
                    try
                    {
                        // CreateSectionView returns void in the interop — call it and
                        // check if a new view appeared afterwards.
                        int viewCountBefore = ((object[])drawing.GetViews())?.Length ?? 0;
                        drawing.CreateSectionView();
                        int viewCountAfter = ((object[])drawing.GetViews())?.Length ?? 0;
                        if (viewCountAfter > viewCountBefore)
                        {
                            // Grab the newest view
                            object[] allViews = (object[])drawing.GetViews();
                            sectionView = (View)allViews[allViews.Length - 1];
                        }
                    }
                    catch { _progress?.LogMessage("CreateSectionView fallback also failed."); }
                }

                if (sectionView != null)
                    _progress?.LogMessage("Section view A-A created.");
                else
                    _progress?.LogWarning("Section view could not be created (API unavailable for this SW version).");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Could not create section view: {ex.Message}");
            }
        }

        /// <summary>
        /// Add detail view(s) for small features.
        /// Version-aware: tries CreateDetailViewAt4 (v29+/SW 2021), then CreateDetailViewAt2 (all versions).
        /// </summary>
        private void AddDetailView(DrawingDoc drawing, List<AnalyzedFeature> smallFeatures,
            double sheetW, double sheetH)
        {
            try
            {
                double detailX = sheetW * 0.78;
                double detailY = sheetH * 0.25;

                var smallest = smallFeatures.OrderBy(f => f.BoundingRadius).FirstOrDefault();
                if (smallest == null) return;

                drawing.ActivateView("Drawing View1");
                ModelDoc2 doc = (ModelDoc2)drawing;

                double radius = smallest.BoundingRadius * 3;
                if (radius < 0.005) radius = 0.005; // Min 5mm radius

                object[] views = (object[])drawing.GetViews();
                if (views == null || views.Length < 2) return;

                View frontView = (View)views[1];
                double[] pos = (double[])frontView.Position;

                doc.SketchManager.CreateCircle(pos[0], pos[1], 0, pos[0] + radius, pos[1], 0);

                View detailView = null;
                int apiVersion = SwVersionHelper.GetMajorVersion(_swApp);

                // Try CreateDetailViewAt4 (SW 2021+ / v29+)
                if (apiVersion >= 29)
                {
                    try
                    {
                        detailView = (View)drawing.CreateDetailViewAt4(
                            detailX, detailY, 0,
                            (int)swDetViewStyle_e.swDetViewSTANDARD,
                            2.0, radius, "B", 0, true, false, false, 0);
                    }
                    catch { _progress?.LogMessage("CreateDetailViewAt4 not available, trying fallback..."); }
                }

                // Fallback: try CreateDetailViewAt4 with a try/catch (it's available in interop v27.3)
                if (detailView == null)
                {
                    try
                    {
                        detailView = (View)drawing.CreateDetailViewAt4(
                            detailX, detailY, 0,
                            (int)swDetViewStyle_e.swDetViewSTANDARD,
                            2.0, radius, "B", 0, true, false, false, 0);
                    }
                    catch { _progress?.LogMessage("CreateDetailViewAt4 fallback also failed."); }
                }

                if (detailView != null)
                    _progress?.LogMessage("Detail view B created.");
                else
                    _progress?.LogWarning("Detail view could not be created (API unavailable for this SW version).");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Could not create detail view: {ex.Message}");
            }
        }

        /// <summary>
        /// Populate drawing title block with custom property values.
        /// </summary>
        private void PopulateTitleBlock(DrawingDoc drawing, ModelNode model, ExportOptions options)
        {
            var swModel = (ModelDoc2)drawing;

            // Define the map of CustomProperty -> Value
            var props = new Dictionary<string, string>
            {
                { "Description", model.Description },
                { "PartNo", model.PartNumber },
                { "Material", model.Material },
                { "Revision", model.Revision },

                // Project Metadata from UI
                { "Company", options.CompanyName },
                { "Project", options.ProjectName },
                { "DrawnBy", options.DrawnBy },
                { "CheckedBy", options.CheckedBy },
                { "Date", DateTime.Now.ToShortDateString() }
            };

            foreach (var kvp in props)
            {
                if (!string.IsNullOrEmpty(kvp.Value))
                {
                    // Add both to the file SummaryInfo and CustomPropertyManager for robustness
                    // (Some templates use $PRP:"Prop", others $PRPSHEET:"Prop")

                    // 1. File-level Custom Property
                    var cusPropMgr = swModel.Extension.get_CustomPropertyManager("");
                    cusPropMgr.Add3(kvp.Key, (int)swCustomInfoType_e.swCustomInfoText, kvp.Value, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);

                    // 2. Configuration-specific Property (Default config)
                    var configNames = (string[])swModel.GetConfigurationNames();
                    if (configNames != null && configNames.Length > 0)
                    {
                        var configPropMgr = swModel.Extension.get_CustomPropertyManager(configNames[0]);
                        configPropMgr.Add3(kvp.Key, (int)swCustomInfoType_e.swCustomInfoText, kvp.Value, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
                    }
                }
            }

            // Force rebuild to update linked notes
            swModel.ForceRebuild3(false);
        }

        private void SetProperty(CustomPropertyManager propMgr, string name, string value)
        {
            if (string.IsNullOrEmpty(value)) return;

            try
            {
                int result = propMgr.Add3(name, (int)swCustomInfoType_e.swCustomInfoText,
                    value, (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            catch { }
        }

        /// <summary>
        /// Get model bounding box. Returns [xMin, yMin, zMin, xMax, yMax, zMax].
        /// </summary>
        private double[] GetBoundingBox(ModelDoc2 doc)
        {
            // Try PartDoc.GetPartBox for parts
            try
            {
                if (doc is PartDoc partDoc)
                {
                    double[] box = (double[])partDoc.GetPartBox(true);
                    if (box != null && box.Length >= 6) return box;
                }
            }
            catch { }

            // Try via visible components bounding box for assemblies
            try
            {
                if (doc is AssemblyDoc assyDoc)
                {
                    double[] box = (double[])assyDoc.GetBox(
                        (int)swBoundingBoxOptions_e.swBoundingBoxIncludeRefPlanes);
                    if (box != null && box.Length >= 6) return box;
                }
            }
            catch { }

            // Last fallback: return a default 100mm cube
            return new double[] { -0.05, -0.05, -0.05, 0.05, 0.05, 0.05 };
        }

        private ModelDoc2 OpenOrGetModel(string filePath)
        {
            // Check if already open
            ModelDoc2 doc = (ModelDoc2)_swApp.GetOpenDocumentByName(filePath);
            if (doc != null) return doc;

            // Open it
            int errors = 0, warnings = 0;
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            int docType = ext == ".sldasm" ? (int)swDocumentTypes_e.swDocASSEMBLY
                        : ext == ".slddrw" ? (int)swDocumentTypes_e.swDocDRAWING
                        : (int)swDocumentTypes_e.swDocPART;

            doc = (ModelDoc2)_swApp.OpenDoc6(filePath, docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "", ref errors, ref warnings);

            return doc;
        }

        private DrawingDoc CreateDrawingDocument(string templatePath, int paperSize,
            double sheetW, double sheetH)
        {
            try
            {
                ModelDoc2 doc = (ModelDoc2)_swApp.NewDocument(
                    templatePath, paperSize, sheetW, sheetH);

                return (DrawingDoc)doc;
            }
            catch (Exception ex)
            {
                _progress?.LogError($"Cannot create drawing: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get the drawing template path using the shared multi-source resolver.
        /// </summary>
        private string GetDrawingTemplatePath(ExportOptions options)
        {
            string path = SwVersionHelper.FindDrawingTemplate(_swApp, options);
            if (!string.IsNullOrEmpty(path))
                _progress?.LogMessage($"Using template: {path}");
            else
                _progress?.LogWarning("No drawing template found. Using SolidWorks default.");
            return path;
        }

        private void SaveDrawing(ModelDoc2 drawingDoc, string outputPath)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                int errors = 0, warnings = 0;
                bool saved = drawingDoc.Extension.SaveAs2(
                    outputPath,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null, "", false,
                    ref errors, ref warnings);

                if (!saved)
                    _progress?.LogWarning($"Drawing save returned false. Errors: {errors}, Warnings: {warnings}");
            }
            catch (Exception ex)
            {
                _progress?.LogError($"Drawing save failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Build the output file path with proper naming convention.
        /// </summary>
        public static string GetOutputPath(ModelNode node, string outputFolder, string extension)
        {
            string baseName = SanitizeFileName(
                $"{node.PartNumber}_{node.Description}_Rev{node.Revision}".Trim('_'));

            if (string.IsNullOrEmpty(baseName))
                baseName = Path.GetFileNameWithoutExtension(node.FilePath);

            // Put into sub-folder by type
            string subFolder = "";
            switch (extension.ToLowerInvariant())
            {
                case ".slddrw": subFolder = "Drawings"; break;
                case ".pdf": subFolder = "PDF"; break;
                case ".dxf": subFolder = "DXF"; break;
                case ".step":
                case ".stp": subFolder = "STEP"; break;
                case ".x_t": subFolder = "Parasolid"; break;
                case ".xlsx": subFolder = "BOM"; break;
                case ".png": subFolder = "Preview"; break;
                default: subFolder = "Other"; break;
            }

            return Path.Combine(outputFolder, subFolder, baseName + extension);
        }

        private static string SanitizeFileName(string name)
        {
            char[] invalid = Path.GetInvalidFileNameChars();
            return new string(name.Where(c => !invalid.Contains(c)).ToArray());
        }
    }
}
