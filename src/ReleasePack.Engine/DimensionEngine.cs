using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Professional annotation engine implementing industrial drawing standards.
    /// 
    /// Annotation pipeline:
    ///   1. Setup drawing layers (Dimensions, Notes, Format)
    ///   2. Insert Model Items (driving dimensions, hole callouts, notes)
    ///   3. Insert Center Marks & Centerlines on all circular geometry
    ///   4. Smart Auto-Dimensioning (baseline scheme)
    ///   5. Pattern Labels (4X EQ SP, PCD, etc.)
    ///   6. Thread Callouts (M8x1.25)
    ///   7. Fillet/Chamfer TYP Notes (R3 TYP, 2x45° TYP)
    ///   8. Colorize annotations to correct layers
    ///   9. Remove overlapping/duplicate dimensions
    /// </summary>
    public class DimensionEngine
    {
        private readonly ISldWorks _swApp;
        private readonly IProgressCallback _progress;

        // Track placed annotation bounding rects for collision avoidance
        private readonly List<AnnotationRect> _placedAnnotations = new List<AnnotationRect>();

        public DimensionEngine(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
        }

        /// <summary>
        /// Full annotation pipeline for a drawing.
        /// </summary>
        public void ApplyDimensions(DrawingDoc drawing, List<AnalyzedFeature> features)
        {
            _placedAnnotations.Clear();

            // 0. Layers
            SetupLayers(drawing);

            // 1. Model Items (Dimensions, Hole Callouts, Notes)
            InsertModelAnnotations(drawing);

            // 2. Center Marks & Centerlines on ALL views
            InsertCenterMarks(drawing);

            // 3. Native SolidWorks Auto-Arrange Dimensions
            AutoArrangeDimensions(drawing);

            // 4. Feature-specific annotations
            if (features != null)
            {
                // Pattern labels
                foreach (var feature in features.Where(f => 
                    f.Category == FeatureCategory.LinearPattern || 
                    f.Category == FeatureCategory.CircularPattern))
                {
                    AddPatternNote(drawing, feature, 
                        feature.Category == FeatureCategory.LinearPattern);
                }

                // Thread callouts
                foreach (var feature in features.Where(f => 
                    f.Category == FeatureCategory.Thread && !string.IsNullOrEmpty(f.ThreadCallout)))
                {
                    AddThreadCallout(drawing, feature);
                }

                // Fillet TYP notes (only label first in each TYP group)
                var labeledGroups = new HashSet<string>();
                foreach (var feature in features.Where(f => f.IsTypical && f.Category == FeatureCategory.Fillet))
                {
                    if (feature.TypicalGroupId != null && !labeledGroups.Contains(feature.TypicalGroupId))
                    {
                        AddFilletAnnotation(drawing, feature);
                        labeledGroups.Add(feature.TypicalGroupId);
                    }
                }

                // Chamfer TYP notes
                foreach (var feature in features.Where(f => f.IsTypical && f.Category == FeatureCategory.Chamfer))
                {
                    if (feature.TypicalGroupId != null && !labeledGroups.Contains(feature.TypicalGroupId))
                    {
                        AddChamferAnnotation(drawing, feature);
                        labeledGroups.Add(feature.TypicalGroupId);
                    }
                }
            }

            // 5. Formatting & Cleanup
            ColorizeDimensions(drawing);
            CleanupOverlappingDimensions(drawing);
        }

        // ======================================================================
        // Layer Setup
        // ======================================================================

        private void SetupLayers(DrawingDoc drawing)
        {
            try
            {
                var layerMgr = (LayerMgr)((ModelDoc2)drawing).GetLayerManager();
                if (layerMgr == null) return;

                CreateLayer(layerMgr, "Dimensions", Color.DarkBlue, swLineWeights_e.swLW_THIN);
                CreateLayer(layerMgr, "Notes", Color.Black, swLineWeights_e.swLW_NORMAL);
                CreateLayer(layerMgr, "Format", Color.DarkRed, swLineWeights_e.swLW_NORMAL);
                CreateLayer(layerMgr, "CenterMarks", Color.Gray, swLineWeights_e.swLW_THIN);
            }
            catch { }
        }

        private void CreateLayer(LayerMgr layerMgr, string name, Color color, swLineWeights_e weight)
        {
            if (layerMgr.GetLayer(name) == null)
            {
                layerMgr.AddLayer(name, "Standard " + name,
                    (int)ColorTranslator.ToWin32(color),
                    (int)swLineStyles_e.swLineCONTINUOUS,
                    (int)weight);
            }
        }

        // ======================================================================
        // 1. Model Items (driving dims, hole callouts, notes)
        // ======================================================================

        private void InsertModelAnnotations(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Inserting model items (dimensions, hole callouts, notes)...");

                // Bitwise flags for annotation types:
                // swInsertDimensionsMarkedForDrawing = 0x1
                // swInsertNotes = 0x2
                // swInsertHoleCallout = 0x4
                int annotTypes = (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing |
                                 (int)swInsertAnnotation_e.swInsertNotes |
                                 4; // swInsertHoleCallout (not in all versions of enum)

                drawing.InsertModelAnnotations3(
                    (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                    annotTypes,
                    true,   // allViews
                    true,   // duplicateDims
                    false,  // hiddenFeatureDims
                    true    // usePlacementInSketch
                );

                _progress?.LogMessage("Model items inserted.");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"InsertModelAnnotations failed: {ex.Message}");
            }
        }

        // ======================================================================
        // 2. Center Marks & Centerlines
        // ======================================================================

        /// <summary>
        /// Insert center marks on ALL circular/arc edges in every non-Isometric view.
        /// Uses IView.AutoInsertCenterMarks API which handles:
        /// - Hole center marks (crosshair)
        /// - Cylindrical centerlines (axis lines)
        /// - Bolt circle centerlines
        /// - Arc center marks
        /// </summary>
        private void InsertCenterMarks(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Inserting center marks and centerlines...");

                object[] sheets = (object[])drawing.GetViews();
                if (sheets == null) return;

                int marksInserted = 0;

                foreach (object sheetObj in sheets)
                {
                    // GetViews returns an array of arrays: first level is sheets, each contains views
                    object[] views = null;
                    if (sheetObj is object[] viewArray)
                    {
                        views = viewArray;
                    }
                    else if (sheetObj is View singleView)
                    {
                        // It's a flat list — use the drawing's sheet views
                        views = (object[])drawing.GetViews();
                        // Process flat list
                        foreach (object vObj in views)
                        {
                            View v = vObj as View;
                            if (v == null) continue;
                            if (IsSheetOrIsoView(v)) continue;

                            InsertCenterMarksOnView(drawing, v, ref marksInserted);
                        }
                        break; // Already processed all views
                    }

                    if (views == null) continue;

                    foreach (object viewObj in views)
                    {
                        View view = viewObj as View;
                        if (view == null) continue;
                        if (IsSheetOrIsoView(view)) continue;

                        InsertCenterMarksOnView(drawing, view, ref marksInserted);
                    }
                }

                _progress?.LogMessage($"Center marks inserted ({marksInserted} views processed).");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Center marks failed: {ex.Message}");
            }
        }

        private void InsertCenterMarksOnView(DrawingDoc drawing, View view, ref int count)
        {
            try
            {
                ModelDoc2 doc = (ModelDoc2)drawing;
                // Select the view first
                doc.Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                // InsertCenterMark3 inserts marks on selected view's circular edges
                // Alternatively, we try to use the auto-insert approach
                // API: IDrawingDoc::InsertCenterMark3(UseDocDefaults, SizeOrOptions...)
                
                // A simpler and more reliable approach: use InsertModelAnnotations
                // with center mark flags. But InsertModelAnnotations3 doesn't always include marks.
                
                // Best approach: directly call InsertCenterMark3 which places center marks
                // on the currently selected edges. We need to select circular edges first.
                // For automation, we iterate visible edges and select circles/arcs.

                // Fallback: The simplest programmatic way is InsertModelAnnotations3
                // with the centermark flag. But since that may already be called, let's
                // try the drawing.InsertCenterMark3 approach on auto-detected edges.
                
                // For maximum compatibility, use the "select all edges in view" approach:
                drawing.InsertCenterMark3(0, true, true);
                count++;
            }
            catch
            {
                // Center marks may fail on views with no circular geometry — that's expected
            }
        }

        private bool IsSheetOrIsoView(View view)
        {
            if (view.Name.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase))
                return true;
            if (view.Name.IndexOf("Isometric", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
            if (view.Name.IndexOf("Iso", StringComparison.OrdinalIgnoreCase) >= 0 &&
                view.Name.Length <= 5)
                return true;
            return false;
        }

        // ======================================================================
        // 3. Smart Auto Arrange
        // ======================================================================

        private void AutoArrangeDimensions(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Running SolidWorks Auto-Arrange Dimensions...");

                object[] views = (object[])drawing.GetViews();
                if (views == null) return;

                ModelDocExtension ext = ((ModelDoc2)drawing).Extension;

                // Process first sheet
                object[] sheetViews = null;
                if (views[0] is object[] sViews)
                    sheetViews = sViews;
                else
                    sheetViews = views; // Flattened array

                foreach (object viewObj in sheetViews)
                {
                    View view = (View)viewObj;
                    if (view == null || IsSheetOrIsoView(view)) continue;

                    _progress?.LogMessage($"Auto-arranging dimensions for view '{view.Name}'");

                    ext.SelectByID2(view.Name, "DRAWINGVIEW", 0, 0, 0, false, 0, null, 0);

                    // Replaces the basic algorithmic bounds staggering with SW Native AutoArrange
                    ext.AlignDimensions((int)swAlignDimensionType_e.swAlignDimensionType_AutoArrange, 0.0);
                }
                
                // Clear selection
                ((ModelDoc2)drawing).ClearSelection2(true);
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"AutoArrange failed: {ex.Message}");
            }
        }

        // ======================================================================
        // 4. Pattern Labels
        // ======================================================================

        /// <summary>
        /// Add pattern label notes following engineering standards:
        /// - Linear: "4X @ 25" or "3X EQ SP"
        /// - Circular: "6X EQ SP ON Ø50 PCD"
        /// </summary>
        private void AddPatternNote(DrawingDoc drawing, AnalyzedFeature feature, bool isLinear)
        {
            try
            {
                if (feature.PatternCount <= 0) return;

                string noteText;
                if (isLinear)
                {
                    double spacingMm = feature.PatternSpacing * 1000; // Convert m → mm
                    if (spacingMm > 0)
                        noteText = $"{feature.PatternCount}X @ {spacingMm:F1}";
                    else
                        noteText = $"{feature.PatternCount}X EQ SP";
                }
                else
                {
                    // Circular pattern — include PCD if available
                    double spacingDeg = feature.PatternSpacing * (180.0 / Math.PI); // radians → degrees

                    if (feature.Diameter > 0)
                    {
                        double pcdMm = feature.Diameter * 1000;
                        noteText = $"{feature.PatternCount}X EQ SP ON <MOD-DIAM>{pcdMm:F1} PCD";
                    }
                    else if (Math.Abs(spacingDeg) > 0.01)
                    {
                        noteText = $"{feature.PatternCount}X @ {spacingDeg:F0}°";
                    }
                    else
                    {
                        noteText = $"{feature.PatternCount}X EQ SP";
                    }
                }

                // Insert a note at the feature's approximate position
                InsertNoteAtFeature(drawing, feature, noteText);

                _progress?.LogMessage($"Pattern label: {noteText} for '{feature.Name}'");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Pattern note failed for '{feature.Name}': {ex.Message}");
            }
        }

        // ======================================================================
        // 5. Thread Callouts
        // ======================================================================

        /// <summary>
        /// Add thread callout note: e.g., "M8x1.25 - 6H" or "M10x1.5"
        /// </summary>
        private void AddThreadCallout(DrawingDoc drawing, AnalyzedFeature feature)
        {
            try
            {
                string callout = feature.ThreadCallout;
                if (string.IsNullOrEmpty(callout)) return;

                // Depth suffix if available
                if (feature.Depth > 0)
                {
                    double depthMm = feature.Depth * 1000;
                    callout += $" x {depthMm:F0} DEEP";
                }

                InsertNoteAtFeature(drawing, feature, callout);
                _progress?.LogMessage($"Thread callout: {callout} for '{feature.Name}'");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Thread callout failed for '{feature.Name}': {ex.Message}");
            }
        }

        // ======================================================================
        // 6. Fillet/Chamfer TYP Notes
        // ======================================================================

        /// <summary>
        /// Add "R{x} TYP" note for fillets that share the same radius.
        /// Only called for the FIRST feature in each TYP group.
        /// </summary>
        private void AddFilletAnnotation(DrawingDoc drawing, AnalyzedFeature feature)
        {
            try
            {
                double radiusMm = feature.Radius * 1000;
                if (radiusMm <= 0) return;

                string note = $"R{radiusMm:F1} TYP";
                InsertNoteAtFeature(drawing, feature, note);
                _progress?.LogMessage($"Fillet TYP: {note} for group '{feature.TypicalGroupId}'");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Fillet TYP note failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Add "CxD TYP" note for chamfers (e.g., "2x45° TYP").
        /// </summary>
        private void AddChamferAnnotation(DrawingDoc drawing, AnalyzedFeature feature)
        {
            try
            {
                double sizeMm = feature.Distance1 * 1000;
                double angle = feature.Angle > 0 ? feature.Angle * (180.0 / Math.PI) : 45;
                if (sizeMm <= 0) return;

                string note = $"{sizeMm:F1}x{angle:F0}° TYP";
                InsertNoteAtFeature(drawing, feature, note);
                _progress?.LogMessage($"Chamfer TYP: {note} for group '{feature.TypicalGroupId}'");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Chamfer TYP note failed: {ex.Message}");
            }
        }

        // ======================================================================
        // Note Insertion Helper
        // ======================================================================

        /// <summary>
        /// Insert a note annotation at the feature's approximate position on the drawing.
        /// </summary>
        private void InsertNoteAtFeature(DrawingDoc drawing, AnalyzedFeature feature, string text)
        {
            try
            {
                ModelDoc2 doc = (ModelDoc2)drawing;

                // Place note at a reasonable position
                // If feature has position data, use it; otherwise use a default offset
                double x = feature.PositionX;
                double y = feature.PositionY;

                // If no position data, use a generic position near center of first view
                if (Math.Abs(x) < 1e-10 && Math.Abs(y) < 1e-10)
                {
                    // Get the first standard view to use as reference
                    object[] views = (object[])drawing.GetViews();
                    if (views != null)
                    {
                        foreach (object vObj in views)
                        {
                            View v = vObj as View;
                            if (v != null && !IsSheetOrIsoView(v))
                            {
                                double[] pos = (double[])v.Position;
                                if (pos != null && pos.Length >= 2)
                                {
                                    x = pos[0] + 0.02; // Offset 20mm right
                                    y = pos[1] + 0.02; // Offset 20mm up
                                }
                                break;
                            }
                        }
                    }
                }

                // Use InsertNote to place the annotation
                Note note = (Note)doc.InsertNote(text);
                if (note != null)
                {
                    Annotation ann = (Annotation)note.GetAnnotation();
                    if (ann != null)
                    {
                        ann.SetPosition2(x, y, 0);
                        ann.Layer = "Notes";
                    }
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"InsertNote failed for '{text}': {ex.Message}");
            }
        }

        // ======================================================================
        // 7. Colorize & Layer Assignment
        // ======================================================================

        private void ColorizeDimensions(DrawingDoc drawing)
        {
            try
            {
                object[] views = (object[])drawing.GetViews();
                if (views == null) return;

                foreach (object viewObj in views)
                {
                    View view = (View)viewObj;
                    Annotation ann = (Annotation)view.GetFirstAnnotation3();
                    while (ann != null)
                    {
                        int type = ann.GetType();
                        if (type == (int)swAnnotationType_e.swDisplayDimension)
                            ann.Layer = "Dimensions";
                        else if (type == (int)swAnnotationType_e.swNote ||
                                 type == (int)swAnnotationType_e.swDatumTag)
                            ann.Layer = "Notes";
                        else if (type == 16 || // swCenterMark
                                 type == 17)   // swCenterLine
                            ann.Layer = "CenterMarks";

                        ann = (Annotation)ann.GetNext3();
                    }
                }
            }
            catch { }
        }

        // ======================================================================
        // 8. Dimension Overlap Cleanup
        // ======================================================================

        /// <summary>
        /// Detect overlapping/nearly-overlapping dimension annotations and nudge them apart.
        /// </summary>
        private void CleanupOverlappingDimensions(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Cleaning up overlapping dimensions...");

                object[] views = (object[])drawing.GetViews();
                if (views == null) return;

                foreach (object viewObj in views)
                {
                    View view = (View)viewObj;
                    var dimAnnotations = new List<Annotation>();

                    Annotation ann = (Annotation)view.GetFirstAnnotation3();
                    while (ann != null)
                    {
                        if (ann.GetType() == (int)swAnnotationType_e.swDisplayDimension)
                        {
                            dimAnnotations.Add(ann);
                        }
                        ann = (Annotation)ann.GetNext3();
                    }

                    // Check for overlaps between each pair
                    for (int i = 0; i < dimAnnotations.Count; i++)
                    {
                        for (int j = i + 1; j < dimAnnotations.Count; j++)
                        {
                            try
                            {
                                double[] posI = (double[])dimAnnotations[i].GetPosition();
                                double[] posJ = (double[])dimAnnotations[j].GetPosition();

                                if (posI == null || posJ == null || posI.Length < 2 || posJ.Length < 2)
                                    continue;

                                // Check if positions are too close (within 5mm)
                                double dist = Math.Sqrt(
                                    Math.Pow(posI[0] - posJ[0], 2) +
                                    Math.Pow(posI[1] - posJ[1], 2));

                                if (dist < 0.005) // 5mm threshold
                                {
                                    // Nudge the second annotation
                                    dimAnnotations[j].SetPosition2(
                                        posJ[0],
                                        posJ[1] + 0.008, // Nudge 8mm
                                        0);
                                }
                            }
                            catch { }
                        }
                    }
                }

                _progress?.LogMessage("Dimension cleanup complete.");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Dimension cleanup failed: {ex.Message}");
            }
        }

        // ======================================================================
        // Internal Types
        // ======================================================================

        private class AnnotationRect
        {
            public double X { get; }
            public double Y { get; }
            public double Width { get; }
            public double Height { get; }

            public AnnotationRect(double x, double y, double width, double height)
            {
                X = x;
                Y = y;
                Width = width;
                Height = height;
            }

            public bool Overlaps(AnnotationRect other)
            {
                return !(X + Width < other.X || other.X + other.Width < X ||
                         Y + Height < other.Y || other.Y + other.Height < Y);
            }
        }
    }
}
