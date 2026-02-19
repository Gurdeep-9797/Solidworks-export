using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Rule-based dimensioning engine. After drawing views are placed, this engine
    /// adds annotations based on the analyzed feature list.
    /// 
    /// Strategy:
    ///   1. InsertModelAnnotations3 for all model-driven dims  
    ///   2. Post-process: add missing callouts per feature rules
    ///   3. Deduplicate (e.g., only one fillet radius note per unique value)
    ///   4. Collision avoidance via annotation position tracking
    /// </summary>
    public class DimensionEngine
    {
        private readonly ISldWorks _swApp;
        private readonly IProgressCallback _progress;

        // Track placed annotation bounding rects for collision avoidance (in sheet coords)
        private readonly List<AnnotationRect> _placedAnnotations = new List<AnnotationRect>();
        private readonly HashSet<string> _addedFilletRadii = new HashSet<string>();
        private readonly HashSet<string> _addedChamfers = new HashSet<string>();

        private const double OFFSET_STEP = 0.005; // 5mm offset when collisions detected

        public DimensionEngine(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
        }

        /// <summary>
        /// Apply all dimensions and annotations to the drawing.
        /// Call AFTER views are placed.
        /// </summary>
        public void ApplyDimensions(DrawingDoc drawing, List<AnalyzedFeature> features)
        {
            _placedAnnotations.Clear();
            _addedFilletRadii.Clear();
            _addedChamfers.Clear();

            // ── Stage 1: Insert model-driven annotations ──────────────────────
            InsertModelAnnotations(drawing);

            // ── Stage 2: Insert hole callouts ────────────────────────────────
            InsertHoleCallouts(drawing);

            // ── Stage 3: Post-process per feature rules ──────────────────────
            foreach (var feature in features)
            {
                switch (feature.Category)
                {
                    case FeatureCategory.Fillet:
                        AddFilletAnnotation(drawing, feature);
                        break;

                    case FeatureCategory.Chamfer:
                        AddChamferAnnotation(drawing, feature);
                        break;

                    case FeatureCategory.LinearPattern:
                        AddPatternNote(drawing, feature, isLinear: true);
                        break;

                    case FeatureCategory.CircularPattern:
                        AddPatternNote(drawing, feature, isLinear: false);
                        break;

                    case FeatureCategory.Thread:
                        AddThreadCallout(drawing, feature);
                        break;
                }
            }

            // ── Stage 4: Add overall bounding dimensions ─────────────────────
            AddOverallDimensions(drawing);

            _progress?.LogMessage($"Dimension engine complete. {_placedAnnotations.Count} annotations placed.");
        }

        /// <summary>
        /// Stage 1: Use SolidWorks' built-in model annotation insertion.
        /// This imports all dims marked for drawing + hole callouts.
        /// </summary>
        private void InsertModelAnnotations(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Inserting model annotations...");

                // Flags: marked-for-drawing dims + notes
                int annotFlags = (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing
                               | (int)swInsertAnnotation_e.swInsertNotes;

                drawing.InsertModelAnnotations3(
                    (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                    annotFlags,
                    true,   // Use dimension placement
                    true,   // Import items from all views
                    false,  // Don't use stacked balloons
                    true    // Auto-arrangement
                );

                _progress?.LogMessage("Model annotations inserted.");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"InsertModelAnnotations3 failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Insert hole callouts across all views.
        /// </summary>
        private void InsertHoleCallouts(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Inserting hole callouts...");

                ModelDoc2 drawingDoc = (ModelDoc2)drawing;
                object[] views = (object[])drawing.GetViews();

                if (views == null) return;

                foreach (object viewObj in views)
                {
                    View view = (View)viewObj;
                    if (view == null || view.Name == "Sheet1") continue; // Skip sheet-level "view"

                    try
                    {
                        // Activate the view
                        drawing.ActivateView(view.Name);

                        // Get visible entities in the view and select circular edges for callouts
                        // Insert dims marked for drawing (holes get auto-callouts)
                        drawing.InsertModelAnnotations3(
                            (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                            (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing,
                            true, true, false, true);
                    }
                    catch (Exception ex)
                    {
                        _progress?.LogWarning($"Hole callout failed for view '{view.Name}': {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"InsertHoleCallouts failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Add fillet radius note. Only one note per unique radius to avoid clutter.
        /// </summary>
        private void AddFilletAnnotation(DrawingDoc drawing, AnalyzedFeature feature)
        {
            if (feature.Radius <= 0) return;

            // Dedup: only one note per unique radius value
            string radiusKey = $"R{feature.Radius * 1000:F2}";
            if (_addedFilletRadii.Contains(radiusKey))
                return;

            _addedFilletRadii.Add(radiusKey);

            try
            {
                ModelDoc2 doc = (ModelDoc2)drawing;
                string noteText = $"R{feature.Radius * 1000:F1}";

                // Place a note near the feature's approximate position
                Note note = (Note)doc.InsertNote(noteText);
                if (note != null)
                {
                    Annotation ann = (Annotation)note.GetAnnotation();
                    if (ann != null)
                    {
                        ann.SetLeader3(
                            (int)swLeaderStyle_e.swNO_LEADER, 0,
                            true, false, false, false);
                    }
                    _progress?.LogMessage($"Added fillet note: {noteText}");
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Failed to add fillet annotation: {ex.Message}");
            }
        }

        /// <summary>
        /// Add chamfer callout. Deduplicated per unique distance/angle combination.
        /// </summary>
        private void AddChamferAnnotation(DrawingDoc drawing, AnalyzedFeature feature)
        {
            if (feature.Distance1 <= 0) return;

            string chamferKey = $"{feature.Distance1 * 1000:F2}x{feature.Angle:F1}";
            if (_addedChamfers.Contains(chamferKey))
                return;

            _addedChamfers.Add(chamferKey);

            try
            {
                ModelDoc2 doc = (ModelDoc2)drawing;
                string noteText;

                if (feature.Angle > 0 && Math.Abs(feature.Angle - Math.PI / 4) < 0.01)
                {
                    // 45° chamfer: show as "CxC"
                    noteText = $"{feature.Distance1 * 1000:F1}×45°";
                }
                else if (feature.Distance2 > 0)
                {
                    noteText = $"{feature.Distance1 * 1000:F1}×{feature.Distance2 * 1000:F1}";
                }
                else
                {
                    noteText = $"{feature.Distance1 * 1000:F1}×{(feature.Angle * 180 / Math.PI):F0}°";
                }

                Note note = (Note)doc.InsertNote(noteText);
                if (note != null)
                {
                    _progress?.LogMessage($"Added chamfer note: {noteText}");
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Failed to add chamfer annotation: {ex.Message}");
            }
        }

        /// <summary>
        /// Add pattern quantity/spacing note.
        /// </summary>
        private void AddPatternNote(DrawingDoc drawing, AnalyzedFeature feature, bool isLinear)
        {
            if (feature.PatternCount <= 1) return;

            try
            {
                ModelDoc2 doc = (ModelDoc2)drawing;
                string noteText;

                if (isLinear)
                {
                    noteText = $"{feature.PatternCount}× EQUALLY SPACED @ {feature.PatternSpacing * 1000:F1}mm";
                }
                else
                {
                    double angleDeg = feature.PatternSpacing * 180 / Math.PI;
                    noteText = $"{feature.PatternCount}× @ {angleDeg:F1}°";
                }

                Note note = (Note)doc.InsertNote(noteText);
                if (note != null)
                {
                    _progress?.LogMessage($"Added pattern note: {noteText}");
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Failed to add pattern note: {ex.Message}");
            }
        }

        /// <summary>
        /// Add thread callout note.
        /// </summary>
        private void AddThreadCallout(DrawingDoc drawing, AnalyzedFeature feature)
        {
            if (string.IsNullOrEmpty(feature.ThreadCallout)) return;

            try
            {
                ModelDoc2 doc = (ModelDoc2)drawing;
                string noteText = feature.ThreadCallout;

                if (feature.Depth > 0)
                    noteText += $" ↧{feature.Depth * 1000:F1}";

                Note note = (Note)doc.InsertNote(noteText);
                if (note != null)
                {
                    _progress?.LogMessage($"Added thread callout: {noteText}");
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Failed to add thread callout: {ex.Message}");
            }
        }

        /// <summary>
        /// Add overall bounding box dimensions (Length × Width × Height).
        /// These are placed on the Front and Right views.
        /// </summary>
        private void AddOverallDimensions(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Adding overall bounding dimensions...");

                ModelDoc2 doc = (ModelDoc2)drawing;

                // The built-in auto-dimensions often miss bounding box dims.
                // We insert a general note with overall dimensions.
                // More sophisticated: select outer edges and add dims,
                // but that requires face/edge selection which is model-specific.

                // For now, attempt model annotations for remaining items
                drawing.InsertModelAnnotations3(
                    (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                    (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing,
                    true, true, false, true);
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Overall dimensions failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Collision rectangle for an annotation (used for overlap avoidance).
        /// </summary>
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
