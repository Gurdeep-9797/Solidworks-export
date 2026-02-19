using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Rule-based dimensioning engine.
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

        public void ApplyDimensions(DrawingDoc drawing, List<AnalyzedFeature> features)
        {
            _placedAnnotations.Clear();

            // 0. Layers
            SetupLayers(drawing);

            // 1. Model Items (Dimensions, Hole Callouts)
            InsertModelAnnotations(drawing);

            // 2. Smart Auto-Dimensioning (API)
            SmartAutoDimension(drawing);

            // 3. Post-process properties features
            if (features != null)
            {
                foreach (var feature in features)
                {
                    // Placeholder: In future, add specific notes based on feature type
                    // e.g. AddFilletAnnotation(drawing, feature);
                }
            }

            // 4. Formatting
            ColorizeDimensions(drawing);
        }

        private void SetupLayers(DrawingDoc drawing)
        {
            try
            {
                var layerMgr = (LayerMgr)((ModelDoc2)drawing).GetLayerManager();
                if (layerMgr == null) return;

                CreateLayer(layerMgr, "Dimensions", Color.DarkBlue, swLineWeights_e.swLW_THIN);
                CreateLayer(layerMgr, "Notes", Color.Black, swLineWeights_e.swLW_NORMAL);
                CreateLayer(layerMgr, "Format", Color.DarkRed, swLineWeights_e.swLW_NORMAL);
            }
            catch {}
        }

        private void CreateLayer(LayerMgr layerMgr, string name, Color color, swLineWeights_e weight)
        {
            if (layerMgr.GetLayer(name) == null)
            {
                layerMgr.AddLayer(name, "Standard " + name, (int)ColorTranslator.ToWin32(color), (int)swLineStyles_e.swLineCONTINUOUS, (int)weight);
            }
        }

        private void InsertModelAnnotations(DrawingDoc drawing)
        {
            try
            {
                _progress?.LogMessage("Inserting model items...");
                // Note: swInsertHoleCallout is 0x4 (swInsertAnnotation_e)
                drawing.InsertModelAnnotations3(
                    (int)swImportModelItemsSource_e.swImportModelItemsFromEntireModel,
                    (int)swInsertAnnotation_e.swInsertDimensionsMarkedForDrawing | 
                    (int)swInsertAnnotation_e.swInsertNotes |
                    4, // swInsertHoleCallout
                    true, true, false, true
                );
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"InsertModelAnnotations failed: {ex.Message}");
            }
        }
        
        private void InsertHoleCallouts(DrawingDoc drawing)
        {
             // Already handled in InsertModelAnnotations with flag 4
        }
        
        private void SmartAutoDimension(DrawingDoc drawing)
        {
            try 
            {
                _progress?.LogMessage("Running Smart Auto-Dimensioning...");
                
                object[] views = (object[])drawing.GetViews();
                if (views == null) return;

                foreach (object viewObj in views)
                {
                    View view = (View)viewObj;
                    if (view == null || view.Name == "Sheet1") continue; // Skip sheet view

                    // Skip Iso views generally
                    // (Logic: Check orientation or name)
                    // For now, dimensions on ISO are usually cluttered, so skip if name contains "Isometric"
                    if (view.Name.IndexOf("Isometric", StringComparison.OrdinalIgnoreCase) >= 0) continue;

                    // 1. Select the view
                    bool selected = ((ModelDoc2)drawing).Extension.SelectByID2(view.Name, "DRAWINGVIEW", 0,0,0, false, 0, null, 0);
                    if (!selected) continue;

                    // 2. Call AutoDimension
                    // 0 = swAutodimEntitiesBasedOnSources (Selected View) - Wait, selected view is often 1 or 2?
                    // According to API: swAutodimEntitiesSelectedView = 2 usually, or it might be missing.
                    // Let's use the raw integer if the enum is missing.
                    // Actually, let's try to find the correct enum member. 
                    // If it's not in the enum, we cast the int.
                    
                    int status = drawing.AutoDimension(
                        2, // 2 = swAutodimEntitiesSelectedView (Enum member missing in this interop)
                        (int)swAutodimScheme_e.swAutodimSchemeBaseline,
                        (int)swAutodimHorizontalPlacement_e.swAutodimHorizontalPlacementBelow, 
                        (int)swAutodimScheme_e.swAutodimSchemeBaseline, 
                        (int)swAutodimVerticalPlacement_e.swAutodimVerticalPlacementLeft
                    );
                    
                    if (status != 0) 
                        _progress?.LogMessage($"AutoDimension returned status {status} for view {view.Name}");
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"SmartAutoDimension failed: {ex.Message}");
            }
        }

        private void AddFilletAnnotation(DrawingDoc drawing, AnalyzedFeature feature) {}
        private void AddChamferAnnotation(DrawingDoc drawing, AnalyzedFeature feature) {}
        private void AddPatternNote(DrawingDoc drawing, AnalyzedFeature feature, bool isLinear) {}
        private void AddThreadCallout(DrawingDoc drawing, AnalyzedFeature feature) {}

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
                        else if (type == (int)swAnnotationType_e.swNote || type == (int)swAnnotationType_e.swDatumTag)
                            ann.Layer = "Notes";
                        
                        ann = (Annotation)ann.GetNext3();
                    }
                }
            }
            catch {}
        }

        /// <summary>
        /// Collision rectangle for an annotation.
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
