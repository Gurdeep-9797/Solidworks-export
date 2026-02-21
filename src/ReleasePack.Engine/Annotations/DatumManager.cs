using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;

namespace ReleasePack.Engine.Annotations
{
    /// <summary>
    /// Intelligently applies datum feature symbols to the geometric view
    /// and strictly enforces the 'Datums' layout layer coloring.
    /// </summary>
    public class DatumManager
    {
        private readonly ModelDoc2 _drawing;

        public DatumManager(ModelDoc2 drawing)
        {
            _drawing = drawing;
        }

        /// <summary>
        /// Inserts a Datum reference (e.g. [A], [B]) on the provided edge or coordinate
        /// and aggressively forces it to the designated layer.
        /// </summary>
        public DatumFeatureSymbol InsertDatum(IView targetView, string label, double x, double y)
        {
            if (_drawing == null || targetView == null) return null;

            // Ensure the correct view is active
            DrawingDoc draw = (DrawingDoc)_drawing;
            draw.ActivateView(targetView.Name);

            // Select an arbitrary point to attach the datum to (usually an edge mid-point or similar)
            _drawing.Extension.SelectByID2("", "FACE", x, y, 0, false, 0, null, 0);

            // Insert Datum
            DatumFeatureSymbol datum = (DatumFeatureSymbol)_drawing.Extension.InsertDatumTag(label);
            if (datum != null)
            {
                IAnnotation ann = datum.GetAnnotation();
                if (ann != null)
                {
                    // Force the standard ISO/ANSI style and push to Layer
                    Layout.LayerManager.PushToLayer(ann, Layout.LayerManager.LAYER_DATUMS);
                }
            }
            
            _drawing.ClearSelection2(true);
            return datum;
        }
    }
}
