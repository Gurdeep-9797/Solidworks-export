using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Drawing;

namespace ReleasePack.Engine.Layout
{
    /// <summary>
    /// V2 AutoCAD-style strict Layer discipline manager.
    /// Injects robust drafting standards into the SolidWorks drawing environment.
    /// </summary>
    public class LayerManager
    {
        private readonly ModelDoc2 _drawingDoc;
        private readonly LayerMgr _layerMgr;

        // Named constants for explicitly allowed layers
        public const string LAYER_VISIBLE = "Visible";
        public const string LAYER_HIDDEN = "Hidden";
        public const string LAYER_CENTERLINES = "CenterLines";
        public const string LAYER_DIMENSIONS = "Dimensions";
        public const string LAYER_NOTES = "Notes";
        public const string LAYER_HATCHING = "Hatching";
        public const string LAYER_DATUMS = "Datums";
        public const string LAYER_CONSTRUCTION = "Construction";
        public const string LAYER_BOM = "BOM";
        public const string LAYER_TITLEBLOCK = "TitleBlock";

        public LayerManager(ModelDoc2 drawingDoc)
        {
            _drawingDoc = drawingDoc ?? throw new ArgumentNullException(nameof(drawingDoc));
            
            DrawingDoc draw = drawingDoc as DrawingDoc;
            if (draw == null) throw new InvalidOperationException("LayerManager requires a DrawingDocument.");

            _layerMgr = draw.GetCurrentLayerManager();
            if (_layerMgr == null) throw new Exception("Failed to access Layer Manager.");
        }

        public void InitializeLayers()
        {
            // The algorithm must not rely on document defaults - explicitly generate all core layers.
            // CreateLayer(name, desc, color, style, weightIndex)

            // Visible: Black, Continuous, 0.5mm (Index 4)
            CreateLayer(LAYER_VISIBLE, "Primary Part Geometry", Color.Black, swLineStyles_e.swLineCONTINUOUS, 4);

            // Hidden: Gray, Dashed, 0.25mm (Index 2)
            CreateLayer(LAYER_HIDDEN, "Unseen Geometry/Edges", Color.Gray, swLineStyles_e.swLineHIDDEN, 2);

            // CenterLines: Gray, Center (Chain), 0.18mm (Index 1)  [0.13mm mapped to closest index 1]
            CreateLayer(LAYER_CENTERLINES, "Symmetry/Bolt circles", Color.Gray, swLineStyles_e.swLineCENTER, 1);

            // Dimensions: Dark Blue, Continuous, 0.18mm (Index 1)
            CreateLayer(LAYER_DIMENSIONS, "All auto-dimensions", Color.DarkBlue, swLineStyles_e.swLineCONTINUOUS, 1);

            // Notes: Black, Continuous, 0.25mm (Index 2)
            CreateLayer(LAYER_NOTES, "Text, Callouts, Titles", Color.Black, swLineStyles_e.swLineCONTINUOUS, 2);

            // Hatching: Dark Gray, Continuous, 0.18mm (Index 1)
            CreateLayer(LAYER_HATCHING, "Section Cut Faces", Color.DarkGray, swLineStyles_e.swLineCONTINUOUS, 1);

            // Datums: Green, Continuous, 0.25mm (Index 2)
            CreateLayer(LAYER_DATUMS, "Reference Frames", Color.Green, swLineStyles_e.swLineCONTINUOUS, 2);

            // Construction: Light Green, Phantom, 0.18mm (Index 1)
            CreateLayer(LAYER_CONSTRUCTION, "Bounding Boxes/Guides", Color.LightGreen, swLineStyles_e.swLinePHANTOM, 1);

            // BOM: Black, Continuous, 0.25mm (Index 2)
            CreateLayer(LAYER_BOM, "Tables/Balloons", Color.Black, swLineStyles_e.swLineCONTINUOUS, 2);

            // TitleBlock: Dark Red, Continuous, 0.35mm (Index 3)
            CreateLayer(LAYER_TITLEBLOCK, "Borders/Logos", Color.DarkRed, swLineStyles_e.swLineCONTINUOUS, 3);
        }

        private void CreateLayer(string name, string desc, Color color, swLineStyles_e style, int weightClass)
        {
            // SW Color Ref: Color is packed as RGB integer = R + G*256 + B*65536
            int swColorInt = color.R + (color.G * 256) + (color.B * 65536);

            // Check if layer exists; if so, update it. If not, add it.
            Layer existingLayer = _layerMgr.GetLayer(name);

            if (existingLayer == null)
            {
                DrawingDoc draw = (DrawingDoc)_drawingDoc;
                draw.CreateLayer2(name, desc, swColorInt, (int)style, weightClass, true);
            }
            else
            {
                // Force conformity even if layer exists (e.g., loaded from a flawed template)
                existingLayer.Color = swColorInt;
                existingLayer.Style = (int)style;
                existingLayer.Width = weightClass;
                existingLayer.Description = desc;
            }
        }

        /// <summary>
        /// Explicit routing mechanism to place annotations firmly on specific layers bypassing SW defaults.
        /// </summary>
        public static void PushToLayer(Annotation swAnnotation, string activeLayer)
        {
            if (swAnnotation != null)
            {
                swAnnotation.Layer = activeLayer;
            }
        }

        public static void PushToLayer(IView swView, string activeLayer)
        {
            if (swView != null)
            {
                // Note: Changing view layer puts all internal entities to that layer unless overridden.
                swView.Layer = activeLayer;
            }
        }
    }
}
