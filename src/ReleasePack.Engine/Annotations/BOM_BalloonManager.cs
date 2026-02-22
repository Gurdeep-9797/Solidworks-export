using System;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine.Annotations
{
    /// <summary>
    /// V2 Intelligent Auto-Ballooning and BOM Formatting rules engine.
    /// Radially places balloons around assembly views and explicitly enforces geometry logic
    /// based on digit lengths (e.g. >= 3 digits forces Hexagonal).
    /// </summary>
    public class BOM_BalloonManager
    {
        private readonly ModelDoc2 _drawing;

        public BOM_BalloonManager(ModelDoc2 drawing)
        {
            _drawing = drawing;
        }

        /// <summary>
        /// Applies automated balloons to a specific assembly view following strict rules.
        /// </summary>
        public void ApplyBOM_Balloons(IView assemblyView)
        {
            if (_drawing == null || assemblyView == null) return;

            DrawingDoc draw = (DrawingDoc)_drawing;
            draw.ActivateView(assemblyView.Name);

            // 1. Initial Auto-Ballooning using Square layout to ensure they ring the view
            BalloonOptions bOptions = _drawing.Extension.CreateBalloonOptions();
            bOptions.Style = (int)swBalloonStyle_e.swBS_Circular; // Start generic
            bOptions.Size = (int)swBalloonFit_e.swBF_2Chars;
            bOptions.UpperTextContent = (int)swBalloonTextContent_e.swBalloonTextItemNumber;

            dynamic dynView = assemblyView;
            var resultBalloons = dynView.AutoBalloon5(bOptions);
            
            if (resultBalloons == null) return;

            // 2. Iterate and enforce digit-length shape logic + Layer assignment
            INote note = (INote)assemblyView.GetFirstNote();
            while (note != null)
            {
                if (note.IsBomBalloon())
                {
                    string text = note.GetText();
                    
                    // Enforce digit length shape rules
                    if (!string.IsNullOrEmpty(text))
                    {
                        // Some balloons text may contain formatting. Strip if necessary, or just measure char length.
                        int pureDigits = ExtractDigits(text).Length;

                        if (pureDigits >= 3)
                        {
                            // Shift to Hexagonal
                            note.SetBalloon((int)swBalloonStyle_e.swBS_Hexagon, (int)swBalloonFit_e.swBF_3Chars);
                        }
                        else
                        {
                            // Keep circular but ensure tight fit
                            note.SetBalloon((int)swBalloonStyle_e.swBS_Circular, (int)swBalloonFit_e.swBF_Tightest);
                        }
                    }

                    // Push to the BOM layer
                    IAnnotation ann = (IAnnotation)note.GetAnnotation();
                    if (ann != null)
                    {
                        // Dynamic dispatch for pushing to layer since Annotation vs IAnnotation differs between V27 headers
                        dynamic dynAnn = ann;
                        Layout.LayerManager.PushToLayer(dynAnn, Layout.LayerManager.LAYER_BOM);
                    }
                }
                note = (INote)note.GetNext();
            }
        }

        private string ExtractDigits(string input)
        {
            // Simple generic numeric extraction (removes SW internal tags if present)
            string result = string.Empty;
            foreach (char c in input)
            {
                if (char.IsDigit(c)) result += c;
            }
            return result;
        }
    }
}
