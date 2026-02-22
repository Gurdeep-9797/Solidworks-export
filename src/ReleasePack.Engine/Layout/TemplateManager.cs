using System;
using System.Linq;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine.Layout
{
    /// <summary>
    /// V2 Strict Drawing Template & Sheet sizing engine.
    /// Manages paper size selection and enforces ISO 5455 valid scales.
    /// </summary>
    public class TemplateManager
    {
        private ISldWorks _swApp;
        
        // Strict ISO 5455 compliant scales.
        // Elements are standard fractional representations (e.g. 1/5 = 0.20)
        private static readonly double[] ISO_5455_SCALES = new double[]
        {
            50.0, 20.0, 10.0, 5.0, 2.0,                 // Enlargement
            1.0,                                        // Full Size
            1.0/2.0, 1.0/5.0, 1.0/10.0, 1.0/20.0,       // Reduction
            1.0/50.0, 1.0/100.0, 1.0/200.0, 1.0/500.0, 1.0/1000.0
        };

        public class StandardSheet
        {
            public string Name { get; set; }
            public double Width { get; set; } // in meters
            public double Height { get; set; } // in meters
            
            // Apply 10mm ISO boundary margins
            public double UsableWidth => Width - 0.020;
            public double UsableHeight => Height - 0.020;
            
            public int SwPaperSize { get; set; }
        }

        // Ordered by area (Smallest to Largest)
        public static readonly List<StandardSheet> ISO_SHEETS = new List<StandardSheet>
        {
            new StandardSheet { Name = "A4 Landscape", Width = 0.297, Height = 0.210, SwPaperSize = (int)swDwgPaperSizes_e.swDwgPaperA4size },
            new StandardSheet { Name = "A3 Landscape", Width = 0.420, Height = 0.297, SwPaperSize = (int)swDwgPaperSizes_e.swDwgPaperA3size },
            new StandardSheet { Name = "A2 Landscape", Width = 0.594, Height = 0.420, SwPaperSize = (int)swDwgPaperSizes_e.swDwgPaperA2size },
            new StandardSheet { Name = "A1 Landscape", Width = 0.841, Height = 0.594, SwPaperSize = (int)swDwgPaperSizes_e.swDwgPaperA1size },
            new StandardSheet { Name = "A0 Landscape", Width = 1.189, Height = 0.841, SwPaperSize = (int)swDwgPaperSizes_e.swDwgPaperA0size }
        };

        public TemplateManager(ISldWorks swApp)
        {
            _swApp = swApp;
        }

        /// <summary>
        /// Given the bounding box of a model, calculates the required footprint to safely house
        /// a standard 3-View + Isometric layout, and selects the optimal sheet + valid ISO scale.
        /// </summary>
        public (StandardSheet BestSheet, double ScaleRatio) ComputeOptimalLayout(double[] boundingBox, bool needTop = true, bool needRight = true, bool needIso = true)
        {
            if (boundingBox == null || boundingBox.Length < 6)
                return (ISO_SHEETS[1], 1.0/5.0); // Fallback: A3 @ 1:5

            // Extract part dimensions in meters
            double dx = Math.Abs(boundingBox[3] - boundingBox[0]);
            double dy = Math.Abs(boundingBox[4] - boundingBox[1]);
            double dz = Math.Abs(boundingBox[5] - boundingBox[2]);

            // Model Space Projection Footprint
            double frontWidth = dx;
            double frontHeight = dy;
            double topHeight = dz;
            double rightWidth = dz;
            
            // Start with Front view base requirement
            double modelSpaceWidth = frontWidth;
            double modelSpaceHeight = frontHeight;

            // Add Right view footprint if needed
            if (needRight)
                modelSpaceWidth += rightWidth + 0.050; // 50mm padding

            // Add Top view footprint if needed
            if (needTop)
                modelSpaceHeight += topHeight + 0.050;

            // Add ISO view footprint estimation if needed
            if (needIso)
            {
                double isoDiag = Math.Sqrt((dx*dx) + (dy*dy) + (dz*dz));
                // Add ISO footprint to either width or height depending on aspect ratio to guarantee fit
                if (modelSpaceWidth > modelSpaceHeight)
                    modelSpaceHeight += isoDiag;
                else
                    modelSpaceWidth += isoDiag;
            }

            // Always add a baseline 30mm margin padding for general safe space
            modelSpaceWidth += 0.030;
            modelSpaceHeight += 0.030;

            // Iterate through sheets from smallest to largest
            foreach (var sheet in ISO_SHEETS)
            {
                // What scale is strictly required to fit this model footprint onto this sheet?
                double requiredScaleW = sheet.UsableWidth / modelSpaceWidth;
                double requiredScaleH = sheet.UsableHeight / modelSpaceHeight;
                
                double rawScale = Math.Min(requiredScaleW, requiredScaleH);
                
                // Snap to nearest smaller (or equal) valid ISO scale
                double validScale = GetValidISO5455Scale(rawScale);

                // If the valid scale isn't ridiculously small, use this sheet!
                // We prefer not to drop below 1:50 on an A4. If it demands that, we upsize the sheet.
                if (validScale >= 1.0/50.0 || sheet.Name == "A0 Landscape")
                {
                    return (sheet, validScale);
                }
            }

            // Absolute fallback
            return (ISO_SHEETS.Last(), GetValidISO5455Scale(0.001));
        }

        /// <summary>
        /// Rounds down a raw mathematical scale to the closest standard ISO 5455 scale strictly.
        /// Never allow arbitrary scaling like 1:3.
        /// </summary>
        public static double GetValidISO5455Scale(double rawScale)
        {
            double bestScale = ISO_5455_SCALES.Last(); // start at smallest
            
            // Because array is sorted from largest logic to smallest:
            for (int i = 0; i < ISO_5455_SCALES.Length; i++)
            {
                if (ISO_5455_SCALES[i] <= rawScale)
                {
                    return ISO_5455_SCALES[i];
                }
            }
            return bestScale;
        }

        /// <summary>
        /// Forces the drawing template geometry (Paper Space) to lock and never scale when views scale.
        /// Applies the decided sheet size.
        /// </summary>
        public void ApplyTemplateToDrawing(DrawingDoc drawing, StandardSheet sheet, double isoScale)
        {
            if (drawing == null) return;
            
            // Set scale in the SW API string format (e.g., 1:5)
            // SolidWorks handles scale by setting numerators and denominators.
            double[] fraction = ConvertDecimalToFraction(isoScale);
            
            bool success = drawing.SetupSheet5(
                ((Sheet)drawing.GetCurrentSheet()).GetName(),
                sheet.SwPaperSize,
                (int)swDwgTemplates_e.swDwgTemplateNone, // We use custom background formatting, ignore SW default
                fraction[0], fraction[1], // Scale numerator : denominator
                true,
                "", // Template path if exists
                sheet.Width,
                sheet.Height,
                "Default",
                true);
        }

        /// <summary>
        /// Converts decimal double (e.g. 0.20) to scaling integers for SW API (e.g. [1, 5])
        /// </summary>
        public static double[] ConvertDecimalToFraction(double value)
        {
            if (value >= 1.0)
            {
                return new double[] { Math.Round(value), 1.0 };
            }
            else
            {
                return new double[] { 1.0, Math.Round(1.0 / value) };
            }
        }
    }
}
