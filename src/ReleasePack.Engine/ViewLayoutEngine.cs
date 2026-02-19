using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Calculates optimal view scales and positions for a professional drawing layout.
    /// </summary>
    public class ViewLayoutEngine
    {
        private const double MARGIN = 0.030; // 30mm margin
        private const double VIEW_GAP = 0.035; // 35mm gap

        public struct LayoutResult
        {
            public double Scale;
            public double FrontX, FrontY;
            public double TopX, TopY;
            public double RightX, RightY;
            public double IsoX, IsoY;
            public bool Fits;
        }

        public LayoutResult CalculateLayout(double sheetW, double sheetH, double modelW, double modelH, double modelD)
        {
            // Available area
            double usableW = sheetW - 2 * MARGIN;
            double usableH = sheetH - 2 * MARGIN;

            // 1. Calculate ideal scale to fit Front + Right + Gaps horizontally, and Front + Top + Gaps vertically
            // Layout assumption:
            //   Top View
            // Front View  Right View
            
            double neededW = modelW + modelD + VIEW_GAP;
            double neededH = modelH + modelD + VIEW_GAP;

             // Start with 1:1 and scale down
            double rawScaleW = usableW / neededW;
            double rawScaleH = usableH / neededH;
            double rawScale = Math.Min(rawScaleW, rawScaleH);

            // Snap to standard scales
            double scale = GetStandardScale(rawScale);

            // 2. Calculate positions (centered on sheet)
            double contentW = (modelW + modelD + VIEW_GAP) * scale;
            double contentH = (modelH + modelD + VIEW_GAP) * scale;

            // Center point of the content block
            double startX = MARGIN + (usableW - contentW) / 2;
            double startY = MARGIN + (usableH - contentH) / 2;

            // Front View (Bottom-Left of the block) reference point is usually center of view, 
            // so we need to offset by half dimensions.
            // Let's calculate centers for SW API
            
            double frontW_Scaled = modelW * scale;
            double frontH_Scaled = modelH * scale;
            double topH_Scaled = modelD * scale;
            double rightW_Scaled = modelD * scale;

            // Front View Center
            double frontCx = startX + frontW_Scaled / 2;
            double frontCy = startY + frontH_Scaled / 2;

            // Right View Center
            double rightCx = startX + frontW_Scaled + VIEW_GAP * scale + rightW_Scaled / 2;
            double rightCy = frontCy; // Aligned with Front

            // Top View Center
            double topCx = frontCx; // Aligned with Front
            double topCy = startY + frontH_Scaled + VIEW_GAP * scale + topH_Scaled / 2;

            // ISO View (Top-Right corner of sheet, disconnected from projection alignment)
            double isoCx = sheetW - MARGIN - (modelW * scale * 0.5);
            double isoCy = sheetH - MARGIN - (modelH * scale * 0.5);

            return new LayoutResult
            {
                Scale = scale,
                FrontX = frontCx, FrontY = frontCy,
                RightX = rightCx, RightY = rightCy,
                TopX = topCx, TopY = topCy,
                IsoX = isoCx, IsoY = isoCy,
                Fits = scale >= 0.05 // Arbitrary fail threshold
            };
        }

        private double GetStandardScale(double raw)
        {
            double[] standards = { 0.1, 0.2, 0.25, 0.5, 1.0, 2.0, 5.0, 10.0 };
            // Find largest standard scale <= raw
            var best = standards.Where(s => s <= raw).DefaultIfEmpty(0.1).Max();
            // If we can go slightly larger (e.g. 1:1.5 isn't standard but 1:2 is too small), maybe stick to standards.
            // Actually, let's allow precise scaling if it helps fill the sheet, but standard is better.
            return best;
        }
    }
}
