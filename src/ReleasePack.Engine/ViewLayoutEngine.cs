using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Calculates optimal view scales and positions for a professional drawing layout.
    /// Supports both 1st Angle (ISO) and 3rd Angle (ANSI) projection.
    /// </summary>
    public class ViewLayoutEngine
    {
        private const double MARGIN = 0.030; // 30mm margin
        private const double VIEW_GAP = 0.035; // 35mm gap between views
        private const double TITLE_BLOCK_HEIGHT = 0.040; // 40mm reserved for title block

        public struct LayoutResult
        {
            public double Scale;
            public double FrontX, FrontY;
            public double TopX, TopY;
            public double RightX, RightY;
            public double IsoX, IsoY;
            public bool Fits;
            public bool HasTopView;
            public bool HasRightView;
        }

        /// <summary>
        /// Calculate layout positions for standard projection views.
        /// </summary>
        /// <param name="sheetW">Sheet width in meters</param>
        /// <param name="sheetH">Sheet height in meters</param>
        /// <param name="modelW">Model width (X extent) in meters</param>
        /// <param name="modelH">Model height (Y extent) in meters</param>
        /// <param name="modelD">Model depth (Z extent) in meters</param>
        /// <param name="standard">Projection standard (1st or 3rd angle)</param>
        /// <param name="includeTop">Whether to include top view</param>
        /// <param name="includeRight">Whether to include right view</param>
        public LayoutResult CalculateLayout(
            double sheetW, double sheetH,
            double modelW, double modelH, double modelD,
            ViewStandard standard = ViewStandard.ThirdAngle,
            bool includeTop = true, bool includeRight = true)
        {
            // Available area (account for title block at bottom-right)
            double usableW = sheetW - 2 * MARGIN;
            double usableH = sheetH - 2 * MARGIN - TITLE_BLOCK_HEIGHT;

            // Calculate how much space is needed
            // Front view: W x H
            // Top view: W x D  (placed above or below depending on standard)
            // Right view: D x H  (placed left or right depending on standard)
            double neededW = modelW + (includeRight ? modelD + VIEW_GAP : 0);
            double neededH = modelH + (includeTop ? modelD + VIEW_GAP : 0);

            // Scale calculation
            double rawScaleW = usableW / neededW;
            double rawScaleH = usableH / neededH;
            double rawScale = Math.Min(rawScaleW, rawScaleH);
            double scale = GetStandardScale(rawScale);

            // Scaled dimensions
            double frontW = modelW * scale;
            double frontH = modelH * scale;
            double topH = modelD * scale;   // Top view height = model depth
            double rightW = modelD * scale; // Right view width = model depth
            double gap = VIEW_GAP * scale;

            // Content block dimensions
            double contentW = frontW + (includeRight ? rightW + gap : 0);
            double contentH = frontH + (includeTop ? topH + gap : 0);

            // Starting position (bottom-left of content block, centered on sheet)
            double startX = MARGIN + (usableW - contentW) / 2;
            double startY = MARGIN + TITLE_BLOCK_HEIGHT + (usableH - contentH) / 2;

            double frontCx, frontCy;
            double topCx, topCy;
            double rightCx, rightCy;
            double isoCx, isoCy;

            if (standard == ViewStandard.ThirdAngle)
            {
                // 3rd Angle (ANSI): Top ABOVE Front, Right to the RIGHT of Front
                // SolidWorks sheet coords: (0,0) at bottom-left, Y increases upward

                // Front View Center (lower-left of content block)
                frontCx = startX + frontW / 2;
                frontCy = startY + frontH / 2;

                // Right View: to the right of Front, same Y
                rightCx = startX + frontW + gap + rightW / 2;
                rightCy = frontCy;

                // Top View: above Front, same X
                topCx = frontCx;
                topCy = startY + frontH + gap + topH / 2;
            }
            else
            {
                // 1st Angle (ISO): Top BELOW Front, Right to the LEFT of Front
                // Per ISO standard, the views are "projected" through the object

                // Front View Center (upper-right area of content block)
                frontCx = startX + (includeRight ? rightW + gap : 0) + frontW / 2;
                frontCy = startY + (includeTop ? topH + gap : 0) + frontH / 2;

                // Right View: to the LEFT of Front, same Y
                rightCx = startX + rightW / 2;
                rightCy = frontCy;

                // Top View: BELOW Front, same X as Front
                topCx = frontCx;
                topCy = startY + topH / 2;
            }

            // ISO View: Top-Right corner, disconnected from projection alignment
            isoCx = sheetW - MARGIN - (modelW * scale * 0.5);
            isoCy = sheetH - MARGIN - (modelH * scale * 0.5);

            return new LayoutResult
            {
                Scale = scale,
                FrontX = frontCx, FrontY = frontCy,
                RightX = rightCx, RightY = rightCy,
                TopX = topCx, TopY = topCy,
                IsoX = isoCx, IsoY = isoCy,
                HasTopView = includeTop,
                HasRightView = includeRight,
                Fits = scale >= 0.05
            };
        }

        private double GetStandardScale(double raw)
        {
            double[] standards = { 0.05, 0.1, 0.2, 0.25, 0.5, 1.0, 2.0, 5.0, 10.0 };
            var best = standards.Where(s => s <= raw).DefaultIfEmpty(0.05).Max();
            return best;
        }
    }
}
