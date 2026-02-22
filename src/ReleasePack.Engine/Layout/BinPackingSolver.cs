using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine.Layout
{
    /// <summary>
    /// V2 View Placement Solver.
    /// Prevents overlapping geometries by calculating strict 2D bounding boxes and margins before placement.
    /// Uses absolute coordinates mapped to Paper Space constraints.
    /// </summary>
    public class BinPackingSolver
    {
        private readonly double _marginMeters;

        public class ViewEnvelope
        {
            public string ViewType { get; set; } // e.g. "FRONT", "TOP", "RIGHT", "ISO"
            public double Width { get; set; }
            public double Height { get; set; }
            public double TargetCenterX { get; set; }
            public double TargetCenterY { get; set; }
            
            // Evaluated Bounding Box
            public double MinX => TargetCenterX - (Width / 2.0);
            public double MaxX => TargetCenterX + (Width / 2.0);
            public double MinY => TargetCenterY - (Height / 2.0);
            public double MaxY => TargetCenterY + (Height / 2.0);

            public bool Overlaps(ViewEnvelope other, double margin)
            {
                // Simple 2D AABB intersection test with margin
                if (this.MaxX + margin < other.MinX || this.MinX - margin > other.MaxX) return false;
                if (this.MaxY + margin < other.MinY || this.MinY - margin > other.MaxY) return false;
                
                return true;
            }
        }

        /// <summary>
        /// Initializes the solver. Uses 10mm margin for ISO and 12mm for ANSI standard by default.
        /// </summary>
        public BinPackingSolver(bool isAnsi = false)
        {
            _marginMeters = isAnsi ? 0.012 : 0.010;
        }

        /// <summary>
        /// Mathematically calculates strict non-overlapping View Envelopes for standard 3-View + ISO layout.
        /// Returns null if the layout fundamentally breaches the sheet boundaries due to scale.
        /// </summary>
        public Dictionary<string, ViewEnvelope> SolveStandardLayout(
            double[] modelBoundingBox, 
            double scale, 
            TemplateManager.StandardSheet sheet,
            bool includeTop,
            bool includeRight,
            bool includeIso)
        {
            if (modelBoundingBox == null || modelBoundingBox.Length < 6) return null;

            // Model dx, dy, dz multiplied by View Scale to get paper size.
            double dx = Math.Abs(modelBoundingBox[3] - modelBoundingBox[0]) * scale;
            double dy = Math.Abs(modelBoundingBox[4] - modelBoundingBox[1]) * scale;
            double dz = Math.Abs(modelBoundingBox[5] - modelBoundingBox[2]) * scale;

            var layout = new Dictionary<string, ViewEnvelope>();

            // Assuming standard projection where Z is depth (Right/Top view geometry).
            var front = new ViewEnvelope { ViewType = "FRONT", Width = dx, Height = dy };
            
            // 1. Anchor FRONT (Start bottom left equivalent, but centered in its quadrant)
            front.TargetCenterX = _marginMeters * 2 + (front.Width / 2.0);
            front.TargetCenterY = _marginMeters * 2 + (front.Height / 2.0);
            layout.Add("FRONT", front);

            // 2. Derive TOP and RIGHT relatively if requested
            if (includeTop)
            {
                var top = new ViewEnvelope { ViewType = "TOP", Width = dx, Height = dz };
                top.TargetCenterX = front.TargetCenterX;
                top.TargetCenterY = front.MaxY + _marginMeters + (top.Height / 2.0);
                layout.Add("TOP", top);
            }

            if (includeRight)
            {
                var right = new ViewEnvelope { ViewType = "RIGHT", Width = dz, Height = dy };
                right.TargetCenterX = front.MaxX + _marginMeters + (right.Width / 2.0);
                right.TargetCenterY = front.TargetCenterY;
                layout.Add("RIGHT", right);
            }

            // 3. Place ISO in top right zone
            if (includeIso)
            {
                double isoDiag = Math.Sqrt((dx*dx) + (dy*dy) + (dz*dz));
                var iso = new ViewEnvelope { ViewType = "ISO", Width = isoDiag, Height = isoDiag };
                
                double maxX = front.MaxX;
                double maxY = front.MaxY;
                if (includeRight) maxX = layout["RIGHT"].TargetCenterX;
                if (includeTop) maxY = layout["TOP"].TargetCenterY;

                iso.TargetCenterX = Math.Max(maxX, (includeTop ? layout["TOP"].MaxX : front.MaxX)) + _marginMeters + (isoDiag / 2.0);
                iso.TargetCenterY = Math.Max(maxY, (includeRight ? layout["RIGHT"].MaxY : front.MaxY)) + _marginMeters + (isoDiag / 2.0);
                layout.Add("ISO", iso);
            }

            // 4. Validate Boundaries (Do any views spill off the physical sheet?)
            // Sheet coordinates are typically [0,0] bottom-left to [Width, Height] top-right.
            // TitleBlock usually occupies right-side (approx 100mm+), so restrict UsableWidth.
            double availableWidth = sheet.UsableWidth - 0.100; // Deduct 100mm for title block

            foreach (var view in layout.Values)
            {
                if (view.MaxX > availableWidth || view.MinX < 0) return null; // Fails to fit, scale must drop
                if (view.MaxY > sheet.UsableHeight || view.MinY < 0) return null; 
            }

            // 5. Overlap Test
            var viewList = layout.Values.ToList();
            for (int i = 0; i < viewList.Count; i++)
            {
                for (int j = i + 1; j < viewList.Count; j++)
                {
                    if (viewList[i].Overlaps(viewList[j], _marginMeters))
                    {
                        return null; // Force downward scale loop
                    }
                }
            }

            return layout;
        }

        /// <summary>
        /// Explicitly targets an existing view and forces it into position, assigning standard view name tags.
        /// </summary>
        public void CommitView(IView swView, ViewEnvelope envelope)
        {
            if (swView == null || envelope == null) return;

            // Rename view to standard
            swView.SetName2(envelope.ViewType);

            // Physically move the center position array (X, Y, Z=0)
            double[] targetLocation = new double[] { envelope.TargetCenterX, envelope.TargetCenterY, 0 };
            swView.Position = targetLocation;
            
            // Make sure the view does not show hidden edges if it's the ISO
            if (envelope.ViewType == "ISO")
            {
                swView.SetDisplayMode3(false, (int)swDisplayMode_e.swSHADED, false, true);
            }
        }
    }
}
