using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine.Annotations
{
    /// <summary>
    /// V2 Intelligent Annotation Solver.
    /// Intercepts SolidWorks AutoDimension outcomes and mathematically resolves overlaps
    /// against geometry and other dimensions using AABB intersection tests.
    /// </summary>
    public class DimensionCollisionSolver
    {
        private readonly double _nudgeStep; // e.g. 5mm step in Paper Space
        private readonly double _padding;   // Security padding around bounding boxes

        public class Rect2D
        {
            public double MinX { get; set; }
            public double MinY { get; set; }
            public double MaxX { get; set; }
            public double MaxY { get; set; }

            public double Width => MaxX - MinX;
            public double Height => MaxY - MinY;

            public Rect2D(double minX, double minY, double maxX, double maxY)
            {
                MinX = minX;
                MinY = minY;
                MaxX = maxX;
                MaxY = maxY;
            }

            public bool Intersects(Rect2D other, double margin = 0)
            {
                if (this.MaxX + margin < other.MinX || this.MinX - margin > other.MaxX) return false;
                if (this.MaxY + margin < other.MinY || this.MinY - margin > other.MaxY) return false;
                return true;
            }
        }

        public DimensionCollisionSolver(double nudgeStepMeters = 0.005, double paddingMeters = 0.002)
        {
            _nudgeStep = nudgeStepMeters;
            _padding = paddingMeters;
        }

        /// <summary>
        /// Takes a populated View, evaluates all placed DisplayDimensions, 
        /// and shoves any intersecting text into clear Paper Space.
        /// </summary>
        public void ResolveViewCollisions(IView swView)
        {
            if (swView == null) return;

            // 1. Get View Geometry Outline [MinX, MinY, MaxX, MaxY] in Paper Space
            double[] outline = (double[])swView.GetOutline();
            if (outline == null || outline.Length < 4) return;
            
            Rect2D viewGeometryBox = new Rect2D(outline[0], outline[1], outline[2], outline[3]);
            List<Rect2D> placedAnnotationBoxes = new List<Rect2D>();

            // 2. Extract dimensions
            IDisplayDimension swDispDim = swView.GetFirstDisplayDimension5();
            while (swDispDim != null)
            {
                IAnnotation swAnn = swDispDim.GetAnnotation();
                if (swAnn != null)
                {
                    // [0]=Left, [1]=Top, [2]=Right, [3]=Bottom, [4]=Z
                    double[] extent = (double[])swAnn.GetExtent();
                    if (extent != null && extent.Length >= 4)
                    {
                        // Note: SW Y-axis points UP in paper space, but Top/Bottom logic varies.
                        // We sort explicitly to get absolute Min/Max
                        double tMinX = Math.Min(extent[0], extent[2]);
                        double tMaxX = Math.Max(extent[0], extent[2]);
                        double tMinY = Math.Min(extent[1], extent[3]);
                        double tMaxY = Math.Max(extent[1], extent[3]);

                        Rect2D txtBox = new Rect2D(tMinX, tMinY, tMaxX, tMaxY);
                        
                        // 3. Check Overlap Condition
                        bool interference = CheckCollision(txtBox, viewGeometryBox, placedAnnotationBoxes);
                        
                        if (interference)
                        {
                            // 4. Resolve Overlap
                            txtBox = NudgeToClearance(swDispDim, swAnn, txtBox, viewGeometryBox, placedAnnotationBoxes);
                        }

                        // Register final position
                        placedAnnotationBoxes.Add(txtBox);
                    }
                }
                swDispDim = swDispDim.GetNext5();
            }
        }

        private bool CheckCollision(Rect2D target, Rect2D geometry, List<Rect2D> obstacles)
        {
            // Dim overlapping main view outline (we don't want Dims inside the part)
            if (target.Intersects(geometry, _padding)) return true;

            // Dim overlapping other dims
            foreach (var obs in obstacles)
            {
                if (target.Intersects(obs, _padding)) return true;
            }
            return false;
        }

        private Rect2D NudgeToClearance(
            IDisplayDimension dim, 
            IAnnotation ann, 
            Rect2D txtBox, 
            Rect2D geometry, 
            List<Rect2D> obstacles)
        {
            // Current text position center in Paper Space
            double[] pos = (double[])ann.GetPosition();
            if (pos == null) return txtBox; // Fallback if API fails

            double cx = pos[0];
            double cy = pos[1];
            double cz = pos[2];

            // Heuristic Nudge Vectors: Try moving outward radially based on quadrant
            // Quadrant relative to geometry center
            double geoCx = geometry.MinX + (geometry.Width / 2.0);
            double geoCy = geometry.MinY + (geometry.Height / 2.0);

            // Vector to push away from center
            double dirX = (cx >= geoCx) ? 1.0 : -1.0;
            double dirY = (cy >= geoCy) ? 1.0 : -1.0;

            // Iterative resolution loop (Limit to 20 attempts to prevent infinite freeze)
            int attempts = 0;
            const int MAX_ATTEMPTS = 20;

            Rect2D currentBox = new Rect2D(txtBox.MinX, txtBox.MinY, txtBox.MaxX, txtBox.MaxY);

            while (CheckCollision(currentBox, geometry, obstacles) && attempts < MAX_ATTEMPTS)
            {
                attempts++;

                // Nudge position
                cx += dirX * _nudgeStep;
                cy += dirY * _nudgeStep;

                // Shift box linearly for testing
                currentBox.MinX += dirX * _nudgeStep;
                currentBox.MaxX += dirX * _nudgeStep;
                currentBox.MinY += dirY * _nudgeStep;
                currentBox.MaxY += dirY * _nudgeStep;
            }

            // Commit position
            ann.SetPosition(cx, cy, cz);
            return currentBox;
        }
    }
}
