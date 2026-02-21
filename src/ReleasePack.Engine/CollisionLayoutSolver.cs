using System;
using System.Collections.Generic;
using System.Linq;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Force-directed annotation layout solver.
    /// Prevents overlapping dimensions and annotations using:
    ///   1. Repulsive forces between overlapping nodes
    ///   2. Repulsive forces from view geometry edges
    ///   3. Attractive spring force toward ideal position
    ///   4. Bounded iteration (max 50 cycles)
    ///   5. Convergence detection (< 0.1mm displacement)
    /// 
    /// Deterministic: same input → same output.
    /// </summary>
    public class CollisionLayoutSolver
    {
        /// <summary>
        /// A positioned annotation node with bounding rectangle.
        /// </summary>
        public class AnnotationNode
        {
            public string Id { get; set; }
            public double X { get; set; }           // Current X position (meters)
            public double Y { get; set; }           // Current Y position
            public double Width { get; set; }       // Bounding width
            public double Height { get; set; }      // Bounding height
            public double IdealX { get; set; }      // Preferred X position
            public double IdealY { get; set; }      // Preferred Y position
            public double AnchorX { get; set; }     // Feature attachment point X
            public double AnchorY { get; set; }     // Feature attachment point Y

            public double Left => X - Width / 2;
            public double Right => X + Width / 2;
            public double Bottom => Y - Height / 2;
            public double Top => Y + Height / 2;
        }

        /// <summary>
        /// View boundary rectangle.
        /// </summary>
        public class ViewBounds
        {
            public double Left { get; set; }
            public double Right { get; set; }
            public double Bottom { get; set; }
            public double Top { get; set; }
        }

        // Layout parameters
        private const int MAX_ITERATIONS = 50;
        private const double CONVERGE_THRESHOLD = 0.0001;  // 0.1mm
        private const double DAMPING = 0.5;
        private const double SPRING_CONSTANT = 0.3;
        private const double MIN_GAP_DEFAULT = 0.008;      // 8mm ISO

        /// <summary>
        /// Resolve all annotation overlaps using force-directed relaxation.
        /// </summary>
        /// <param name="nodes">Annotation nodes to position.</param>
        /// <param name="viewBounds">View bounding rectangles to avoid.</param>
        /// <param name="minGap">Minimum gap between annotations (meters).</param>
        public void Solve(
            List<AnnotationNode> nodes,
            List<ViewBounds> viewBounds,
            double minGap = MIN_GAP_DEFAULT)
        {
            if (nodes == null || nodes.Count < 2) return;

            // Initialize at ideal positions
            foreach (var node in nodes)
            {
                node.X = node.IdealX;
                node.Y = node.IdealY;
            }

            // Push outside view bounds initially
            foreach (var node in nodes)
            {
                foreach (var vb in viewBounds)
                {
                    PushOutsideView(node, vb, minGap);
                }
            }

            // Iterative relaxation
            for (int iter = 0; iter < MAX_ITERATIONS; iter++)
            {
                double maxDisplacement = 0;

                for (int i = 0; i < nodes.Count; i++)
                {
                    double forceX = 0, forceY = 0;

                    // ── Repulsive force from other annotations ──
                    for (int j = 0; j < nodes.Count; j++)
                    {
                        if (i == j) continue;

                        double overlap = ComputeOverlap(nodes[i], nodes[j], minGap);
                        if (overlap > 0)
                        {
                            double dx = nodes[i].X - nodes[j].X;
                            double dy = nodes[i].Y - nodes[j].Y;
                            double dist = Math.Sqrt(dx * dx + dy * dy);

                            if (dist < 1e-8)
                            {
                                // Exactly overlapping — push in deterministic direction
                                dx = 0.001 * (i - j);
                                dy = 0.001;
                                dist = Math.Sqrt(dx * dx + dy * dy);
                            }

                            double magnitude = overlap + minGap;
                            forceX += (dx / dist) * magnitude;
                            forceY += (dy / dist) * magnitude;
                        }
                    }

                    // ── Repulsive force from view boundaries ──
                    foreach (var vb in viewBounds)
                    {
                        // Check if node overlaps with view
                        if (nodes[i].Right > vb.Left && nodes[i].Left < vb.Right &&
                            nodes[i].Top > vb.Bottom && nodes[i].Bottom < vb.Top)
                        {
                            // Push in direction of minimum escape
                            double escapeLeft = vb.Left - nodes[i].Right;
                            double escapeRight = nodes[i].Left - vb.Right;
                            double escapeDown = vb.Bottom - nodes[i].Top;
                            double escapeUp = nodes[i].Bottom - vb.Top;

                            double minEscape = double.MaxValue;
                            double pushX = 0, pushY = 0;

                            if (Math.Abs(escapeLeft) < Math.Abs(minEscape)) { minEscape = escapeLeft; pushX = escapeLeft - minGap; pushY = 0; }
                            if (Math.Abs(escapeRight) < Math.Abs(minEscape)) { minEscape = escapeRight; pushX = -escapeRight + minGap; pushY = 0; }
                            if (Math.Abs(escapeDown) < Math.Abs(minEscape)) { minEscape = escapeDown; pushX = 0; pushY = escapeDown - minGap; }
                            if (Math.Abs(escapeUp) < Math.Abs(minEscape)) { minEscape = escapeUp; pushX = 0; pushY = -escapeUp + minGap; }

                            forceX += pushX;
                            forceY += pushY;
                        }
                    }

                    // ── Attractive spring force toward ideal position ──
                    forceX += (nodes[i].IdealX - nodes[i].X) * SPRING_CONSTANT;
                    forceY += (nodes[i].IdealY - nodes[i].Y) * SPRING_CONSTANT;

                    // Apply force with damping
                    double dispX = forceX * DAMPING;
                    double dispY = forceY * DAMPING;

                    nodes[i].X += dispX;
                    nodes[i].Y += dispY;

                    double disp = Math.Sqrt(dispX * dispX + dispY * dispY);
                    if (disp > maxDisplacement) maxDisplacement = disp;
                }

                // Convergence check
                if (maxDisplacement < CONVERGE_THRESHOLD)
                    break;
            }

            // Final stacking for any remaining overlaps
            FinalStack(nodes, minGap);
        }

        /// <summary>
        /// Compute the minimum push distance to separate two rectangles.
        /// Returns 0 if no overlap.
        /// </summary>
        private double ComputeOverlap(AnnotationNode a, AnnotationNode b, double gap)
        {
            double overlapX = Math.Max(0,
                Math.Min(a.Right + gap, b.Right + gap) -
                Math.Max(a.Left - gap, b.Left - gap) -
                (a.Width + b.Width + 2 * gap));

            // Simplified: check if rectangles (with gap) overlap
            bool overlapH = a.Right + gap > b.Left && a.Left - gap < b.Right;
            bool overlapV = a.Top + gap > b.Bottom && a.Bottom - gap < b.Top;

            if (overlapH && overlapV)
            {
                double sepX = Math.Min(
                    Math.Abs(a.Right - b.Left + gap),
                    Math.Abs(b.Right - a.Left + gap));
                double sepY = Math.Min(
                    Math.Abs(a.Top - b.Bottom + gap),
                    Math.Abs(b.Top - a.Bottom + gap));
                return Math.Min(sepX, sepY);
            }

            return 0;
        }

        private void PushOutsideView(AnnotationNode node, ViewBounds vb, double gap)
        {
            // If node is inside view bounds, push it outside
            if (node.Right > vb.Left && node.Left < vb.Right &&
                node.Top > vb.Bottom && node.Bottom < vb.Top)
            {
                // Push to whichever edge is closest
                double distLeft = Math.Abs(node.Right - vb.Left);
                double distRight = Math.Abs(node.Left - vb.Right);
                double distBottom = Math.Abs(node.Top - vb.Bottom);
                double distTop = Math.Abs(node.Bottom - vb.Top);

                double min = Math.Min(Math.Min(distLeft, distRight), Math.Min(distBottom, distTop));

                if (min == distLeft) node.X = vb.Left - node.Width / 2 - gap;
                else if (min == distRight) node.X = vb.Right + node.Width / 2 + gap;
                else if (min == distBottom) node.Y = vb.Bottom - node.Height / 2 - gap;
                else node.Y = vb.Top + node.Height / 2 + gap;
            }
        }

        private void FinalStack(List<AnnotationNode> nodes, double gap)
        {
            // Last resort: vertically stack any remaining overlaps
            for (int i = 0; i < nodes.Count; i++)
            {
                for (int j = i + 1; j < nodes.Count; j++)
                {
                    if (ComputeOverlap(nodes[i], nodes[j], gap) > 0)
                    {
                        nodes[j].Y = nodes[i].Top + gap + nodes[j].Height / 2;
                    }
                }
            }
        }
    }
}
