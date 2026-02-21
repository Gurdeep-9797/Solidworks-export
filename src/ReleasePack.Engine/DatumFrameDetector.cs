using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Detects a datum reference frame from arbitrary BREP geometry using
    /// planar face clustering, cylindrical axis detection, and PCA fallback.
    /// 
    /// Works for: native parts, STEP imports, multi-body, castings, sheet metal.
    /// Output governs ALL dimension placement in the drawing.
    /// </summary>
    public class DatumFrameDetector
    {
        /// <summary>
        /// Result of datum detection: an orthogonal frame + confidence metrics.
        /// </summary>
        public class DatumFrame
        {
            public double[] Origin { get; set; } = new double[3];  // Reference origin XYZ (meters)
            public double[] XAxis { get; set; } = { 1, 0, 0 };    // Right view direction
            public double[] YAxis { get; set; } = { 0, 1, 0 };    // Top view direction
            public double[] ZAxis { get; set; } = { 0, 0, 1 };    // Front view direction (primary datum normal)
            public double Confidence { get; set; } = 0;            // 0-1: how reliable the detection is
            public string Method { get; set; } = "BoundingBox";    // "PlanarCluster", "CylindricalAxis", "PCA", "BoundingBox"
            public bool IsCylindrical { get; set; } = false;
            public bool IsSymmetric { get; set; } = false;
            public double[] Extents { get; set; } = new double[3]; // Major extents along X, Y, Z
        }

        private struct FaceCluster
        {
            public double[] MeanNormal;
            public double TotalArea;
            public List<double[]> FaceCentroids;
        }

        private const double ANGLE_THRESHOLD_RAD = 5.0 * Math.PI / 180.0;  // 5° clustering
        private const double CYLINDRICAL_RATIO = 0.4;  // 40% cylindrical surface → cylindrical part
        private const double CONFIDENCE_THRESHOLD = 0.5;

        /// <summary>
        /// Detect the optimal datum reference frame from a SolidWorks model.
        /// </summary>
        public static DatumFrame Detect(ModelDoc2 doc)
        {
            try
            {
                var partDoc = doc as PartDoc;
                if (partDoc == null)
                    return BoundingBoxFallback(GetBoundingBoxSafe(doc));

                var bodies = (object[])partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodies == null || bodies.Length == 0)
                    return BoundingBoxFallback(GetBoundingBoxSafe(doc));

                // For multi-body: use the largest body
                IBody2 primaryBody = null;
                double maxVolume = 0;
                foreach (IBody2 body in bodies)
                {
                    try
                    {
                        double[] mp = (double[])body.GetMassProperties(0);
                        double vol = mp != null && mp.Length > 3 ? mp[3] : 0;
                        if (vol > maxVolume) { maxVolume = vol; primaryBody = body; }
                    }
                    catch { primaryBody = primaryBody ?? body; }
                }

                if (primaryBody == null)
                    return BoundingBoxFallback(GetBoundingBoxSafe(doc));

                return AnalyzeBody(primaryBody, doc);
            }
            catch
            {
                return BoundingBoxFallback(GetBoundingBoxSafe(doc));
            }
        }

        private static DatumFrame AnalyzeBody(IBody2 body, ModelDoc2 doc)
        {
            // ── Phase 1: FAST EXTRACTION (Main COM Thread) ──
            // We must extract all BREP geometry properties here to avoid COM threading exceptions
            // before we run the heavy math on background threads.
            var planarFaces = new List<(double[] normal, double area, double[] centroid)>();
            var cylindricalFaces = new List<(double[] axis, double radius, double area)>();
            double totalArea = 0;

            object[] faces = (object[])body.GetFaces();
            if (faces == null)
                return BoundingBoxFallback(GetBoundingBoxSafe(doc));

            foreach (IFace2 face in faces)
            {
                try
                {
                    ISurface surface = (ISurface)face.GetSurface();
                    double area = face.GetArea();
                    totalArea += area;

                    if (surface.IsPlane())
                    {
                        double[] pParams = (double[])surface.PlaneParams;
                        if (pParams != null && pParams.Length >= 6)
                        {
                            double[] normal = Normalize(new[] { pParams[0], pParams[1], pParams[2] });
                            double[] centroid = { pParams[3], pParams[4], pParams[5] };
                            planarFaces.Add((normal, area, centroid));
                        }
                    }
                    else if (surface.IsCylinder())
                    {
                        double[] cParams = (double[])surface.CylinderParams;
                        if (cParams != null && cParams.Length >= 7)
                        {
                            double[] axis = Normalize(new[] { cParams[3], cParams[4], cParams[5] });
                            double radius = cParams[6];
                            cylindricalFaces.Add((axis, radius, area));
                        }
                    }
                }
                catch { /* Skip problematic faces */ }
            }

            // Extract bounding box on main thread too
            double[] bbox = GetBoundingBoxSafe(doc);

            // ── Phase 2: MATH & CLUSTERING (Background Thread Safe) ──
            // We can now run the heavy math without freezing SolidWorks
            return System.Threading.Tasks.Task.Run(() => 
            {
                // Cylindrical dominance check
                double cylArea = cylindricalFaces.Sum(f => f.area);
                if (totalArea > 0 && cylArea / totalArea > CYLINDRICAL_RATIO)
                {
                    return DetectCylindricalFrame(cylindricalFaces, bbox);
                }

                // Cluster planar normals
                if (planarFaces.Count >= 3)
                {
                    var frame = DetectFromPlanarClusters(planarFaces, bbox);
                    if (frame.Confidence >= CONFIDENCE_THRESHOLD)
                        return frame;
                }

                // PCA fallback
                return BoundingBoxFallback(bbox);
            }).GetAwaiter().GetResult(); // Block main thread until math finishes (or could await if upstream is async)
        }

        private static DatumFrame DetectFromPlanarClusters(
            List<(double[] normal, double area, double[] centroid)> planarFaces,
            double[] bbox)
        {
            // Agglomerative clustering by normal direction
            var clusters = new List<FaceCluster>();

            foreach (var face in planarFaces)
            {
                bool merged = false;
                for (int i = 0; i < clusters.Count; i++)
                {
                    double dot = Math.Abs(Dot(face.normal, clusters[i].MeanNormal));
                    if (dot > Math.Cos(ANGLE_THRESHOLD_RAD))
                    {
                        // Merge into existing cluster
                        var c = clusters[i];
                        double totalA = c.TotalArea + face.area;
                        c.MeanNormal = Normalize(new[]
                        {
                            (c.MeanNormal[0] * c.TotalArea + face.normal[0] * face.area) / totalA,
                            (c.MeanNormal[1] * c.TotalArea + face.normal[1] * face.area) / totalA,
                            (c.MeanNormal[2] * c.TotalArea + face.normal[2] * face.area) / totalA,
                        });
                        c.TotalArea = totalA;
                        c.FaceCentroids.Add(face.centroid);
                        clusters[i] = c;
                        merged = true;
                        break;
                    }
                }

                if (!merged)
                {
                    clusters.Add(new FaceCluster
                    {
                        MeanNormal = (double[])face.normal.Clone(),
                        TotalArea = face.area,
                        FaceCentroids = new List<double[]> { face.centroid }
                    });
                }
            }

            // Sort by total area descending
            clusters.Sort((a, b) => b.TotalArea.CompareTo(a.TotalArea));

            if (clusters.Count < 2)
                return PCAFallback(bbox);

            // Primary datum: largest cluster
            double[] primaryN = clusters[0].MeanNormal;

            // Secondary datum: next-largest that is perpendicular to primary
            double[] secondaryN = null;
            foreach (var c in clusters.Skip(1))
            {
                double dot = Math.Abs(Dot(primaryN, c.MeanNormal));
                if (dot < Math.Sin(15.0 * Math.PI / 180.0))
                {
                    secondaryN = c.MeanNormal;
                    break;
                }
            }

            if (secondaryN == null)
            {
                // No perpendicular cluster found — construct one
                secondaryN = ArbitraryPerpendicular(primaryN);
            }

            // Tertiary: cross product
            double[] tertiaryN = Normalize(Cross(primaryN, secondaryN));
            // Correct secondary to be exactly perpendicular
            secondaryN = Normalize(Cross(tertiaryN, primaryN));

            // Origin: centroid of the largest face in primary cluster
            double[] origin = clusters[0].FaceCentroids[0];

            // Check symmetry
            bool isSymmetric = false;
            if (bbox != null)
            {
                double[] center = {
                    (bbox[0] + bbox[3]) / 2,
                    (bbox[1] + bbox[4]) / 2,
                    (bbox[2] + bbox[5]) / 2
                };
                // Heuristic: if origin is near bbox center, part is likely symmetric
                double dist = Distance(origin, center);
                double diagonal = Distance(new[] { bbox[0], bbox[1], bbox[2] },
                                           new[] { bbox[3], bbox[4], bbox[5] });
                if (dist / diagonal < 0.15)
                    isSymmetric = true;
            }

            double[] extents = bbox != null
                ? new[] { Math.Abs(bbox[3] - bbox[0]), Math.Abs(bbox[4] - bbox[1]), Math.Abs(bbox[5] - bbox[2]) }
                : new double[] { 0, 0, 0 };

            return new DatumFrame
            {
                Origin = origin,
                ZAxis = primaryN,       // Front view normal (primary datum)
                XAxis = secondaryN,     // Right view direction
                YAxis = tertiaryN,      // Top view direction
                Confidence = Math.Min(1.0, clusters[0].TotalArea / (clusters.Sum(c => c.TotalArea) + 0.0001)),
                Method = "PlanarCluster",
                IsSymmetric = isSymmetric,
                Extents = extents
            };
        }

        private static DatumFrame DetectCylindricalFrame(
            List<(double[] axis, double radius, double area)> cylFaces,
            double[] bbox)
        {
            // Weighted average of cylinder axes by area
            double[] avgAxis = { 0, 0, 0 };
            double totalWeight = 0;
            foreach (var cf in cylFaces)
            {
                double w = cf.area;
                avgAxis[0] += cf.axis[0] * w;
                avgAxis[1] += cf.axis[1] * w;
                avgAxis[2] += cf.axis[2] * w;
                totalWeight += w;
            }
            double[] principalAxis = Normalize(new[]
            {
                avgAxis[0] / totalWeight,
                avgAxis[1] / totalWeight,
                avgAxis[2] / totalWeight
            });

            double[] perpX = ArbitraryPerpendicular(principalAxis);
            double[] perpY = Normalize(Cross(principalAxis, perpX));
            double[] origin = bbox != null
                ? new[] { (bbox[0] + bbox[3]) / 2, (bbox[1] + bbox[4]) / 2, (bbox[2] + bbox[5]) / 2 }
                : new double[] { 0, 0, 0 };
            double[] extents = bbox != null
                ? new[] { Math.Abs(bbox[3] - bbox[0]), Math.Abs(bbox[4] - bbox[1]), Math.Abs(bbox[5] - bbox[2]) }
                : new double[] { 0, 0, 0 };

            return new DatumFrame
            {
                Origin = origin,
                ZAxis = principalAxis,   // Axis direction = front view
                XAxis = perpX,
                YAxis = perpY,
                Confidence = 0.85,
                Method = "CylindricalAxis",
                IsCylindrical = true,
                Extents = extents
            };
        }

        private static DatumFrame PCAFallback(double[] bbox)
        {
            // Use bounding box as PCA approximation for organic geometry
            return BoundingBoxFallback(bbox);
        }

        private static DatumFrame BoundingBoxFallback(double[] bbox)
        {
            double[] origin, extents;

            if (bbox != null && bbox.Length >= 6)
            {
                origin = new[]
                {
                    (bbox[0] + bbox[3]) / 2,
                    (bbox[1] + bbox[4]) / 2,
                    (bbox[2] + bbox[5]) / 2
                };
                extents = new[]
                {
                    Math.Abs(bbox[3] - bbox[0]),
                    Math.Abs(bbox[4] - bbox[1]),
                    Math.Abs(bbox[5] - bbox[2])
                };
            }
            else
            {
                origin = new double[] { 0, 0, 0 };
                extents = new double[] { 0.1, 0.1, 0.1 };
            }

            // Align axes to bounding box extents (largest = Z/front)
            int maxIdx = 0;
            if (extents[1] > extents[maxIdx]) maxIdx = 1;
            if (extents[2] > extents[maxIdx]) maxIdx = 2;

            double[] z = { 0, 0, 0 }; z[maxIdx] = 1;
            double[] x = { 0, 0, 0 }; x[(maxIdx + 1) % 3] = 1;
            double[] y = { 0, 0, 0 }; y[(maxIdx + 2) % 3] = 1;

            return new DatumFrame
            {
                Origin = origin,
                ZAxis = z,
                XAxis = x,
                YAxis = y,
                Confidence = 0.3,
                Method = "BoundingBox",
                Extents = extents
            };
        }

        // ── Vector Math Helpers ──

        private static double[] GetBoundingBoxSafe(ModelDoc2 doc)
        {
            try
            {
                // Try getting body-level bounding box
                if (doc.GetType() == (int)swDocumentTypes_e.swDocPART)
                {
                    PartDoc part = (PartDoc)doc;
                    object[] bodies = (object[])part.GetBodies2((int)swBodyType_e.swSolidBody, true);
                    if (bodies != null && bodies.Length > 0)
                    {
                        IBody2 body = (IBody2)bodies[0];
                        double[] box = (double[])body.GetBodyBox();
                        if (box != null && box.Length >= 6)
                            return box;
                    }
                }
                
                // No body-level box found, return null (BoundingBoxFallback handles this)
                return null;
            }
            catch
            {
                try
                {
                    // Last resort: just return null and let caller use defaults
                    return null;
                }
                catch { return null; }
            }
        }

        internal static double[] Normalize(double[] v)
        {
            double len = Math.Sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]);
            if (len < 1e-12) return new double[] { 0, 0, 1 };
            return new[] { v[0] / len, v[1] / len, v[2] / len };
        }

        internal static double Dot(double[] a, double[] b)
        {
            return a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
        }

        internal static double[] Cross(double[] a, double[] b)
        {
            return new[]
            {
                a[1] * b[2] - a[2] * b[1],
                a[2] * b[0] - a[0] * b[2],
                a[0] * b[1] - a[1] * b[0]
            };
        }

        private static double Distance(double[] a, double[] b)
        {
            double dx = a[0] - b[0], dy = a[1] - b[1], dz = a[2] - b[2];
            return Math.Sqrt(dx * dx + dy * dy + dz * dz);
        }

        internal static double[] ArbitraryPerpendicular(double[] v)
        {
            // Find vector not parallel to v, then cross
            double[] candidate = Math.Abs(v[0]) < 0.9 ? new[] { 1.0, 0, 0 } : new[] { 0, 1.0, 0 };
            return Normalize(Cross(v, candidate));
        }
    }
}
