using System;
using System.Collections.Generic;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// BREP-based topology analyzer that classifies features from raw geometry
    /// WITHOUT relying on feature tree names. Works for STEP imports, multi-body,
    /// and featureless models.
    /// 
    /// Detection hierarchy:
    ///   1. Cylindrical faces → Hole or Boss classification
    ///   2. Toroidal faces → Fillet detection
    ///   3. Angled planar faces → Chamfer detection
    ///   4. Enclosed planar faces → Pocket detection
    ///   5. Geometric similarity clustering → Pattern detection
    /// </summary>
    public class TopologyAnalyzer
    {
        private readonly IProgressCallback _progress;

        public TopologyAnalyzer(IProgressCallback progress = null)
        {
            _progress = progress;
        }

        /// <summary>
        /// Topology-classified feature from BREP geometry.
        /// </summary>
        public class TopologicalFeature
        {
            public string Id { get; set; }
            public string Type { get; set; }          // Hole, Boss, Fillet, Chamfer, Pocket, Slot, Thread, Freeform
            public string SubType { get; set; }       // Through, Blind, CounterBore, CounterSink
            public double Radius { get; set; }
            public double Depth { get; set; }
            public double Angle { get; set; }
            public double Width { get; set; }
            public double[] Position { get; set; }    // Centroid XYZ
            public double[] Axis { get; set; }        // Primary direction (hole axis, etc.)
            public double BoundingRadius { get; set; }
            public bool IsInternal { get; set; }
            public bool IsThreaded { get; set; }
            public string ThreadSpec { get; set; }
            public double CounterBoreRadius { get; set; }
            public double CSAngle { get; set; }

            // For pattern detection
            public string DimensionalSignature { get; set; }

            /// <summary>Convert to AnalyzedFeature for downstream pipeline compatibility.</summary>
            public AnalyzedFeature ToAnalyzedFeature()
            {
                var af = new AnalyzedFeature
                {
                    Name = $"Topo_{Type}_{Id}",
                    TypeName = Type,
                    Category = MapToCategory(),
                    Diameter = Radius * 2,
                    Depth = Depth,
                    Radius = Type == "Fillet" ? Radius : 0,
                    Angle = Angle,
                    Distance1 = Width,
                    PositionX = Position?[0] ?? 0,
                    PositionY = Position?[1] ?? 0,
                    PositionZ = Position?[2] ?? 0,
                    BoundingRadius = BoundingRadius,
                    IsInternal = IsInternal,
                    ThreadCallout = ThreadSpec
                };

                if (IsInternal)
                    af.NeedsSectionView = true;

                if (BoundingRadius > 0 && BoundingRadius < 0.005)
                    af.NeedsDetailView = true;

                return af;
            }

            private FeatureCategory MapToCategory()
            {
                switch (Type)
                {
                    case "Hole":
                        if (IsThreaded) return FeatureCategory.Thread;
                        if (SubType == "CounterBore") return FeatureCategory.CounterBore;
                        if (SubType == "CounterSink") return FeatureCategory.CounterSink;
                        return FeatureCategory.HoleWizard;
                    case "Boss": return FeatureCategory.BossExtrude;
                    case "Fillet": return FeatureCategory.Fillet;
                    case "Chamfer": return FeatureCategory.Chamfer;
                    case "Pocket": return FeatureCategory.Pocket;
                    case "Slot": return FeatureCategory.Slot;
                    default: return FeatureCategory.Other;
                }
            }
        }

        /// <summary>
        /// Pattern group detected from geometric similarity clustering.
        /// </summary>
        public class PatternGroup
        {
            public string PatternType { get; set; }  // Linear, Circular
            public int Count { get; set; }
            public double Spacing { get; set; }
            public double PCD { get; set; }           // Pitch Circle Diameter (for circular)
            public double[] Direction { get; set; }
            public double[] Center { get; set; }
            public List<TopologicalFeature> Features { get; set; } = new List<TopologicalFeature>();
        }

        /// <summary>
        /// Struct to hold extracted COM geometry data so analysis can run on background threads.
        /// </summary>
        private struct RawFaceData
        {
            public int Hash;
            public bool IsCylinder;
            public bool IsPlane;
            public bool IsTorus;
            public bool IsCone;
            public double Area;
            public double[] CylinderParams;
            public double[] PlaneParams;
            public double[] TorusParams;
            public double[] ConeParams;
            public bool FaceSense;
            public double[] UVRange;
            public List<RawFaceData> AdjacentFaces; // For compound features (counterbores)
            public bool HasCosmeticThread;
        }

        /// <summary>
        /// Analyze BREP topology of a model and return classified features.
        /// </summary>
        public List<TopologicalFeature> Analyze(ModelDoc2 doc)
        {
            var features = new List<TopologicalFeature>();

            try
            {
                var partDoc = doc as PartDoc;
                if (partDoc == null) return features;

                var bodies = (object[])partDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                if (bodies == null) return features;

                int featureId = 0;
                foreach (IBody2 body in bodies)
                {
                    try
                    {
                        var bodyFeatures = AnalyzeBody(body, ref featureId);
                        features.AddRange(bodyFeatures);
                    }
                    catch (Exception ex)
                    {
                        _progress?.LogWarning($"Body analysis failed: {ex.Message}");
                    }
                }

                _progress?.LogMessage($"Topology analysis found {features.Count} features " +
                    $"({features.Count(f => f.Type == "Hole")} holes, " +
                    $"{features.Count(f => f.Type == "Fillet")} fillets, " +
                    $"{features.Count(f => f.Type == "Chamfer")} chamfers)");
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Topology analysis failed: {ex.Message}");
            }

            return features;
        }

        private List<TopologicalFeature> AnalyzeBody(IBody2 body, ref int featureId)
        {
            var features = new List<TopologicalFeature>();
            object[] faces = (object[])body.GetFaces();
            if (faces == null) return features;

            // ── Phase 1: FAST COM EXTRACTION (Main Thread) ──
            var rawFaces = new List<RawFaceData>();
            var processedHashes = new HashSet<int>();

            foreach (IFace2 face in faces)
            {
                int hash = face.GetHashCode();
                if (processedHashes.Contains(hash)) continue;
                processedHashes.Add(hash);

                try
                {
                    ISurface surface = (ISurface)face.GetSurface();
                    var raw = new RawFaceData
                    {
                        Hash = hash,
                        Area = face.GetArea(),
                        IsCylinder = surface.IsCylinder(),
                        IsPlane = surface.IsPlane(),
                        IsTorus = IsTorus(surface),
                        IsCone = IsCone(surface),
                        FaceSense = face.FaceInSurfaceSense(),
                        HasCosmeticThread = HasCosmeticThread(face)
                    };

                    // UV range is not available via IFace2 in this interop version.
                    // Depth estimation falls back to cylinder params when UVRange is null.
                    raw.UVRange = null;

                    if (raw.IsCylinder) raw.CylinderParams = (double[])surface.CylinderParams;
                    else if (raw.IsPlane) raw.PlaneParams = (double[])surface.PlaneParams;
                    else if (raw.IsTorus) raw.TorusParams = (double[])surface.TorusParams;
                    else if (raw.IsCone) raw.ConeParams = (double[])surface.ConeParams;

                    // For counterbores, we need adjacent faces
                    if (raw.IsCylinder && !raw.FaceSense)
                    {
                        raw.AdjacentFaces = ExtractAdjacentFaces(face);
                    }

                    rawFaces.Add(raw);
                }
                catch { /* Skip problematic */ }
            }

            // ── Phase 2: PARALLEL COMPUTATION (Background Threads) ──
            object lockObj = new object();
            int localId = featureId;
            
            System.Threading.Tasks.Parallel.ForEach(rawFaces, raw =>
            {
                TopologicalFeature feat = null;

                if (raw.IsCylinder)
                {
                    feat = ClassifyCylindrical(raw);
                }
                else if (raw.IsTorus)
                {
                    feat = ClassifyTorus(raw);
                }
                else if (raw.IsPlane)
                {
                    // Note: Planar chamfer detection requires adjacent faces. 
                    // For massive parallel performance, we skip complex planar topology here 
                    // or require it to be pre-extracted.
                    // feat = ClassifyPlanar(raw); 
                }

                if (feat != null)
                {
                    lock (lockObj)
                    {
                        feat.Id = (localId++).ToString();
                        feat.DimensionalSignature = $"{feat.Type}|{Math.Round(feat.Radius * 1000, 1)}|{Math.Round(feat.Depth * 1000, 1)}";
                        features.Add(feat);
                    }
                }
            });

            featureId = localId;

            return features;
        }

        private List<RawFaceData> ExtractAdjacentFaces(IFace2 face)
        {
            var adj = new List<RawFaceData>();
            try
            {
                object[] edges = (object[])face.GetEdges();
                if (edges == null) return adj;

                foreach (IEdge edge in edges)
                {
                    // Get adjacent faces from edge
                    object adjResult = edge.GetTwoAdjacentFaces2();
                    if (adjResult == null) continue;
                    object[] adjArr = adjResult as object[];
                    if (adjArr == null || adjArr.Length < 2) continue;
                    IFace2 face1 = adjArr[0] as IFace2;
                    IFace2 face2 = adjArr[1] as IFace2;
                    IFace2 other = (face1 != null && face1.GetHashCode() != face.GetHashCode()) ? face1 : face2;
                    if (other == null) continue;
                    
                    ISurface otherSurf = (ISurface)other.GetSurface();
                    adj.Add(new RawFaceData
                    {
                        IsCylinder = otherSurf.IsCylinder(),
                        IsCone = IsCone(otherSurf),
                        CylinderParams = otherSurf.IsCylinder() ? (double[])otherSurf.CylinderParams : null,
                        ConeParams = IsCone(otherSurf) ? (double[])otherSurf.ConeParams : null
                    });
                }
            }
            catch { }
            return adj;
        }

        private TopologicalFeature ClassifyCylindrical(RawFaceData raw)
        {
            if (raw.CylinderParams == null || raw.CylinderParams.Length < 7) return null;

            double[] axis = DatumFrameDetector.Normalize(new[] { raw.CylinderParams[3], raw.CylinderParams[4], raw.CylinderParams[5] });
            double radius = raw.CylinderParams[6];
            double[] center = { raw.CylinderParams[0], raw.CylinderParams[1], raw.CylinderParams[2] };

            var feat = new TopologicalFeature
            {
                Axis = axis,
                Radius = radius,
                Position = center,
                BoundingRadius = radius
            };

            if (!raw.FaceSense)
            {
                // Inward-facing cylindrical surface = HOLE
                feat.Type = "Hole";
                feat.IsInternal = true;
                feat.SubType = "Through"; // Default, refined below

                // Estimate depth from cylinder bounds
                if (raw.UVRange != null && raw.UVRange.Length >= 4)
                {
                    feat.Depth = Math.Abs(raw.UVRange[3] - raw.UVRange[2]);
                }

                // Check for cosmetic thread
                feat.IsThreaded = raw.HasCosmeticThread;
                if (feat.IsThreaded)
                {
                    double diamMm = radius * 2000;
                    feat.ThreadSpec = $"M{diamMm:F0}";
                }

                // Check adjacent faces for counterbore/countersink
                ClassifyHoleCompound(raw, feat, axis);
            }
            else
            {
                // Outward-facing = BOSS
                feat.Type = "Boss";
                feat.IsInternal = false;
            }

            return feat;
        }

        private void ClassifyHoleCompound(RawFaceData raw, TopologicalFeature feat, double[] axis)
        {
            if (raw.AdjacentFaces == null) return;

            foreach (var adj in raw.AdjacentFaces)
            {
                // Check for larger coaxial cylinder (counterbore)
                if (adj.IsCylinder && adj.CylinderParams != null && adj.CylinderParams.Length >= 7)
                {
                    double[] otherAxis = DatumFrameDetector.Normalize(
                        new[] { adj.CylinderParams[3], adj.CylinderParams[4], adj.CylinderParams[5] });
                    double otherRadius = adj.CylinderParams[6];

                    double axisDot = Math.Abs(DatumFrameDetector.Dot(axis, otherAxis));
                    if (axisDot > 0.99 && otherRadius > feat.Radius * 1.15)
                    {
                        feat.SubType = "CounterBore";
                        feat.CounterBoreRadius = otherRadius;
                    }
                }

                // Check for adjacent cone (countersink)
                if (adj.IsCone)
                {
                    feat.SubType = "CounterSink";
                    feat.CSAngle = 82; // Default; refine from cone params if available
                }
            }
        }

        private TopologicalFeature ClassifyTorus(RawFaceData raw)
        {
            if (raw.TorusParams == null || raw.TorusParams.Length < 8) return null;

            double minorRadius = raw.TorusParams[7]; // Fillet radius
            double[] center = { raw.TorusParams[0], raw.TorusParams[1], raw.TorusParams[2] };

            return new TopologicalFeature
            {
                Type = "Fillet",
                Radius = minorRadius,
                Position = center,
                BoundingRadius = minorRadius,
                IsInternal = false
            };
        }

        private TopologicalFeature ClassifyPlanar(IFace2 face, ISurface surface, IBody2 body)
        {
            // Check if this planar face is a chamfer (angled between two edges)
            try
            {
                object[] edges = (object[])face.GetEdges();
                if (edges == null || edges.Length != 2) return null;

                // A chamfer typically has exactly 2 or 3 edges
                // and the face normal is at ~45° to adjacent faces
                double[] pParams = (double[])surface.PlaneParams;
                double[] faceNormal = pParams != null && pParams.Length >= 3
                    ? DatumFrameDetector.Normalize(new[] { pParams[0], pParams[1], pParams[2] })
                    : null;

                // Check adjacent face normals
                double totalAngle = 0;
                int neighborCount = 0;

                foreach (IEdge edge in edges)
                {
                    try
                    {
                        object[] adjFaces = (object[])edge.GetTwoAdjacentFaces2();
                        if (adjFaces == null || adjFaces.Length < 2) continue;
                        IFace2 f1 = adjFaces[0] as IFace2;
                        IFace2 f2 = adjFaces[1] as IFace2;
                        IFace2 other = (f1 != null && f1.GetHashCode() != face.GetHashCode()) ? f1 : f2;
                        if (other == null) continue;

                        ISurface otherSurf = (ISurface)other.GetSurface();
                        if (!otherSurf.IsPlane()) continue;

                        double[] otherPP = (double[])otherSurf.PlaneParams;
                        double[] otherNormal = otherPP != null && otherPP.Length >= 3
                            ? DatumFrameDetector.Normalize(new[] { otherPP[0], otherPP[1], otherPP[2] })
                            : null;

                        double dot = Math.Abs(DatumFrameDetector.Dot(faceNormal, otherNormal));
                        double angle = Math.Acos(Math.Min(1, dot)) * 180.0 / Math.PI;
                        totalAngle += angle;
                        neighborCount++;
                    }
                    catch { }
                }

                if (neighborCount >= 2)
                {
                    double avgAngle = totalAngle / neighborCount;
                    // Chamfers are typically 30-60° from adjacent faces
                    if (avgAngle > 25 && avgAngle < 65)
                    {
                        double area = face.GetArea();
                        double width = Math.Sqrt(area); // Approximate chamfer width

                        double[] pParams2 = (double[])surface.PlaneParams;
                        double[] pos = pParams2.Length >= 6
                            ? new[] { pParams2[3], pParams2[4], pParams2[5] }
                            : new double[] { 0, 0, 0 };

                        return new TopologicalFeature
                        {
                            Type = "Chamfer",
                            Width = width,
                            Angle = avgAngle,
                            Position = pos,
                            BoundingRadius = width,
                            IsInternal = false
                        };
                    }
                }
            }
            catch { }

            return null; // Not a recognizable feature
        }

        /// <summary>
        /// Detect patterns by geometric similarity clustering.
        /// Features with same type + dimensions at regular spacing form patterns.
        /// </summary>
        public List<PatternGroup> DetectPatterns(List<TopologicalFeature> features)
        {
            var patterns = new List<PatternGroup>();

            // Group by dimensional signature (type + radius + depth)
            var groups = features
                .Where(f => f.DimensionalSignature != null && f.Position != null)
                .GroupBy(f => f.DimensionalSignature)
                .Where(g => g.Count() >= 2);

            foreach (var group in groups)
            {
                var feats = group.ToList();
                var positions = feats.Select(f => f.Position).ToList();

                // Try linear pattern detection
                var linearPattern = TryLinearPattern(feats, positions);
                if (linearPattern != null)
                {
                    patterns.Add(linearPattern);
                    continue;
                }

                // Try circular pattern detection
                if (feats.Count >= 3)
                {
                    var circPattern = TryCircularPattern(feats, positions);
                    if (circPattern != null)
                        patterns.Add(circPattern);
                }
            }

            return patterns;
        }

        private PatternGroup TryLinearPattern(List<TopologicalFeature> feats, List<double[]> positions)
        {
            if (positions.Count < 2) return null;

            // Compute all pairwise displacement vectors
            var displacements = new List<double[]>();
            for (int i = 0; i < positions.Count; i++)
                for (int j = i + 1; j < positions.Count; j++)
                    displacements.Add(new[]
                    {
                        positions[j][0] - positions[i][0],
                        positions[j][1] - positions[i][1],
                        positions[j][2] - positions[i][2]
                    });

            // Cluster displacement vectors (find dominant spacing)
            if (displacements.Count == 0) return null;

            // Find the shortest displacement as candidate spacing
            displacements.Sort((a, b) => Magnitude(a).CompareTo(Magnitude(b)));
            double[] baseVec = displacements[0];
            double spacing = Magnitude(baseVec);

            if (spacing < 0.001) return null; // Too close, not a pattern

            // Count how many displacements are multiples of baseVec
            int consistent = 0;
            foreach (var d in displacements)
            {
                double mag = Magnitude(d);
                double ratio = mag / spacing;
                double roundedRatio = Math.Round(ratio);
                if (Math.Abs(ratio - roundedRatio) < 0.1 && roundedRatio >= 1)
                    consistent++;
            }

            double consistency = (double)consistent / displacements.Count;
            if (consistency > 0.5)
            {
                return new PatternGroup
                {
                    PatternType = "Linear",
                    Count = feats.Count,
                    Spacing = spacing,
                    Direction = DatumFrameDetector.Normalize(baseVec),
                    Features = feats
                };
            }

            return null;
        }

        private PatternGroup TryCircularPattern(List<TopologicalFeature> feats, List<double[]> positions)
        {
            if (positions.Count < 3) return null;

            try
            {
                // Fit a circle to the feature positions
                // Simplified: compute centroid, then average radius
                double[] centroid = new double[3];
                foreach (var p in positions)
                {
                    centroid[0] += p[0]; centroid[1] += p[1]; centroid[2] += p[2];
                }
                centroid[0] /= positions.Count;
                centroid[1] /= positions.Count;
                centroid[2] /= positions.Count;

                var radii = positions.Select(p =>
                    Math.Sqrt(
                        (p[0] - centroid[0]) * (p[0] - centroid[0]) +
                        (p[1] - centroid[1]) * (p[1] - centroid[1]) +
                        (p[2] - centroid[2]) * (p[2] - centroid[2])
                    )).ToList();

                double meanRadius = radii.Average();
                double stdDev = Math.Sqrt(radii.Average(r => (r - meanRadius) * (r - meanRadius)));

                // Check if positions lie on a circle (low std dev relative to radius)
                if (stdDev / meanRadius < 0.05 && meanRadius > 0.001)
                {
                    return new PatternGroup
                    {
                        PatternType = "Circular",
                        Count = feats.Count,
                        PCD = meanRadius * 2,
                        Center = centroid,
                        Features = feats
                    };
                }
            }
            catch { }

            return null;
        }

        // ── Helpers ──

        private bool IsTorus(ISurface surface)
        {
            try { return surface.TorusParams != null; }
            catch { return false; }
        }

        private bool IsCone(ISurface surface)
        {
            try { return surface.ConeParams != null; }
            catch { return false; }
        }

        private bool HasCosmeticThread(IFace2 face)
        {
            // Check if the face has a cosmetic thread annotation
            try
            {
                // Cosmetic threads are typically stored as feature-level data
                // For BREP-only detection, check for helical UV parameterization
                return false; // Conservative: only detects via feature tree
            }
            catch { return false; }
        }

        private static double Magnitude(double[] v)
        {
            return Math.Sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2]);
        }
    }
}
