using System;
using System.Collections.Generic;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Walks the SolidWorks feature tree and classifies each feature for
    /// the dimensioning engine. Determines if section/detail views are needed.
    /// </summary>
    public class FeatureAnalyzer
    {
        private readonly IProgressCallback _progress;

        // Feature types that indicate internal geometry (need section view)
        private static readonly HashSet<string> InternalFeatureTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "HoleWzd", "Cut", "CutExtrude", "RevCut", "Pocket",
            "ICE", "CutSweep", "CutLoft"
        };

        // Feature types that are typically small and benefit from detail views
        private static readonly HashSet<string> SmallFeatureTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Fillet", "ChamferFeature", "BreakCorner"
        };

        // Threshold below which we consider a feature "small" (meters)
        private const double DETAIL_VIEW_THRESHOLD = 0.005; // 5mm

        public FeatureAnalyzer(IProgressCallback progress = null)
        {
            _progress = progress;
        }

        /// <summary>
        /// Analyze all features in the model and return classified list.
        /// </summary>
        public List<AnalyzedFeature> Analyze(ModelDoc2 doc)
        {
            var features = new List<AnalyzedFeature>();
            bool hasInternalFeatures = false;
            bool hasSmallFeatures = false;

            Feature feat = (Feature)doc.FirstFeature();
            while (feat != null)
            {
                try
                {
                    string typeName = feat.GetTypeName2();
                    var analyzed = ClassifyFeature(feat, typeName);

                    if (analyzed != null)
                    {
                        features.Add(analyzed);

                        if (analyzed.IsInternal)
                            hasInternalFeatures = true;

                        if (analyzed.NeedsDetailView)
                            hasSmallFeatures = true;
                    }
                }
                catch (Exception ex)
                {
                    _progress?.LogWarning($"Could not analyze feature '{feat.Name}': {ex.Message}");
                }

                feat = (Feature)feat.GetNextFeature();
            }

            // If ANY internal feature found, mark the first internal feature as needing section
            if (hasInternalFeatures)
            {
                var firstInternal = features.Find(f => f.IsInternal);
                if (firstInternal != null)
                    firstInternal.NeedsSectionView = true;
            }

            _progress?.LogMessage($"Analyzed {features.Count} features. " +
                $"Internal: {(hasInternalFeatures ? "Yes" : "No")}, " +
                $"Small features: {(hasSmallFeatures ? "Yes" : "No")}");

            return features;
        }

        /// <summary>
        /// Classify a single feature into an AnalyzedFeature.
        /// </summary>
        private AnalyzedFeature ClassifyFeature(Feature feat, string typeName)
        {
            // Skip system/reference features
            if (IsSystemFeature(typeName))
                return null;

            var analyzed = new AnalyzedFeature
            {
                Name = feat.Name,
                TypeName = typeName,
                Category = MapCategory(typeName),
                IsInternal = InternalFeatureTypes.Contains(typeName)
            };

            // Extract dimensions from the feature's specific data object
            ExtractFeatureData(feat, typeName, analyzed);

            // Determine if detail view is needed based on bounding size
            if (SmallFeatureTypes.Contains(typeName) || analyzed.BoundingRadius < DETAIL_VIEW_THRESHOLD)
            {
                analyzed.NeedsDetailView = analyzed.BoundingRadius > 0 &&
                                            analyzed.BoundingRadius < DETAIL_VIEW_THRESHOLD;
            }

            return analyzed;
        }

        /// <summary>
        /// Extract dimensional data from feature definition objects.
        /// </summary>
        private void ExtractFeatureData(Feature feat, string typeName, AnalyzedFeature analyzed)
        {
            try
            {
                object defObj = feat.GetDefinition();

                switch (typeName)
                {
                    case "HoleWzd":
                        ExtractHoleWizardData(feat, analyzed);
                        break;

                    case "Fillet":
                    case "ConstRadiusFillet":
                        ExtractFilletData(feat, analyzed);
                        break;

                    case "ChamferFeature":
                        ExtractChamferData(feat, analyzed);
                        break;

                    case "LPattern":
                        ExtractLinearPatternData(feat, analyzed);
                        break;

                    case "CirPattern":
                        ExtractCircularPatternData(feat, analyzed);
                        break;

                    case "Cut":
                    case "Boss":
                    case "CutExtrude":
                    case "BossExtrude":
                        ExtractExtrudeData(feat, analyzed);
                        break;

                    case "SMBaseFlange":
                    case "EdgeFlange":
                    case "ICE":
                        analyzed.Category = FeatureCategory.SheetMetalFlange;
                        break;

                    case "Thread":
                        ExtractThreadData(feat, analyzed);
                        break;
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning(
                    $"Could not extract data from feature '{feat.Name}' ({typeName}): {ex.Message}");
            }
        }

        private void ExtractHoleWizardData(Feature feat, AnalyzedFeature analyzed)
        {
            analyzed.Category = FeatureCategory.HoleWizard;
            analyzed.IsInternal = true;

            // Try to get hole wizard feature data
            try
            {
                object defObj = feat.GetDefinition();
                if (defObj is IWizardHoleFeatureData2 holeData)
                {
                    analyzed.Diameter = holeData.Diameter;
                    analyzed.Depth = holeData.Depth;
                    analyzed.BoundingRadius = holeData.Diameter / 2.0;

                    // Check for tapped hole types
                    try
                    {
                        int holeType = holeData.Type;
                        // Tapped hole types: 3 (Straight Tap), 4 (Tapered Tap), 5 (Bottom Tap)
                        if (holeType >= 3 && holeType <= 5)
                        {
                            analyzed.Category = FeatureCategory.Thread;
                            analyzed.ThreadCallout = $"M{holeData.Diameter * 1000:F1}";
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // Fallback: try reading dimensions from the feature directly
            if (analyzed.Diameter == 0)
            {
                TryReadDimensionsFromFeature(feat, analyzed);
            }
        }

        private void ExtractFilletData(Feature feat, AnalyzedFeature analyzed)
        {
            analyzed.Category = FeatureCategory.Fillet;

            try
            {
                object defObj = feat.GetDefinition();
                if (defObj is ISimpleFilletFeatureData2 filletData)
                {
                    analyzed.Radius = filletData.DefaultRadius;
                    analyzed.BoundingRadius = filletData.DefaultRadius;
                }
            }
            catch { }

            if (analyzed.Radius == 0)
                TryReadDimensionsFromFeature(feat, analyzed);
        }

        private void ExtractChamferData(Feature feat, AnalyzedFeature analyzed)
        {
            analyzed.Category = FeatureCategory.Chamfer;

            // IChamferFeatureData2 property names vary across SW versions;
            // use display dimension fallback which works universally.
            TryReadDimensionsFromFeature(feat, analyzed);
        }

        private void ExtractLinearPatternData(Feature feat, AnalyzedFeature analyzed)
        {
            analyzed.Category = FeatureCategory.LinearPattern;

            try
            {
                object defObj = feat.GetDefinition();
                if (defObj is ILinearPatternFeatureData patternData)
                {
                    analyzed.PatternCount = patternData.D1TotalInstances;
                    analyzed.PatternSpacing = patternData.D1Spacing;
                }
            }
            catch { }
        }

        private void ExtractCircularPatternData(Feature feat, AnalyzedFeature analyzed)
        {
            analyzed.Category = FeatureCategory.CircularPattern;

            try
            {
                object defObj = feat.GetDefinition();
                if (defObj is ICircularPatternFeatureData patternData)
                {
                    analyzed.PatternCount = patternData.TotalInstances;
                    analyzed.PatternSpacing = patternData.Spacing;
                }
            }
            catch { }
        }

        private void ExtractExtrudeData(Feature feat, AnalyzedFeature analyzed)
        {
            bool isCut = feat.GetTypeName2().Contains("Cut");
            analyzed.Category = isCut ? FeatureCategory.CutExtrude : FeatureCategory.BossExtrude;
            analyzed.IsInternal = isCut;

            try
            {
                object defObj = feat.GetDefinition();
                if (defObj is IExtrudeFeatureData2 extData)
                {
                    analyzed.Depth = extData.GetDepth(true); // direction 1
                }
            }
            catch { }
        }

        private void ExtractThreadData(Feature feat, AnalyzedFeature analyzed)
        {
            analyzed.Category = FeatureCategory.Thread;
            analyzed.IsInternal = true;

            // Thread features: extract dims via display dimensions fallback
            // since IThreadFeatureData property names vary across versions
            TryReadDimensionsFromFeature(feat, analyzed);
            if (analyzed.Diameter > 0)
            {
                analyzed.ThreadCallout = $"M{analyzed.Diameter * 1000:F0}";
                analyzed.BoundingRadius = analyzed.Diameter / 2.0;
            }
        }

        /// <summary>
        /// Fallback: read dimensions directly from the feature's sub-features/display dimensions.
        /// </summary>
        private void TryReadDimensionsFromFeature(Feature feat, AnalyzedFeature analyzed)
        {
            try
            {
                DisplayDimension dispDim = (DisplayDimension)feat.GetFirstDisplayDimension();
                while (dispDim != null)
                {
                    Dimension dim = (Dimension)dispDim.GetDimension2(0);
                    if (dim != null)
                    {
                        double val = dim.Value;
                        string name = dim.Name?.ToLowerInvariant() ?? "";

                        if (name.Contains("dia") || name.Contains("d1"))
                            analyzed.Diameter = val;
                        else if (name.Contains("depth"))
                            analyzed.Depth = val;
                        else if (name.Contains("radius") || name.Contains("rad"))
                            analyzed.Radius = val;
                        else if (name.Contains("angle"))
                            analyzed.Angle = val;
                        else if (analyzed.Distance1 == 0)
                            analyzed.Distance1 = val;
                        else
                            analyzed.Distance2 = val;

                        if (analyzed.BoundingRadius == 0 && val > 0)
                            analyzed.BoundingRadius = val;
                    }

                    dispDim = (DisplayDimension)feat.GetNextDisplayDimension(dispDim);
                }
            }
            catch { }
        }

        private FeatureCategory MapCategory(string typeName)
        {
            switch (typeName)
            {
                case "HoleWzd": return FeatureCategory.HoleWizard;
                case "Cut":
                case "CutExtrude": return FeatureCategory.CutExtrude;
                case "Boss":
                case "BossExtrude": return FeatureCategory.BossExtrude;
                case "Fillet":
                case "ConstRadiusFillet": return FeatureCategory.Fillet;
                case "ChamferFeature": return FeatureCategory.Chamfer;
                case "LPattern": return FeatureCategory.LinearPattern;
                case "CirPattern": return FeatureCategory.CircularPattern;
                case "Thread": return FeatureCategory.Thread;
                case "SMBaseFlange":
                case "EdgeFlange": return FeatureCategory.SheetMetalFlange;
                case "ICE":
                case "SketchBend": return FeatureCategory.SheetMetalBend;
                case "Shell": return FeatureCategory.Shell;
                case "Mirror": return FeatureCategory.Mirror;
                case "Rib": return FeatureCategory.Rib;
                default: return FeatureCategory.Other;
            }
        }

        private bool IsSystemFeature(string typeName)
        {
            // System/reference features that we skip during analysis
            return typeName == "OriginProfileFeature" ||
                   typeName == "RefPlane" ||
                   typeName == "RefAxis" ||
                   typeName == "MaterialFolder" ||
                   typeName == "HistoryFolder" ||
                   typeName == "SensorFolder" ||
                   typeName == "DetailCabinet" ||
                   typeName == "SolidBodyFolder" ||
                   typeName == "SurfaceBodyFolder" ||
                   typeName == "ProfileFeature" ||
                   typeName == "3DProfileFeature" ||
                   typeName == "MateGroup" ||
                   typeName == "FlatPattern" ||
                   typeName == "CommentsFolder" ||
                   typeName == "FavoriteFolder" ||
                   typeName == "DesignTableFolder" ||
                   typeName == "SelectionSetFolder" ||
                   typeName == "BendTableFolder";
        }
    }
}
