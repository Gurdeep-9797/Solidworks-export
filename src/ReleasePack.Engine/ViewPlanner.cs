using System;
using System.Collections.Generic;
using System.Linq;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Decides which drawing views are needed based on feature analysis.
    /// Follows ISO 128 / ASME Y14.3 standards for view selection.
    /// </summary>
    public class ViewPlanner
    {
        /// <summary>
        /// Result of the view planning process.
        /// </summary>
        public class ViewPlan
        {
            public bool NeedFrontView { get; set; } = true;  // Always
            public bool NeedTopView { get; set; } = true;
            public bool NeedRightView { get; set; } = true;
            public bool NeedIsoView { get; set; } = true;
            public bool NeedSectionView { get; set; }
            public bool NeedDetailViews { get; set; }
            public List<double> AuxiliaryAngles { get; set; } = new List<double>();
            public List<AnalyzedFeature> DetailFeatures { get; set; } = new List<AnalyzedFeature>();
            
            // TYP groups: key = groupId, value = list of features sharing the same size
            public Dictionary<string, List<AnalyzedFeature>> TypicalGroups { get; set; } 
                = new Dictionary<string, List<AnalyzedFeature>>();

            /// <summary>
            /// Summary string for logging.
            /// </summary>
            public string Summary
            {
                get
                {
                    int viewCount = 1; // Front always
                    if (NeedTopView) viewCount++;
                    if (NeedRightView) viewCount++;
                    if (NeedIsoView) viewCount++;
                    if (NeedSectionView) viewCount++;
                    viewCount += AuxiliaryAngles.Count;
                    
                    return $"{viewCount} views planned" +
                           (NeedSectionView ? " (incl. Section)" : "") +
                           (NeedDetailViews ? $" + {DetailFeatures.Count} detail(s)" : "") +
                           (AuxiliaryAngles.Count > 0 ? $" + {AuxiliaryAngles.Count} auxiliary" : "") +
                           (TypicalGroups.Count > 0 ? $" | {TypicalGroups.Count} TYP groups" : "");
                }
            }
        }

        /// <summary>
        /// Analyze features and model geometry to determine the optimal set of views.
        /// </summary>
        public static ViewPlan Plan(List<AnalyzedFeature> features, double modelW, double modelH, double modelD)
        {
            var plan = new ViewPlan();

            if (features == null || features.Count == 0)
                return plan;

            // --- 1. Section View Decision ---
            // If any internal features exist (holes, cuts, pockets), add section view
            plan.NeedSectionView = features.Any(f => f.IsInternal);

            // --- 2. Detail View Decision ---
            var detailFeatures = features.Where(f => f.NeedsDetailView).ToList();
            plan.NeedDetailViews = detailFeatures.Count > 0;
            plan.DetailFeatures = detailFeatures;

            // --- 3. Auxiliary View Decision ---
            var auxFeatures = features.Where(f => f.NeedsAuxiliaryView).ToList();
            foreach (var f in auxFeatures)
            {
                if (f.AuxiliaryAngle != 0 && !plan.AuxiliaryAngles.Contains(f.AuxiliaryAngle))
                    plan.AuxiliaryAngles.Add(f.AuxiliaryAngle);
            }

            // --- 4. View Reduction for Simple Parts ---
            // Cylindrical/turned parts: If model is predominantly revolved,
            // only Front + one Side view is needed (2 views).
            // Heuristic: If depth and width are similar (within 20%) AND 
            // most features are internal, it's likely cylindrical.
            bool likelyCylindrical = Math.Abs(modelW - modelD) / Math.Max(modelW, 0.001) < 0.2
                                     && features.Count(f => f.IsInternal) > features.Count * 0.5;

            if (likelyCylindrical && features.Count < 10)
            {
                // Simple turned part: Front + Right is sufficient
                plan.NeedTopView = false;
            }

            // Very simple parts (< 3 features, no internal): Front + Iso only
            int meaningfulFeatures = features.Count(f => 
                f.Category != FeatureCategory.Other && 
                f.Category != FeatureCategory.Fillet && 
                f.Category != FeatureCategory.Chamfer);
            
            if (meaningfulFeatures <= 2 && !plan.NeedSectionView)
            {
                // Simple extruded part: 3 views usually enough, but always include Iso
                // Keep all 3 standard views for safety
            }

            // --- 5. TYP Grouping ---
            // Group fillets by radius, chamfers by size, holes by diameter
            GroupTypicalFeatures(features, plan);

            return plan;
        }

        /// <summary>
        /// Group features of the same type and size for "TYP" labeling.
        /// E.g., 4 fillets with R=3mm → label one as "R3 TYP"
        /// </summary>
        private static void GroupTypicalFeatures(List<AnalyzedFeature> features, ViewPlan plan)
        {
            // Group fillets by radius
            var fillets = features.Where(f => f.Category == FeatureCategory.Fillet && f.Radius > 0)
                                  .GroupBy(f => Math.Round(f.Radius * 1000, 1)); // Round to 0.1mm

            foreach (var group in fillets.Where(g => g.Count() > 1))
            {
                string groupId = $"FILLET_R{group.Key:F1}";
                var list = group.ToList();
                plan.TypicalGroups[groupId] = list;
                
                // Mark all as typical, first one gets the label
                foreach (var f in list)
                {
                    f.IsTypical = true;
                    f.TypicalGroupId = groupId;
                }
            }

            // Group chamfers by size
            var chamfers = features.Where(f => f.Category == FeatureCategory.Chamfer && f.Distance1 > 0)
                                   .GroupBy(f => Math.Round(f.Distance1 * 1000, 1));

            foreach (var group in chamfers.Where(g => g.Count() > 1))
            {
                string groupId = $"CHAMFER_{group.Key:F1}";
                var list = group.ToList();
                plan.TypicalGroups[groupId] = list;
                
                foreach (var f in list)
                {
                    f.IsTypical = true;
                    f.TypicalGroupId = groupId;
                }
            }

            // Group holes by diameter
            var holes = features.Where(f => (f.Category == FeatureCategory.HoleWizard || 
                                             f.Category == FeatureCategory.CutExtrude) && f.Diameter > 0)
                                .GroupBy(f => Math.Round(f.Diameter * 1000, 1));

            foreach (var group in holes.Where(g => g.Count() > 1))
            {
                string groupId = $"HOLE_D{group.Key:F1}";
                var list = group.ToList();
                plan.TypicalGroups[groupId] = list;
                
                foreach (var f in list)
                {
                    f.IsTypical = true;
                    f.TypicalGroupId = groupId;
                }
            }
        }
    }
}
