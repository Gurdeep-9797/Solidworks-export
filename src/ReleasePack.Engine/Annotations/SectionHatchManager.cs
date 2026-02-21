using System;
using SolidWorks.Interop.sldworks;

namespace ReleasePack.Engine.Annotations
{
    /// <summary>
    /// Enforces strict section view compliance.
    /// Alternates hatch angles on multi-body/assembly cut faces and applies correct scaling.
    /// </summary>
    public class SectionHatchManager
    {
        private const double ANGLE_45_RAD = Math.PI / 4.0;
        private const double ANGLE_135_RAD = 3.0 * Math.PI / 4.0;

        /// <summary>
        /// Analyzes a generated section view and forces standardized 45/135 degree alternating hatches.
        /// </summary>
        public void FormatSectionViewHatches(IView sectionView, double viewScale)
        {
            if (sectionView == null) return;

            // Get all hatches in this section view
            object[] faceHatches = (object[])sectionView.GetFaceHatches();
            if (faceHatches == null || faceHatches.Length == 0) return;

            bool useAlternateAngle = false;

            // Iterate through every detected hatch face
            foreach (var hatchObj in faceHatches)
            {
                IFaceHatch hatch = hatchObj as IFaceHatch;
                if (hatch != null)
                {
                    // 1. Force the ISO standard hatch pattern "ANSI31" (Iron/General) if available
                    hatch.Pattern = "ANSI31";
                    
                    // 2. Alternate angles between touching components (45 and 135 degrees)
                    hatch.Angle = useAlternateAngle ? ANGLE_135_RAD : ANGLE_45_RAD;
                    useAlternateAngle = !useAlternateAngle;

                    // 3. Set the scale proportional to the View Scale. 
                    // This prevents high-scale detail views from having solid-black tight hatches.
                    // Scale factor 1.0 is default. If view is 1:5, hatch scale should compensate up.
                    double inverseScale = 1.0 / viewScale;
                    double idealHatchScale = Math.Max(1.0, inverseScale * 0.5); // Adjust multiplier as aesthetically needed

                    hatch.Scale2 = idealHatchScale;

                    // Note: Hatches inherit the View's layer natively unless components are explicitly overridden.
                }
            }
        }
    }
}
