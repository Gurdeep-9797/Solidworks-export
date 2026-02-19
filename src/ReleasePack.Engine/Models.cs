using System.Collections.Generic;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Represents a node in the assembly dependency tree.
    /// </summary>
    public class ModelNode
    {
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public ModelNodeType NodeType { get; set; }
        public string ConfigurationName { get; set; }
        public Dictionary<string, string> CustomProperties { get; set; } = new Dictionary<string, string>();
        public List<ModelNode> Children { get; set; } = new List<ModelNode>();

        /// <summary>Convenience property: PartNumber from custom properties or filename.</summary>
        public string PartNumber =>
            CustomProperties.ContainsKey("PartNumber") ? CustomProperties["PartNumber"] :
            CustomProperties.ContainsKey("Part Number") ? CustomProperties["Part Number"] :
            System.IO.Path.GetFileNameWithoutExtension(FilePath);

        /// <summary>Revision from custom properties, defaults to "A".</summary>
        public string Revision =>
            CustomProperties.ContainsKey("Revision") ? CustomProperties["Revision"] :
            CustomProperties.ContainsKey("Rev") ? CustomProperties["Rev"] : "A";

        /// <summary>Description from custom properties.</summary>
        public string Description =>
            CustomProperties.ContainsKey("Description") ? CustomProperties["Description"] : "";

        /// <summary>Material from custom properties.</summary>
        public string Material =>
            CustomProperties.ContainsKey("Material") ? CustomProperties["Material"] : "";

        /// <summary>Whether this is a sheet metal part (detected during scanning).</summary>
        public bool IsSheetMetal { get; set; }
    }

    public enum ModelNodeType
    {
        Part,
        Assembly,
        Drawing
    }

    /// <summary>
    /// Analyzed feature information from the feature tree.
    /// </summary>
    public class AnalyzedFeature
    {
        public string Name { get; set; }
        public FeatureCategory Category { get; set; }
        public string TypeName { get; set; }

        // Dimension data extracted from the feature
        public double Diameter { get; set; }
        public double Depth { get; set; }
        public double Radius { get; set; }
        public double Angle { get; set; }
        public double Distance1 { get; set; }
        public double Distance2 { get; set; }
        public int PatternCount { get; set; }
        public double PatternSpacing { get; set; }

        // Position info
        public double PositionX { get; set; }
        public double PositionY { get; set; }
        public double PositionZ { get; set; }

        // Bounding size of the feature (for deciding detail views)
        public double BoundingRadius { get; set; }

        // Flags
        public bool NeedsSectionView { get; set; }
        public bool NeedsDetailView { get; set; }
        public bool IsInternal { get; set; }
        public bool NeedsAuxiliaryView { get; set; }
        public double AuxiliaryAngle { get; set; }

        // TYP grouping (fillets/chamfers of same size)
        public bool IsTypical { get; set; }
        public string TypicalGroupId { get; set; }

        /// <summary>Thread callout string if applicable (e.g., "M8x1.25").</summary>
        public string ThreadCallout { get; set; }
    }

    public enum FeatureCategory
    {
        HoleWizard,
        CutExtrude,
        BossExtrude,
        Fillet,
        Chamfer,
        LinearPattern,
        CircularPattern,
        Thread,
        SheetMetalBend,
        SheetMetalFlange,
        Pocket,
        Slot,
        Rib,
        Shell,
        Mirror,
        CounterBore,
        CounterSink,
        Sweep,
        Loft,
        Other
    }

    /// <summary>
    /// User-selected options for the release pack generation.
    /// </summary>
    public class ExportOptions
    {
        // Scope
        public ExportScope Scope { get; set; } = ExportScope.CurrentDocument;
        public string RemoteFilePath { get; set; }
        
        // Project Metadata (for Title Block)
        public string CompanyName { get; set; }
        public string ProjectName { get; set; }
        public string DrawnBy { get; set; }
        public string CheckedBy { get; set; }

        // Output types
        public bool GenerateDrawing { get; set; } = true;
        public bool ExportPDF { get; set; } = true;
        public bool ExportDXF { get; set; } = true;
        public bool ExportSTEP { get; set; } = false;
        public bool ExportParasolid { get; set; } = false;
        public bool ExportBOM { get; set; } = true;
        public bool ExportPreviewImage { get; set; } = false;

        // Drawing options
        public ViewStandard ViewStandard { get; set; } = ViewStandard.ThirdAngle;
        public SheetSizeOption SheetSize { get; set; } = SheetSizeOption.Auto;
        public string DrawingTemplatePath { get; set; }
        public string BomTemplatePath { get; set; }

        // Output folder
        public string OutputFolder { get; set; } // null = auto (next to source file)
        public bool UseCustomFolder { get; set; } = false;
    }

    public enum ExportScope
    {
        CurrentDocument,
        CurrentAndChildren,
        RemoteFile
    }

    public enum ViewStandard
    {
        ThirdAngle,  // ANSI
        FirstAngle   // ISO
    }

    public enum SheetSizeOption
    {
        Auto,
        A4_Landscape,
        A3_Landscape,
        A2_Landscape,
        A1_Landscape,
        A0_Landscape
    }

    /// <summary>
    /// Progress callback interface for UI updates.
    /// </summary>
    public interface IProgressCallback
    {
        void ReportProgress(int percent, string message);
        void LogMessage(string message);
        void LogWarning(string message);
        void LogError(string message);
    }

    /// <summary>
    /// Result of a single export operation.
    /// </summary>
    public class ExportResult
    {
        public string SourceFile { get; set; }
        public string OutputFile { get; set; }
        public string ExportType { get; set; }
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
    }
}
