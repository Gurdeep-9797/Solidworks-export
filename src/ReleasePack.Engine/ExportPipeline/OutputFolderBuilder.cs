using System;
using System.IO;

namespace ReleasePack.Engine.Exporting
{
    /// <summary>
    /// V2 Pipeline to enforce rigid output directory structures.
    /// Eliminates file scatter and ensures all deliverables are cleanly sorted by type.
    /// </summary>
    public class OutputFolderBuilder
    {
        public string RootPath { get; private set; }

        public string DrawingsFolder { get; private set; }
        public string PdfFolder { get; private set; }
        public string DxfFolder { get; private set; }
        public string StepFolder { get; private set; }
        public string ParasolidFolder { get; private set; }
        public string BomFolder { get; private set; }

        public OutputFolderBuilder(string basePath, string projectName)
        {
            if (string.IsNullOrWhiteSpace(basePath)) throw new ArgumentNullException(nameof(basePath));
            if (string.IsNullOrWhiteSpace(projectName)) projectName = "Release";

            // Root structure: C:\...\Release_YYYYMMDD_ProjectName
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
            RootPath = Path.Combine(basePath, $"{projectName}_Release_{timestamp}");

            // Define Subfolders
            DrawingsFolder = Path.Combine(RootPath, "01_Drawings_SLDDRW");
            PdfFolder = Path.Combine(RootPath, "02_PDFs");
            StepFolder = Path.Combine(RootPath, "03_STEP");
            DxfFolder = Path.Combine(RootPath, "04_DXF");
            ParasolidFolder = Path.Combine(RootPath, "05_Parasolid");
            BomFolder = Path.Combine(RootPath, "06_BOM");
        }

        /// <summary>
        /// Generates the physical folders on the disk.
        /// </summary>
        public void BuildTree()
        {
            if (!Directory.Exists(RootPath)) Directory.CreateDirectory(RootPath);

            Directory.CreateDirectory(DrawingsFolder);
            Directory.CreateDirectory(PdfFolder);
            Directory.CreateDirectory(StepFolder);
            Directory.CreateDirectory(DxfFolder);
            Directory.CreateDirectory(ParasolidFolder);
            Directory.CreateDirectory(BomFolder);
        }

        /// <summary>
        /// Retrieves the target path based on extension.
        /// </summary>
        public string GetTargetFolder(string extension)
        {
            extension = extension.ToLower().TrimStart('.');
            switch (extension)
            {
                case "slddrw": return DrawingsFolder;
                case "pdf": return PdfFolder;
                case "step":
                case "stp": return StepFolder;
                case "dxf": return DxfFolder;
                case "x_t": return ParasolidFolder;
                case "xlsx":
                case "csv": return BomFolder;
                default: return RootPath;
            }
        }
    }
}
