using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Orchestrates the full Release Pack generation pipeline.
    /// For each model in the worktree, runs the selected exports.
    /// </summary>
    public class ExportPipeline
    {
        private readonly ISldWorks _swApp;
        private readonly IProgressCallback _progress;
        private readonly DrawingGenerator _drawingGen;
        private readonly AssemblyDrawingGenerator _assyDrawingGen;
        private readonly BomExtractor _bomExtractor;
        private readonly DependencyScanner _scanner;

        public ExportPipeline(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
            _drawingGen = new DrawingGenerator(swApp, progress);
            _assyDrawingGen = new AssemblyDrawingGenerator(swApp, progress);
            _bomExtractor = new BomExtractor(swApp, progress);
            _scanner = new DependencyScanner(swApp, progress);
        }

        /// <summary>
        /// Run the full export pipeline.
        /// Returns list of all generated file results.
        /// </summary>
        public List<ExportResult> Execute(ExportOptions options)
        {
            var results = new List<ExportResult>();

            try
            {
                // ── Step 1: Scan dependencies ─────────────────────
                _progress?.ReportProgress(5, "Scanning dependencies...");
                List<ModelNode> worktree = _scanner.Scan(options);

                if (worktree == null || worktree.Count == 0)
                {
                    _progress?.LogError("No models found to export.");
                    return results;
                }

                // Flatten to unique file list
                List<ModelNode> allNodes = DependencyScanner.Flatten(worktree);
                _progress?.LogMessage($"Found {allNodes.Count} unique model(s) in worktree.");

                // ── Step 2: Determine output folder ───────────────
                string outputFolder = DetermineOutputFolder(options, worktree[0]);
                Directory.CreateDirectory(outputFolder);
                _progress?.LogMessage($"Output folder: {outputFolder}");

                int totalSteps = allNodes.Count * CountSelectedOutputs(options);
                int currentStep = 0;

                // ── Step 3: Process each model ────────────────────
                foreach (var node in allNodes)
                {
                    _progress?.LogMessage($"\n━━━ Processing: {node.FileName} ━━━");

                    // Drawing generation
                    if (options.GenerateDrawing)
                    {
                        currentStep++;
                        _progress?.ReportProgress(ProgressPercent(currentStep, totalSteps),
                            $"Generating drawing: {node.FileName}");

                        string drawingPath;
                        if (node.NodeType == ModelNodeType.Assembly)
                        {
                            drawingPath = _assyDrawingGen.Generate(node, outputFolder, options);
                        }
                        else
                        {
                            drawingPath = _drawingGen.Generate(node, outputFolder, options);
                        }

                        results.Add(new ExportResult
                        {
                            SourceFile = node.FilePath,
                            OutputFile = drawingPath,
                            ExportType = "Drawing",
                            Success = drawingPath != null
                        });
                    }

                    // PDF export
                    if (options.ExportPDF)
                    {
                        currentStep++;
                        _progress?.ReportProgress(ProgressPercent(currentStep, totalSteps),
                            $"Exporting PDF: {node.FileName}");

                        var result = ExportFile(node, outputFolder, ".pdf",
                            (int)swSaveAsVersion_e.swSaveAsCurrentVersion, "PDF");
                        results.Add(result);
                    }

                    // DXF export
                    if (options.ExportDXF)
                    {
                        currentStep++;
                        _progress?.ReportProgress(ProgressPercent(currentStep, totalSteps),
                            $"Exporting DXF: {node.FileName}");

                        if (node.IsSheetMetal)
                        {
                            var result = ExportSheetMetalDxf(node, outputFolder);
                            results.Add(result);
                        }
                        else
                        {
                            var result = ExportFile(node, outputFolder, ".dxf",
                                (int)swSaveAsVersion_e.swSaveAsCurrentVersion, "DXF");
                            results.Add(result);
                        }
                    }

                    // STEP export
                    if (options.ExportSTEP)
                    {
                        currentStep++;
                        _progress?.ReportProgress(ProgressPercent(currentStep, totalSteps),
                            $"Exporting STEP: {node.FileName}");

                        var result = ExportFile(node, outputFolder, ".step",
                            (int)swSaveAsVersion_e.swSaveAsCurrentVersion, "STEP");
                        results.Add(result);
                    }

                    // Parasolid export
                    if (options.ExportParasolid)
                    {
                        currentStep++;
                        _progress?.ReportProgress(ProgressPercent(currentStep, totalSteps),
                            $"Exporting Parasolid: {node.FileName}");

                        var result = ExportFile(node, outputFolder, ".x_t",
                            (int)swSaveAsVersion_e.swSaveAsCurrentVersion, "Parasolid");
                        results.Add(result);
                    }

                    // Preview image
                    if (options.ExportPreviewImage)
                    {
                        currentStep++;
                        _progress?.ReportProgress(ProgressPercent(currentStep, totalSteps),
                            $"Capturing preview: {node.FileName}");

                        var result = ExportPreviewImage(node, outputFolder);
                        results.Add(result);
                    }
                }

                // ── Step 4: BOM export (assembly-level only) ──────
                if (options.ExportBOM)
                {
                    var assemblies = allNodes.Where(n => n.NodeType == ModelNodeType.Assembly).ToList();
                    foreach (var assy in assemblies)
                    {
                        _progress?.ReportProgress(95, $"Exporting BOM: {assy.FileName}");
                        string bomPath = _bomExtractor.Export(assy, outputFolder);
                        results.Add(new ExportResult
                        {
                            SourceFile = assy.FilePath,
                            OutputFile = bomPath,
                            ExportType = "BOM Excel",
                            Success = bomPath != null
                        });
                    }
                }

                // ── Summary ───────────────────────────────────────
                int success = results.Count(r => r.Success);
                int failed = results.Count(r => !r.Success);
                _progress?.ReportProgress(100, "Complete!");
                _progress?.LogMessage($"\n══════════════════════════════════════");
                _progress?.LogMessage($"Release Pack Complete!");
                _progress?.LogMessage($"  Successful: {success}");
                _progress?.LogMessage($"  Failed:     {failed}");
                _progress?.LogMessage($"  Output:     {outputFolder}");
                _progress?.LogMessage($"══════════════════════════════════════");
            }
            catch (Exception ex)
            {
                _progress?.LogError($"Pipeline failed: {ex.Message}\n{ex.StackTrace}");
            }

            return results;
        }

        /// <summary>
        /// Export a model file to a specified format using SaveAs.
        /// </summary>
        private ExportResult ExportFile(ModelNode node, string outputFolder,
            string extension, int saveVersion, string typeName)
        {
            var result = new ExportResult
            {
                SourceFile = node.FilePath,
                ExportType = typeName,
                Success = false
            };

            try
            {
                ModelDoc2 doc = OpenOrGetModel(node.FilePath);
                if (doc == null)
                {
                    result.ErrorMessage = "Could not open model.";
                    return result;
                }

                string outputPath = DrawingGenerator.GetOutputPath(node, outputFolder, extension);
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                int errors = 0, warnings = 0;
                bool saved = doc.Extension.SaveAs2(
                    outputPath, saveVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null, "", false, ref errors, ref warnings);

                result.OutputFile = outputPath;
                result.Success = saved;

                if (!saved)
                    result.ErrorMessage = $"SaveAs returned false. Errors: {errors}";

                _progress?.LogMessage($"  {typeName}: {(saved ? "✓" : "✗")} {Path.GetFileName(outputPath)}");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                _progress?.LogWarning($"  {typeName}: ✗ {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Export sheet metal flat pattern DXF using ExportToDWG2.
        /// </summary>
        private ExportResult ExportSheetMetalDxf(ModelNode node, string outputFolder)
        {
            var result = new ExportResult
            {
                SourceFile = node.FilePath,
                ExportType = "DXF (Flat Pattern)",
                Success = false
            };

            try
            {
                ModelDoc2 doc = OpenOrGetModel(node.FilePath);
                if (doc == null || !(doc is PartDoc partDoc))
                {
                    result.ErrorMessage = "Could not open as part document.";
                    return result;
                }

                string outputPath = DrawingGenerator.GetOutputPath(node, outputFolder, ".dxf");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                // ExportToDWG2 parameters for flat pattern
                // dataAlignment: 1 = align to geometry
                // views: 1 = flat pattern view
                bool exported = partDoc.ExportToDWG2(
                    outputPath,
                    node.FilePath,
                    (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal,
                    true,   // Include bend lines
                    null,   // Use default settings
                    false,  // Don't show map
                    false,  // Don't overwrite
                    0,      // Sheet index
                    null    // Views
                );

                result.OutputFile = outputPath;
                result.Success = exported;

                if (!exported)
                {
                    // Fallback: try regular DXF export
                    _progress?.LogWarning("Flat pattern export failed. Trying regular DXF...");
                    return ExportFile(node, outputFolder, ".dxf",
                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion, "DXF");
                }

                _progress?.LogMessage($"  DXF (Flat): ✓ {Path.GetFileName(outputPath)}");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                _progress?.LogWarning($"  DXF (Flat): ✗ {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Capture a preview PNG image of the model.
        /// </summary>
        private ExportResult ExportPreviewImage(ModelNode node, string outputFolder)
        {
            var result = new ExportResult
            {
                SourceFile = node.FilePath,
                ExportType = "Preview PNG",
                Success = false
            };

            try
            {
                ModelDoc2 doc = OpenOrGetModel(node.FilePath);
                if (doc == null)
                {
                    result.ErrorMessage = "Could not open model.";
                    return result;
                }

                string outputPath = DrawingGenerator.GetOutputPath(node, outputFolder, ".png");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                // Use SaveBMP with PNG extension (SolidWorks auto-detects format)
                bool saved = doc.SaveBMP(outputPath, 1920, 1080);

                result.OutputFile = outputPath;
                result.Success = saved;

                _progress?.LogMessage($"  Preview: {(saved ? "✓" : "✗")} {Path.GetFileName(outputPath)}");
            }
            catch (Exception ex)
            {
                result.ErrorMessage = ex.Message;
                _progress?.LogWarning($"  Preview: ✗ {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Determine the output folder path.
        /// </summary>
        private string DetermineOutputFolder(ExportOptions options, ModelNode rootNode)
        {
            if (options.UseCustomFolder && !string.IsNullOrEmpty(options.OutputFolder))
            {
                return options.OutputFolder;
            }

            // Auto: create folder next to source file
            string sourceDir = Path.GetDirectoryName(rootNode.FilePath);
            string folderName = Path.GetFileNameWithoutExtension(rootNode.FilePath) + "_ReleasePack";
            return Path.Combine(sourceDir, folderName);
        }

        private int CountSelectedOutputs(ExportOptions options)
        {
            int count = 0;
            if (options.GenerateDrawing) count++;
            if (options.ExportPDF) count++;
            if (options.ExportDXF) count++;
            if (options.ExportSTEP) count++;
            if (options.ExportParasolid) count++;
            if (options.ExportPreviewImage) count++;
            return Math.Max(count, 1);
        }

        private int ProgressPercent(int current, int total)
        {
            if (total == 0) return 100;
            return Math.Min(95, (int)(5 + 90.0 * current / total));
        }

        private ModelDoc2 OpenOrGetModel(string filePath)
        {
            ModelDoc2 doc = (ModelDoc2)_swApp.GetOpenDocumentByName(filePath);
            if (doc != null) return doc;

            int errors = 0, warnings = 0;
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            int docType = ext == ".sldasm" ? (int)swDocumentTypes_e.swDocASSEMBLY
                        : ext == ".slddrw" ? (int)swDocumentTypes_e.swDocDRAWING
                        : (int)swDocumentTypes_e.swDocPART;

            return (ModelDoc2)_swApp.OpenDoc6(filePath, docType,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                "", ref errors, ref warnings);
        }
    }
}
