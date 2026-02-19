using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Extracts BOM data from SolidWorks assemblies and exports to Excel (xlsx)
    /// with industrial-grade formatting using ClosedXML.
    /// </summary>
    public class BomExtractor
    {
        private readonly ISldWorks _swApp;
        private readonly IProgressCallback _progress;

        public BomExtractor(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
        }

        /// <summary>
        /// Extract BOM from an assembly and export to Excel.
        /// Returns the path to the generated xlsx file.
        /// </summary>
        public string Export(ModelNode assemblyNode, string outputFolder)
        {
            _progress?.LogMessage($"── Extracting BOM for: {assemblyNode.FileName} ──");

            if (assemblyNode.NodeType != ModelNodeType.Assembly)
            {
                _progress?.LogWarning("BOM extraction skipped: not an assembly.");
                return null;
            }

            try
            {
                // Collect BOM rows from the worktree
                var bomRows = CollectBomRows(assemblyNode);

                if (bomRows.Count == 0)
                {
                    _progress?.LogWarning("No BOM items found.");
                    return null;
                }

                // Generate Excel file
                string outputPath = DrawingGenerator.GetOutputPath(assemblyNode, outputFolder, ".xlsx");
                Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

                WriteExcel(bomRows, assemblyNode, outputPath);

                _progress?.LogMessage($"BOM exported: {outputPath} ({bomRows.Count} items)");
                return outputPath;
            }
            catch (Exception ex)
            {
                _progress?.LogError($"BOM extraction failed: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Collect BOM rows from the worktree by traversing assembly children.
        /// </summary>
        private List<BomRow> CollectBomRows(ModelNode assemblyNode)
        {
            var rows = new List<BomRow>();
            var partCounts = new Dictionary<string, BomRow>(StringComparer.OrdinalIgnoreCase);
            int itemNumber = 1;

            CollectRecursive(assemblyNode, partCounts, ref itemNumber);

            return partCounts.Values.OrderBy(r => r.ItemNumber).ToList();
        }

        private void CollectRecursive(ModelNode node, Dictionary<string, BomRow> partCounts, ref int itemNumber)
        {
            foreach (var child in node.Children)
            {
                string key = child.FilePath ?? child.FileName;

                if (partCounts.ContainsKey(key))
                {
                    // Increment quantity for duplicate parts
                    partCounts[key].Quantity++;
                }
                else
                {
                    var row = new BomRow
                    {
                        ItemNumber = itemNumber++,
                        PartNumber = child.PartNumber,
                        Description = child.Description,
                        Material = child.Material,
                        Quantity = 1,
                        FilePath = child.FilePath,
                        NodeType = child.NodeType
                    };

                    // Try to get weight from custom properties
                    if (child.CustomProperties.ContainsKey("Weight"))
                    {
                        double.TryParse(child.CustomProperties["Weight"], out double weight);
                        row.Weight = weight;
                    }

                    partCounts[key] = row;
                }

                // Recurse into sub-assemblies
                if (child.Children != null && child.Children.Count > 0)
                {
                    CollectRecursive(child, partCounts, ref itemNumber);
                }
            }
        }

        /// <summary>
        /// Write BOM to Excel with professional formatting.
        /// </summary>
        private void WriteExcel(List<BomRow> rows, ModelNode assemblyNode, string outputPath)
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add("BOM");

                // ── Title Row ─────────────────────────────────────
                ws.Cell("A1").Value = $"BILL OF MATERIALS — {assemblyNode.PartNumber}";
                ws.Range("A1:H1").Merge();
                ws.Cell("A1").Style
                    .Font.SetBold(true)
                    .Font.SetFontSize(14)
                    .Font.SetFontColor(XLColor.White)
                    .Fill.SetBackgroundColor(XLColor.FromHtml("#1B3A5C"))
                    .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                // ── Info Row ──────────────────────────────────────
                ws.Cell("A2").Value = $"Assembly: {assemblyNode.FileName}";
                ws.Cell("D2").Value = $"Rev: {assemblyNode.Revision}";
                ws.Cell("F2").Value = $"Date: {DateTime.Now:yyyy-MM-dd}";
                ws.Range("A2:H2").Style
                    .Font.SetItalic(true)
                    .Font.SetFontSize(9)
                    .Fill.SetBackgroundColor(XLColor.FromHtml("#E8EDF2"));

                // ── Headers ───────────────────────────────────────
                int headerRow = 4;
                string[] headers = { "Item #", "Part Number", "Description", "Material",
                                      "Qty", "Weight (kg)", "Total Wt (kg)", "Type" };

                for (int i = 0; i < headers.Length; i++)
                {
                    var cell = ws.Cell(headerRow, i + 1);
                    cell.Value = headers[i];
                    cell.Style
                        .Font.SetBold(true)
                        .Font.SetFontSize(10)
                        .Font.SetFontColor(XLColor.White)
                        .Fill.SetBackgroundColor(XLColor.FromHtml("#2C5F8A"))
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                        .Border.SetOutsideBorderColor(XLColor.White);
                }

                // ── Data Rows ─────────────────────────────────────
                double totalWeight = 0;
                int dataStartRow = headerRow + 1;

                for (int i = 0; i < rows.Count; i++)
                {
                    var row = rows[i];
                    int rowNum = dataStartRow + i;
                    bool isEven = i % 2 == 0;

                    ws.Cell(rowNum, 1).Value = row.ItemNumber;
                    ws.Cell(rowNum, 2).Value = row.PartNumber;
                    ws.Cell(rowNum, 3).Value = row.Description;
                    ws.Cell(rowNum, 4).Value = row.Material;
                    ws.Cell(rowNum, 5).Value = row.Quantity;
                    ws.Cell(rowNum, 6).Value = row.Weight;
                    ws.Cell(rowNum, 7).Value = row.Weight * row.Quantity;
                    ws.Cell(rowNum, 8).Value = row.NodeType.ToString();

                    totalWeight += row.Weight * row.Quantity;

                    // Alternating row colors
                    var bgColor = isEven ? XLColor.FromHtml("#F5F7FA") : XLColor.White;
                    ws.Range(rowNum, 1, rowNum, 8).Style
                        .Fill.SetBackgroundColor(bgColor)
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                        .Border.SetOutsideBorderColor(XLColor.FromHtml("#D0D5DD"))
                        .Font.SetFontSize(10);

                    // Center alignment for numeric columns
                    ws.Cell(rowNum, 1).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    ws.Cell(rowNum, 5).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                    ws.Cell(rowNum, 6).Style.NumberFormat.Format = "0.000";
                    ws.Cell(rowNum, 7).Style.NumberFormat.Format = "0.000";
                }

                // ── Totals Row ────────────────────────────────────
                int totalRow = dataStartRow + rows.Count;
                ws.Cell(totalRow, 1).Value = "";
                ws.Cell(totalRow, 4).Value = "TOTAL:";
                ws.Cell(totalRow, 5).Value = rows.Sum(r => r.Quantity);
                ws.Cell(totalRow, 7).Value = totalWeight;

                ws.Range(totalRow, 1, totalRow, 8).Style
                    .Font.SetBold(true)
                    .Font.SetFontSize(10)
                    .Fill.SetBackgroundColor(XLColor.FromHtml("#D4E6F1"))
                    .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

                ws.Cell(totalRow, 7).Style.NumberFormat.Format = "0.000";

                // ── Auto-fit columns ──────────────────────────────
                ws.Columns().AdjustToContents();

                // Set minimum column widths
                ws.Column(1).Width = 8;    // Item #
                ws.Column(2).Width = 18;   // Part Number
                ws.Column(3).Width = 30;   // Description
                ws.Column(4).Width = 18;   // Material
                ws.Column(5).Width = 6;    // Qty
                ws.Column(6).Width = 12;   // Weight
                ws.Column(7).Width = 14;   // Total Weight
                ws.Column(8).Width = 12;   // Type

                // ── Print settings ────────────────────────────────
                ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
                ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
                ws.PageSetup.FitToPages(1, 1);

                workbook.SaveAs(outputPath);
            }
        }

        /// <summary>
        /// BOM data row.
        /// </summary>
        private class BomRow
        {
            public int ItemNumber { get; set; }
            public string PartNumber { get; set; }
            public string Description { get; set; }
            public string Material { get; set; }
            public int Quantity { get; set; }
            public double Weight { get; set; }
            public string FilePath { get; set; }
            public ModelNodeType NodeType { get; set; }
        }
    }
}
