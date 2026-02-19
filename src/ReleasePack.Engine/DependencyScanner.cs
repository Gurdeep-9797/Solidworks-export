using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Scans a SolidWorks model's dependency tree and builds a worktree of ModelNodes.
    /// Handles assemblies by recursively traversing all components.
    /// </summary>
    public class DependencyScanner
    {
        private readonly ISldWorks _swApp;
        private readonly HashSet<string> _visitedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly IProgressCallback _progress;

        public DependencyScanner(ISldWorks swApp, IProgressCallback progress = null)
        {
            _swApp = swApp;
            _progress = progress;
        }

        /// <summary>
        /// Scans the active document based on the selected scope.
        /// </summary>
        public List<ModelNode> Scan(ExportOptions options)
        {
            _visitedPaths.Clear();
            var results = new List<ModelNode>();

            ModelDoc2 rootDoc = null;
            bool openedRemote = false;

            switch (options.Scope)
            {
                case ExportScope.CurrentDocument:
                case ExportScope.CurrentAndChildren:
                    rootDoc = (ModelDoc2)_swApp.ActiveDoc;
                    break;

                case ExportScope.RemoteFile:
                    if (string.IsNullOrEmpty(options.RemoteFilePath) || !File.Exists(options.RemoteFilePath))
                        throw new FileNotFoundException("Remote file not found: " + options.RemoteFilePath);

                    int errors = 0, warnings = 0;
                    rootDoc = (ModelDoc2)_swApp.OpenDoc6(
                        options.RemoteFilePath,
                        GetDocType(options.RemoteFilePath),
                        (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
                        "", ref errors, ref warnings);

                    if (rootDoc == null)
                        throw new InvalidOperationException($"Failed to open '{options.RemoteFilePath}'. Error code: {errors}");

                    openedRemote = true;
                    break;
            }

            if (rootDoc == null)
                throw new InvalidOperationException("No active document found in SolidWorks.");

            var rootNode = BuildNode(rootDoc);

            if (options.Scope == ExportScope.CurrentAndChildren || options.Scope == ExportScope.RemoteFile)
            {
                // If it's an assembly, recurse into children
                if (rootDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    _progress?.LogMessage("Scanning assembly components...");
                    ScanAssemblyChildren(rootDoc, rootNode);
                }
            }

            results.Add(rootNode);
            return results;
        }

        /// <summary>
        /// Recursively scan all components within an assembly.
        /// </summary>
        private void ScanAssemblyChildren(ModelDoc2 assemblyDoc, ModelNode parentNode)
        {
            AssemblyDoc assy = (AssemblyDoc)assemblyDoc;
            object[] components = (object[])assy.GetComponents(false); // top-level only

            if (components == null) return;

            foreach (object comp in components)
            {
                Component2 swComp = (Component2)comp;

                try
                {
                    // Skip suppressed components
                    int suppression = swComp.GetSuppression2();
                    if (suppression == (int)swComponentSuppressionState_e.swComponentSuppressed)
                    {
                        _progress?.LogWarning($"Skipping suppressed component: {swComp.Name2}");
                        continue;
                    }

                    // Try to resolve lightweight components
                    if (suppression == (int)swComponentSuppressionState_e.swComponentLightweight)
                    {
                        _progress?.LogMessage($"Resolving lightweight component: {swComp.Name2}");
                        swComp.SetSuppression2((int)swComponentSuppressionState_e.swComponentFullyResolved);
                    }

                    ModelDoc2 compDoc = (ModelDoc2)swComp.GetModelDoc2();
                    if (compDoc == null)
                    {
                        _progress?.LogWarning($"Could not get model for component: {swComp.Name2}");
                        continue;
                    }

                    string compPath = compDoc.GetPathName();

                    // Avoid circular references
                    if (!string.IsNullOrEmpty(compPath) && _visitedPaths.Contains(compPath))
                        continue;

                    var childNode = BuildNode(compDoc);
                    parentNode.Children.Add(childNode);

                    // Recurse if this child is also an assembly
                    if (compDoc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
                    {
                        ScanAssemblyChildren(compDoc, childNode);
                    }

                    _progress?.LogMessage($"Found: {childNode.FileName} ({childNode.NodeType})");
                }
                catch (Exception ex)
                {
                    _progress?.LogWarning($"Error scanning component {swComp.Name2}: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Build a ModelNode from a ModelDoc2 instance.
        /// </summary>
        private ModelNode BuildNode(ModelDoc2 doc)
        {
            string path = doc.GetPathName();
            if (!string.IsNullOrEmpty(path))
                _visitedPaths.Add(path);

            var node = new ModelNode
            {
                FilePath = path,
                FileName = Path.GetFileName(path),
                ConfigurationName = doc.ConfigurationManager?.ActiveConfiguration?.Name ?? "",
                NodeType = GetNodeType(doc),
                IsSheetMetal = DetectSheetMetal(doc)
            };

            // Extract custom properties
            ExtractCustomProperties(doc, node);

            return node;
        }

        /// <summary>
        /// Extract custom properties from the active configuration and the file.
        /// </summary>
        private void ExtractCustomProperties(ModelDoc2 doc, ModelNode node)
        {
            try
            {
                // File-level custom properties
                CustomPropertyManager filePropMgr = (CustomPropertyManager)doc.Extension.CustomPropertyManager[""];
                if (filePropMgr != null)
                {
                    ReadProperties(filePropMgr, node);
                }

                // Config-level custom properties (override file-level)
                if (!string.IsNullOrEmpty(node.ConfigurationName))
                {
                    CustomPropertyManager configPropMgr =
                        (CustomPropertyManager)doc.Extension.CustomPropertyManager[node.ConfigurationName];
                    if (configPropMgr != null)
                    {
                        ReadProperties(configPropMgr, node);
                    }
                }
            }
            catch (Exception ex)
            {
                _progress?.LogWarning($"Could not read custom properties for {node.FileName}: {ex.Message}");
            }
        }

        private void ReadProperties(CustomPropertyManager propMgr, ModelNode node)
        {
            object propNames = null;
            object propTypes = null;
            object propValues = null;
            object propResolved = null;
            object propLinks = null;

            int count = propMgr.GetAll3(ref propNames, ref propTypes, ref propValues, ref propResolved, ref propLinks);

            if (count > 0 && propNames is string[] names && propResolved is string[] resolved)
            {
                for (int i = 0; i < names.Length; i++)
                {
                    if (!string.IsNullOrEmpty(names[i]))
                    {
                        node.CustomProperties[names[i]] = (i < resolved.Length) ? resolved[i] : "";
                    }
                }
            }
        }

        /// <summary>
        /// Detect if a part is sheet metal by looking for sheet metal features.
        /// </summary>
        private bool DetectSheetMetal(ModelDoc2 doc)
        {
            if (doc.GetType() != (int)swDocumentTypes_e.swDocPART) return false;

            try
            {
                Feature feat = (Feature)doc.FirstFeature();
                while (feat != null)
                {
                    string typeName = feat.GetTypeName2();
                    if (typeName == "SheetMetal" || typeName == "SMBaseFlange" ||
                        typeName == "FlatPattern" || typeName == "EdgeFlange" ||
                        typeName == "SketchBend" || typeName == "Hem" ||
                        typeName == "Jog" || typeName == "BreakCorner")
                    {
                        return true;
                    }
                    feat = (Feature)feat.GetNextFeature();
                }
            }
            catch { }

            return false;
        }

        private ModelNodeType GetNodeType(ModelDoc2 doc)
        {
            switch (doc.GetType())
            {
                case (int)swDocumentTypes_e.swDocPART: return ModelNodeType.Part;
                case (int)swDocumentTypes_e.swDocASSEMBLY: return ModelNodeType.Assembly;
                case (int)swDocumentTypes_e.swDocDRAWING: return ModelNodeType.Drawing;
                default: return ModelNodeType.Part;
            }
        }

        private int GetDocType(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            switch (ext)
            {
                case ".sldprt": return (int)swDocumentTypes_e.swDocPART;
                case ".sldasm": return (int)swDocumentTypes_e.swDocASSEMBLY;
                case ".slddrw": return (int)swDocumentTypes_e.swDocDRAWING;
                default: return (int)swDocumentTypes_e.swDocPART;
            }
        }

        /// <summary>
        /// Flatten the worktree into a distinct list of unique model nodes (by file path).
        /// </summary>
        public static List<ModelNode> Flatten(List<ModelNode> roots)
        {
            var result = new List<ModelNode>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            FlattenRecursive(roots, result, seen);
            return result;
        }

        private static void FlattenRecursive(List<ModelNode> nodes, List<ModelNode> result, HashSet<string> seen)
        {
            foreach (var node in nodes)
            {
                if (!string.IsNullOrEmpty(node.FilePath) && !seen.Contains(node.FilePath))
                {
                    seen.Add(node.FilePath);
                    result.Add(node);
                }
                if (node.Children != null && node.Children.Count > 0)
                {
                    FlattenRecursive(node.Children, result, seen);
                }
            }
        }
    }
}
