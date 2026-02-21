using System;
using System.Collections.Generic;
using System.Linq;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Assigns hierarchical part numbers for assemblies.
    /// 
    /// Schema:
    ///   Top-level assembly: {ProjectPrefix}
    ///   Sub-assembly A:     A
    ///     Part 1:           A-1
    ///     Part 2:           A-2
    ///   Sub-assembly B:     B
    ///     Part 1:           B-1
    ///     Sub-assy B.1:     B.1
    ///       Part 1:         B.1-1
    ///   Loose part:         1, 2, 3...
    /// 
    /// This numbering helps in:
    ///   - BOM table item labeling
    ///   - Balloon annotation
    ///   - Sub-assembly/main-assembly cross-referencing
    ///   - Drawing sheet organization
    /// </summary>
    public class PartNumberingEngine
    {
        /// <summary>
        /// Numbered node wrapping a ModelNode with its assigned item number.
        /// </summary>
        public class NumberedNode
        {
            public ModelNode Original { get; set; }
            public string ItemNumber { get; set; }      // e.g., "A-1", "B.1-2"
            public string DisplayName { get; set; }     // e.g., "A-1 Bracket"
            public int Level { get; set; }              // 0 = top assembly, 1 = sub-assy, 2 = part
            public List<NumberedNode> Children { get; set; } = new List<NumberedNode>();
        }

        /// <summary>
        /// Assign hierarchical part numbers to the entire assembly tree.
        /// </summary>
        /// <param name="root">Root assembly node.</param>
        /// <param name="projectPrefix">Optional project prefix (e.g., "PRJ-001").</param>
        /// <returns>Numbered tree with item numbers assigned.</returns>
        public static NumberedNode AssignNumbers(ModelNode root, string projectPrefix = "")
        {
            var numbered = new NumberedNode
            {
                Original = root,
                ItemNumber = string.IsNullOrEmpty(projectPrefix) ? "ASSY" : projectPrefix,
                DisplayName = root.FileName,
                Level = 0
            };

            AssignChildNumbers(root, numbered, "");

            return numbered;
        }

        private static void AssignChildNumbers(ModelNode source, NumberedNode parent, string prefix)
        {
            char subAssyLetter = 'A';
            int partIndex = 1;

            foreach (var child in source.Children)
            {
                var numberedChild = new NumberedNode
                {
                    Original = child,
                    Level = parent.Level + 1
                };

                if (child.NodeType == ModelNodeType.Assembly)
                {
                    // Sub-assemblies get letters
                    string letter = subAssyLetter.ToString();
                    subAssyLetter++;

                    string childPrefix = string.IsNullOrEmpty(prefix)
                        ? letter
                        : $"{prefix}.{letter}";

                    numberedChild.ItemNumber = childPrefix;
                    numberedChild.DisplayName = $"{childPrefix} {child.FileName}";

                    // Recurse into sub-assembly
                    AssignChildNumbers(child, numberedChild, childPrefix);
                }
                else
                {
                    // Parts get numbers
                    string partNum = string.IsNullOrEmpty(prefix)
                        ? partIndex.ToString()
                        : $"{prefix}-{partIndex}";

                    numberedChild.ItemNumber = partNum;
                    numberedChild.DisplayName = $"{partNum} {child.FileName}";
                    partIndex++;
                }

                parent.Children.Add(numberedChild);
            }
        }

        /// <summary>
        /// Flatten the numbered tree into a sorted list for BOM table.
        /// </summary>
        public static List<NumberedNode> FlattenForBOM(NumberedNode root, bool includeSubAssemblies = true)
        {
            var flat = new List<NumberedNode>();
            FlattenRecursive(root, flat, includeSubAssemblies);
            return flat.OrderBy(n => n.ItemNumber, new ItemNumberComparer()).ToList();
        }

        private static void FlattenRecursive(NumberedNode node, List<NumberedNode> flat, bool includeSubAssemblies)
        {
            // Skip the root assembly itself
            if (node.Level > 0)
            {
                if (node.Original.NodeType == ModelNodeType.Assembly && !includeSubAssemblies)
                {
                    // Skip sub-assembly entries, but include their children
                }
                else
                {
                    flat.Add(node);
                }
            }

            foreach (var child in node.Children)
            {
                FlattenRecursive(child, flat, includeSubAssemblies);
            }
        }

        /// <summary>
        /// Custom comparer for item numbers to sort correctly:
        /// A before B, A-1 before A-2, A-10 after A-9
        /// </summary>
        private class ItemNumberComparer : IComparer<string>
        {
            public int Compare(string x, string y)
            {
                if (x == null && y == null) return 0;
                if (x == null) return -1;
                if (y == null) return 1;

                var partsX = SplitItemNumber(x);
                var partsY = SplitItemNumber(y);

                int minLen = Math.Min(partsX.Length, partsY.Length);
                for (int i = 0; i < minLen; i++)
                {
                    int cmp;
                    if (int.TryParse(partsX[i], out int nx) && int.TryParse(partsY[i], out int ny))
                        cmp = nx.CompareTo(ny);
                    else
                        cmp = string.Compare(partsX[i], partsY[i], StringComparison.OrdinalIgnoreCase);

                    if (cmp != 0) return cmp;
                }

                return partsX.Length.CompareTo(partsY.Length);
            }

            private string[] SplitItemNumber(string s)
            {
                // Split by . and - to get sortable components
                return s.Split(new[] { '.', '-' }, StringSplitOptions.RemoveEmptyEntries);
            }
        }

        /// <summary>
        /// Generate a summary string showing the numbering tree.
        /// </summary>
        public static string PrintTree(NumberedNode root, int indent = 0)
        {
            var lines = new List<string>();
            PrintTreeRecursive(root, lines, indent);
            return string.Join(Environment.NewLine, lines);
        }

        private static void PrintTreeRecursive(NumberedNode node, List<string> lines, int indent)
        {
            string prefix = new string(' ', indent * 2);
            string type = node.Original.NodeType == ModelNodeType.Assembly ? "[ASSY]" : "[PART]";
            lines.Add($"{prefix}{node.ItemNumber} {type} {node.Original.FileName}");

            foreach (var child in node.Children)
            {
                PrintTreeRecursive(child, lines, indent + 1);
            }
        }
    }
}
