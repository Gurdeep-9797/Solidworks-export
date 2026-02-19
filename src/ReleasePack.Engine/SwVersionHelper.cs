using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace ReleasePack.Engine
{
    /// <summary>
    /// Central utility for SolidWorks version detection and cross-version helpers.
    /// Major API versions: 26=2018, 27=2019, 28=2020, 29=2021, 30=2022, 31=2023, 32=2024.
    /// </summary>
    public static class SwVersionHelper
    {
        // ───────────── Version detection ─────────────

        /// <summary>
        /// Parse the major API version from ISldWorks.RevisionNumber (format "XX.Y.Z").
        /// Returns 0 if parsing fails.
        /// </summary>
        public static int GetMajorVersion(ISldWorks swApp)
        {
            try
            {
                string rev = swApp.RevisionNumber();
                if (string.IsNullOrEmpty(rev)) return 0;
                string[] parts = rev.Split('.');
                if (parts.Length >= 1 && int.TryParse(parts[0], out int major))
                    return major;
            }
            catch { }
            return 0;
        }

        /// <summary>Check if the running SolidWorks is at least the given major API version.</summary>
        public static bool IsVersionAtLeast(ISldWorks swApp, int minMajor)
        {
            return GetMajorVersion(swApp) >= minMajor;
        }

        /// <summary>Map major API version to marketing year (26 → 2018, …).</summary>
        public static int GetSwYear(int majorVersion)
        {
            return majorVersion + 1992; // 26+1992=2018, 28+1992=2020, etc.
        }

        /// <summary>Get a human-readable version string, e.g. "SolidWorks 2020 (API 28)".</summary>
        public static string GetVersionString(ISldWorks swApp)
        {
            int major = GetMajorVersion(swApp);
            return major > 0 ? $"SolidWorks {GetSwYear(major)} (API {major})" : "SolidWorks (unknown version)";
        }

        // ───────────── Template discovery ─────────────

        /// <summary>
        /// Find the best available drawing template (.drwdot) by searching multiple sources:
        ///   1. User-specified path from ExportOptions
        ///   2. SW user preference (swDefaultTemplateDrawing)
        ///   3. SW file-locations template folder preference
        ///   4. ProgramData standard locations
        ///   5. Registry-based install-dir locations
        ///   6. Empty string fallback (SW uses its own default)
        /// </summary>
        public static string FindDrawingTemplate(ISldWorks swApp, ExportOptions options)
        {
            // 1. User-specified
            if (options != null &&
                !string.IsNullOrEmpty(options.DrawingTemplatePath) &&
                File.Exists(options.DrawingTemplatePath))
            {
                return options.DrawingTemplatePath;
            }

            // 2. SW default template preference
            try
            {
                string pref = swApp.GetUserPreferenceStringValue(
                    (int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
                if (!string.IsNullOrEmpty(pref) && File.Exists(pref))
                    return pref;
            }
            catch { }

            // 3. SW file-locations template folder
            try
            {
                string folders = swApp.GetUserPreferenceStringValue(
                    (int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates);
                if (!string.IsNullOrEmpty(folders))
                {
                    foreach (string folder in folders.Split(';'))
                    {
                        if (string.IsNullOrWhiteSpace(folder)) continue;
                        string found = SearchFolderForTemplate(folder.Trim());
                        if (found != null) return found;
                    }
                }
            }
            catch { }

            // 4. ProgramData standard locations (SW 2018-2026)
            for (int year = 2026; year >= 2018; year--)
            {
                string pdPath = Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData),
                    "SOLIDWORKS", $"SOLIDWORKS {year}", "templates");
                string found = SearchFolderForTemplate(pdPath);
                if (found != null) return found;
            }

            // 5. Registry-based install directory
            try
            {
                string installDir = GetInstallDirFromRegistry();
                if (!string.IsNullOrEmpty(installDir))
                {
                    string[] subPaths = new[]
                    {
                        @"lang\english\Tutorial\draw.drwdot",
                        @"lang\english\draw.drwdot",
                        @"adefault\draw.drwdot",
                    };
                    foreach (string sub in subPaths)
                    {
                        string full = Path.Combine(installDir, sub);
                        if (File.Exists(full)) return full;
                    }
                }
            }
            catch { }

            // 6. Try the executable directory
            try
            {
                string exeDir = Path.GetDirectoryName(swApp.GetExecutablePath());
                if (!string.IsNullOrEmpty(exeDir))
                {
                    string[] subPaths = new[]
                    {
                        @"lang\english\Tutorial\draw.drwdot",
                        @"lang\english\draw.drwdot",
                        @"adefault\draw.drwdot",
                    };
                    foreach (string sub in subPaths)
                    {
                        string full = Path.Combine(exeDir, sub);
                        if (File.Exists(full)) return full;
                    }
                }
            }
            catch { }

            return ""; // SW will use its own default
        }

        private static string SearchFolderForTemplate(string folder)
        {
            if (!Directory.Exists(folder)) return null;

            // Prefer "draw.drwdot", then any .drwdot
            string standard = Path.Combine(folder, "draw.drwdot");
            if (File.Exists(standard)) return standard;

            try
            {
                string[] templates = Directory.GetFiles(folder, "*.drwdot", SearchOption.TopDirectoryOnly);
                if (templates.Length > 0) return templates[0];
            }
            catch { }
            return null;
        }

        private static string GetInstallDirFromRegistry()
        {
            // Try each year from newest to oldest
            for (int year = 2026; year >= 2018; year--)
            {
                try
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(
                        $@"SOFTWARE\SolidWorks\SolidWorks {year}\Setup"))
                    {
                        if (key != null)
                        {
                            string loc = key.GetValue("SolidWorks Location") as string;
                            if (!string.IsNullOrEmpty(loc) && Directory.Exists(loc))
                                return loc;
                        }
                    }
                }
                catch { }
            }
            return null;
        }
    }
}
