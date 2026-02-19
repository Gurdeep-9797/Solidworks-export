using System;
using System.IO;
using Microsoft.Win32;

namespace ReleasePack.Installer
{
    public static class SolidWorksDetector
    {
        public static string FindSolidWorksExecutable()
        {
            // Strategy 1: Check COM Registration (Most reliable)
            string comPath = GetPathFromCOM();
            if (!string.IsNullOrEmpty(comPath) && File.Exists(comPath)) return comPath;

            // Strategy 2: Check Registry (HKLM and HKCU)
            for (int year = 2026; year >= 2018; year--)
            {
                string path = GetInstallDirFromRegistry(year);
                if (!string.IsNullOrEmpty(path))
                {
                    string exePath = Path.Combine(path, "SLDWORKS.exe");
                    if (File.Exists(exePath)) return exePath;
                }
            }

            // Strategy 3: Check Standard Paths
            string[] standardPaths = new[]
            {
                @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SLDWORKS.exe",
                @"C:\Program Files\SolidWorks\SLDWORKS.exe"
            };

            foreach (var path in standardPaths)
            {
                if (File.Exists(path)) return path;
            }

            return null;
        }

        private static string GetPathFromCOM()
        {
            try
            {
                // Look up CLSID for SolidWorks.Application
                using (var root = Registry.ClassesRoot.OpenSubKey("SolidWorks.Application"))
                {
                    if (root == null) return null;
                    using (var clsidKey = root.OpenSubKey("CLSID"))
                    {
                        string clsid = clsidKey?.GetValue("") as string;
                        if (string.IsNullOrEmpty(clsid)) return null;

                        // Look up LocalServer32 for that CLSID
                        using (var serverKey = Registry.ClassesRoot.OpenSubKey($@"CLSID\{clsid}\LocalServer32"))
                        {
                            string path = serverKey?.GetValue("") as string;
                            if (string.IsNullOrEmpty(path)) return null;

                            // Cleanup quotes/args if present
                            path = path.Trim('"');
                            if (path.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                                return path;
                        }
                    }
                }
            }
            catch { }
            return null;
        }

        private static string GetInstallDirFromRegistry(int year)
        {
            string[] roots = { "HKEY_LOCAL_MACHINE", "HKEY_CURRENT_USER" };
            string subKey = $@"SOFTWARE\SolidWorks\SolidWorks {year}\Setup";

            foreach (var root in roots)
            {
                try
                {
                    string loc = Registry.GetValue($@"{root}\{subKey}", "SolidWorks Location", null) as string;
                    if (!string.IsNullOrEmpty(loc) && Directory.Exists(loc))
                        return loc;
                }
                catch { }
            }
            return null;
        }
    }
}
