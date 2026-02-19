using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace ReleasePack.Installer
{
    public static class RegistrationHelper
    {
        public static string RegisterAddIn(string dllPath)
        {
            try
            {
                if (!File.Exists(dllPath))
                    return "Error: DLL not found.";

                // Load the assembly
                // Validating 64-bit execution for correct registry writing
                if (System.IntPtr.Size == 4)
                {
                    return "Error: Installer is running in 32-bit mode. SolidWorks requires 64-bit registration.";
                }

                System.Reflection.Assembly asm = System.Reflection.Assembly.LoadFrom(dllPath);
                var reg = new RegistrationServices();
                
                // Register: SetCodeBase is critical for finding the DLL from SW
                bool result = reg.RegisterAssembly(asm, AssemblyRegistrationFlags.SetCodeBase);
                
                if (result)
                    return "Success";
                else
                    return "Registration returned false (no types found to register?).";
            }
            catch (Exception ex)
            {
                // Detailed error info
                return $"Error: {ex.Message}\nStack: {ex.StackTrace}";
            }
        }
    }
}
