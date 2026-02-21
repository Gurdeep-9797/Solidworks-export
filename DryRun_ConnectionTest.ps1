$ErrorActionPreference = "Stop"

Write-Output "=========================================="
Write-Output "SolidWorks V2 Subsystem: Dry-Run COM Test"
Write-Output "=========================================="
Write-Output "Attempting native COM connection via PowerShell..."

try {
    # Try attaching to an active SolidWorks instance
    $swApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("SldWorks.Application")
    Write-Output "[SUCCESS] Connected to active SolidWorks instance."
    
    # Get active Document
    $doc = $swApp.IActiveDoc2
    if ($doc) {
        Write-Output "[INFO] Active Document: $($doc.GetPathName())"
    }
    else {
        Write-Output "[WARNING] Connected to SolidWorks, but no document is currently open."
    }

}
catch {
    # Fallback to creating a new background instance
    Write-Output "[INFO] No active instance found. Backgrounding a new SolidWorks instance..."
    try {
        $swApp = New-Object -ComObject SldWorks.Application
        Write-Output "[SUCCESS] Instantiated new SolidWorks invisible process."
        
        $swApp.ExitApp()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($swApp) | Out-Null
        Write-Output "[INFO] Closed background SolidWorks instance."
    }
    catch {
        Write-Output "[FATAL ERROR] Could not connect to or instantiate SolidWorks COM."
        Write-Output $_.Exception.Message
        Exit 1
    }
}

Write-Output "=========================================="
Write-Output "Dry-Run Connection Test COMPLETE."
Write-Output "=========================================="
