# Example script for using the DskMonPS7 DLL.
# This script can be run in PowerShell 7+.

# Get the directory of the script.
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# The DLLs are expected to be in a 'dist' directory relative to the script.
$DllPath = Join-Path $ScriptDir "..\dist\DskMon_ps7_x64.dll" # Assuming x64

# Load the assembly
try {
    [System.Reflection.Assembly]::LoadFrom($DllPath) | Out-Null
} catch {
    Write-Error "Failed to load the DskMon DLL from '$DllPath'. Please make sure the DLL exists at this path."
    return
}

Write-Host "Waiting for disk events... Press Ctrl+C to exit."

while ($true) {
    # GetNextEvent() blocks until a disk is attached or detached.
    $DskEvent = [DskMon.DskMon]::GetNextEvent()

    if ($DskEvent) {
        # The event object is a PSCustomObject, ready to use.
        $DskEvent | Format-List
    }
}
