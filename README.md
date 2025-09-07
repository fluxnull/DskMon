# DskMon

**DskMon.dll** is a .NET library for Windows that monitors physical disk attach and detach events, delivering a clean, PowerShell-friendly `PSCustomObject` with trimmed properties for logging, routing, or ingest workflows. It uses WMI (`Win32_DiskDrive`) for event detection and enriches data with CIM (`MSFT_Disk`) and fallback queries for reliability across diverse hardware.

Designed for simplicity and determinism, it replaces JSON-emitting console apps with a lightweight DLL callable from PowerShell 5.1 or 7.x. Ideal for kiosk systems, automated backups, or disk monitoring workflows.

## Features

- **Event-driven:** Captures disk attach/detach with low CPU usage (WMI `WITHIN 2` seconds polling).
- **Rich output:** Returns a `PSCustomObject` with fields like `diskNumber`, `serialNumber`, `model`, `mountPoint`, `sizeInBytes`, and more.
- **Robust fallbacks:** Sources data from CIM (`MSFT_Disk`), WMI (`Win32_*`), and PnP parsing for maximum compatibility.
- **PowerShell-native:** Works seamlessly in PowerShell 5.1 (.NET Framework 4.8) and 7.x (.NET 8.0).
- **Trimmed strings:** All string fields are `.Trim()`ed to eliminate vendor padding or whitespace.

## Prerequisites

- **OS:** Windows 7 SP1 or later (Windows 10/11 recommended for full CIM support).
- **PowerShell:**
  - For `DskMonPS5` (.NET Framework 4.8): PowerShell 5.1 (included with Windows).
  - For `DskMonPS7` (.NET 8.0): PowerShell 7.4 or later.
- **Permissions:** Standard user context suffices for most fields; elevated rights may be needed for some `SerialNumber` queries.
- **Build Environment:** A system with .NET SDK 8.0 and Mono for building both targets (see Jules build instructions below).

## Installation

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/fluxnull/DskMon
   cd DskMon
   ```

2. **Build the DLLs:**
   Follow the instructions in the [Building for Jules](#building-for-jules) section to generate `DskMon32.dll` and `DskMon64.dll` for both PowerShell 5.1 and 7.x.

3. **Copy DLLs:**
   Place the compiled `DskMon.dll` (from either `DskMonPS5` or `DskMonPS7` build) into your project directory or a known path.

## Usage

### PowerShell Example

```powershell
# Load the DLL
Add-Type -Path ".\DskMon.dll"

# Monitor disk events in a loop
while ($true) {
    $event = [DskMon.DskMon]::GetNextEvent()  # Blocks until an event; use timeoutMs for non-blocking
    if ($event) {
        $event | Format-Table  # Display event details
        # Example: Log to file or process
        # $event | ConvertTo-Json | Out-File "disk_events.log" -Append
    }
}
```

### Output Example

For a USB disk attach event:

```powershell
@{
    eventType        = 'Attached'
    diskNumber       = 2
    name             = '\\.\PHYSICALDRIVE2'
    serialNumber     = 'WX51A28CL8R0'
    model            = 'G-DRIVE mobile USB Device'
    firmwareVersion  = '1019'
    manufacturer     = 'G-DRIVE'
    pnpDeviceID      = 'USBSTOR\DISK&VEN_G-DRIVE&PROD_MOBILE&REV_1019\57583531413238434C385230&0'
    friendlyName     = 'G-DRIVE mobile USB Device'
    caption          = 'G-DRIVE mobile USB Device'
    hardwareID       = 'USBSTOR\DiskG-DRIVE_Mobile____1019'
    mountPoint       = 'D:'
    volumeName       = 'New Volume'
    interfaceType    = 'USB'
    fileSystem       = 'NTFS'
    sizeInBytes      = 1000151707648
    freeSpaceInBytes = 998598656000
}
```

## Building for Jules

The following instructions are tailored for the Jules AI agent's build environment (Ubuntu 24.04, as specified in `Build_20250903-213955.txt`). This environment includes .NET SDK 8.0, Mono, and NuGet for building both .NET Framework 4.8 and .NET 8.0 targets.

### Prerequisites (Jules Environment)

Ensure the Jules build environment is set up as per `Build_20250903-213955.txt`. Key components:
- **.NET SDK 8.0:** For building `DskMonPS7` (.NET 8.0).
- **Mono:** For building `DskMonPS5` (.NET Framework 4.8).
- **NuGet CLI:** For package restoration.
- **PATH:** Includes `/usr/share/dotnet`, `/usr/bin/mono`, `/opt/nuget`, and `/opt/microsoft/powershell/7`.

### Build Instructions

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/fluxnull/DskMon
   cd DskMon
   ```

2. **Build DskMonPS5 (.NET Framework 4.8 for PowerShell 5.1):**
   ```bash
   # Restore packages using NuGet (Mono)
   mono /opt/nuget/nuget.exe restore DskMonPS5.csproj

   # Build for 32-bit
   msbuild DskMonPS5.csproj -p:Configuration=Release -p:TargetFramework=net48 -p:Platform=x86 -p:OutputPath=bin/Release/net48/x86
   mv bin/Release/net48/x86/DskMon.dll bin/Release/net48/x86/DskMon32.dll

   # Build for 64-bit
   msbuild DskMonPS5.csproj -p:Configuration=Release -p:TargetFramework=net48 -p:Platform=x64 -p:OutputPath=bin/Release/net48/x64
   mv bin/Release/net48/x64/DskMon.dll bin/Release/net48/x64/DskMon64.dll
   ```

   **Output:**
   - `bin/Release/net48/x86/DskMon32.dll` (for PowerShell 5.1, 32-bit)
   - `bin/Release/net48/x64/DskMon64.dll` (for PowerShell 5.1, 64-bit)

3. **Build DskMonPS7 (.NET 8.0 for PowerShell 7.x):**
   ```bash
   # Restore and build for 32-bit
   dotnet build DskMonPS7.csproj -c Release -f net8.0 -p:Platform=x86 -o bin/Release/net8.0/x86
   mv bin/Release/net8.0/x86/DskMon.dll bin/Release/net8.0/x86/DskMon32.dll

   # Restore and build for 64-bit
   dotnet build DskMonPS7.csproj -c Release -f net8.0 -p:Platform=x64 -o bin/Release/net8.0/x64
   mv bin/Release/net8.0/x64/DskMon.dll bin/Release/net8.0/x64/DskMon64.dll
   ```

   **Output:**
   - `bin/Release/net8.0/x86/DskMon32.dll` (for PowerShell 7.x, 32-bit)
   - `bin/Release/net8.0/x64/DskMon64.dll` (for PowerShell 7.x, 64-bit)

4. **Verify Builds:**
   ```bash
   # Check .NET Framework builds
   ls bin/Release/net48/x86/DskMon32.dll
   ls bin/Release/net48/x64/DskMon64.dll

   # Check .NET 8.0 builds
   ls bin/Release/net8.0/x86/DskMon32.dll
   ls bin/Release/net8.0/x64/DskMon64.dll
   ```

5. **Test the DLLs (Optional):**
   On a Windows system with PowerShell 5.1 or 7.x:
   ```powershell
   # Test PowerShell 5.1 DLL
   Add-Type -Path ".\bin\Release\net48\x64\DskMon64.dll"
   [DskMon.DskMon]::GetNextEvent(5000)  # Test with 5s timeout

   # Test PowerShell 7.x DLL
   Add-Type -Path ".\bin\Release\net8.0\x64\DskMon64.dll"
   [DskMon.DskMon]::GetNextEvent(5000)  # Test with 5s timeout
   ```

## Testing

### Functional Tests
- **Attach/Detach:** Test with SATA, USB, NVMe, and external drives.
- **File Systems:** Verify with NTFS, exFAT, FAT32, and unformatted disks.
- **PnP Variants:** Test with different USB chipsets (ASMedia, JMicron, etc.).

### Reliability Tests
- **Timeout:** Call `GetNextEvent(1000)` and confirm `null` return.
- **Mount Delay:** Attach a disk and verify drive letter capture within 4s.
- **Permissions:** Test under standard and elevated accounts to ensure fallback behavior.

## Limitations

- **Single Mount:** Returns the first mount point found; multi-partition disks may need future array support.
- **Flaky Hardware:** Some USB bridges may obscure serial numbers or firmware; PnP parsing is best-effort.
- **AutoPlay:** Does not modify system AutoPlay settings; use external scripts for kiosk workflows.

## Contributing

- **Issues:** Report bugs or request features via GitHub Issues.
- **Pull Requests:** Ensure builds pass for both `DskMonPS5` and `DskMonPS7` projects.
- **Code Style:** Follow C# conventions; ensure all strings are trimmed in output.

## License

MIT License. See `LICENSE` file for details.

## Acknowledgments

- Built for compatibility with Jules AI agent's build environment.
- Leverages WMI and CIM for robust disk monitoring.
- Inspired by the need for a lightweight, PowerShell-native disk event solution.
