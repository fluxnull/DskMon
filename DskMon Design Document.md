# DskMon.dll — Design Document

*Last updated: 2025-09-03 (America/New\_York)*

## 1) What this is (and why)

**DskMon.dll** is a small .NET library that watches Windows for **physical disk attach/detach** and returns a **PowerShell-native** object with clean, trimmed fields ready for logging, routing, and copy/ingest workflows.

* **Primary goal:** replace a console EXE that emitted JSON with a **DLL** that your PowerShell loop can call directly.
* **Output:** a **`PSCustomObject`** (actually a `PSObject` with `NoteProperty`s) whose **string fields are `.Trim()`ed**.
* **Reliability:** uses **CIM (MSFT\_Disk)** where available, then **WMI** as fallback for every field that matters.

If you grew up on SCSI IDs and IRQ jumpers, you get the vibe: **simple, deterministic, zero fluff.**

---

## 2) High-level architecture

```
PowerShell loop
     │
     ▼
[DskMon.DskMon]::GetNextEvent(timeoutMs)
     │   ├─ Creates 2 WMI event watchers (creation/deletion) on Win32_DiskDrive
     │   ├─ Blocks until one fires (or returns null on timeout)
     │   └─ Rebinds target instance to Win32_DiskDrive and builds output object
     │        ├─ Pulls MSFT_Disk (CIM) by Disk Number (if present)
     │        ├─ Field-by-field fallback (CIM→WMI→PnP parse)
     │        └─ Maps Disk→Partition→LogicalDisk→Volume for mount/fs/free
     ▼
PSCustomObject (trimmed properties)
```

**Eventing model:** WMI intrinsic events

* `__InstanceCreationEvent` of `Win32_DiskDrive` ⇒ `"Attached"`
* `__InstanceDeletionEvent` of `Win32_DiskDrive` ⇒ `"Detached"`

**Polling interval:** `WITHIN 2` seconds (low CPU, good enough responsiveness).
**Mount timing:** a short post-attach poll (≤4s) to catch the drive letter assignment.

---

## 3) Public surface (how you call it)

### 3.1 Static API

```csharp
public static class DskMon
{
    /// Blocks until an attach/detach occurs (or timeout).
    /// Returns a PSCustomObject (trimmed strings) or null on timeout.
    public static System.Management.Automation.PSObject GetNextEvent(int timeoutMilliseconds = Timeout.Infinite);
}
```

### 3.2 PowerShell usage

```powershell
Add-Type -Path .\DskMon.dll

while ($true) {
  $evt = [DskMon.DskMon]::GetNextEvent()  # blocks; null only on timeout if you pass one
  if ($evt) {
    $evt  # already trimmed; ready to log/route/copy
  }
}
```

> **Type name note:** The compiled type is `DskMon.DskMon`.
> If you want to call `[DskMon]::GetNextEvent()` exactly, compile without a namespace or add a type alias in your environment.

---

## 4) Data model (what you get back)

### 4.1 Property set (camelCase)

| Property           | Type     | Notes                                 |
| ------------------ | -------- | ------------------------------------- |
| `eventType`        | `string` | `"Attached"` or `"Detached"`          |
| `diskNumber`       | `int`    | Windows Disk #                        |
| `name`             | `string` | `\\.\PHYSICALDRIVE#`                  |
| `serialNumber`     | `string` | Trimmed; multiple fallback paths      |
| `model`            | `string` | Device model/caption                  |
| `firmwareVersion`  | `string` | Firmware/Revision                     |
| `manufacturer`     | `string` | From WMI or parsed from PnP           |
| `pnpDeviceID`      | `string` | Raw PnP ID                            |
| `friendlyName`     | `string` | CIM/WMI; falls back to model          |
| `caption`          | `string` | WMI caption                           |
| `hardwareID`       | `string` | PnP HardwareID\[0]                    |
| `mountPoint`       | `string` | Example: `D:` (blank on early detach) |
| `volumeName`       | `string` | Volume label                          |
| `interfaceType`    | `string` | USB/SATA/NVMe/etc.                    |
| `fileSystem`       | `string` | NTFS/exFAT/etc.                       |
| `sizeInBytes`      | `long`   | Physical disk size                    |
| `freeSpaceInBytes` | `long`   | Volume free space (best effort)       |

**All string fields are `Trim()`ed** before being added to the PS object.

### 4.2 Example output (typical attach)

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

---

## 5) Field sourcing & fallback logic

### 5.1 Source classes

* **CIM (preferred):** `root\microsoft\windows\storage: MSFT_Disk`
* **WMI (classic):** `root\CIMV2: Win32_DiskDrive`, `Win32_DiskPartition`, `Win32_LogicalDisk`, `Win32_Volume`, `Win32_PnPEntity`, `Win32_PhysicalMedia`

### 5.2 Fallback matrix

| Field              | Primary                         | Fallback #1                                 | Fallback #2                                         | Fallback #3                   |
| ------------------ | ------------------------------- | ------------------------------------------- | --------------------------------------------------- | ----------------------------- |
| `diskNumber`       | `Win32_DiskDrive.Index`         | `MSFT_Disk.Number`                          | —                                                   | —                             |
| `name`             | `Win32_DiskDrive.DeviceID`      | —                                           | —                                                   | —                             |
| `serialNumber`     | `MSFT_Disk.SerialNumber`        | `Win32_DiskDrive.SerialNumber`              | `Win32_PhysicalMedia.SerialNumber` (Tag = DeviceID) | parse from `PNPDeviceID` tail |
| `model`            | `MSFT_Disk.Model`               | `Win32_DiskDrive.Model`                     | `Win32_DiskDrive.Caption`                           | —                             |
| `firmwareVersion`  | `MSFT_Disk.FirmwareVersion`     | `Win32_DiskDrive.FirmwareRevision`          | —                                                   | —                             |
| `manufacturer`     | `Win32_DiskDrive.Manufacturer`  | parsed from `PNPDeviceID (&VEN_)`           | `MSFT_Disk.FriendlyName`                            | —                             |
| `pnpDeviceID`      | `Win32_DiskDrive.PNPDeviceID`   | `Win32_PnPEntity.PNPDeviceID` (by DeviceID) | —                                                   | —                             |
| `friendlyName`     | `MSFT_Disk.FriendlyName`        | `Win32_DiskDrive.FriendlyName`              | `model`                                             | —                             |
| `caption`          | `Win32_DiskDrive.Caption`       | `model`                                     | —                                                   | —                             |
| `hardwareID`       | `Win32_PnPEntity.HardwareID[0]` | —                                           | —                                                   | —                             |
| `mountPoint`       | `Win32_Volume.DriveLetter`      | `Win32_LogicalDisk.DeviceID`                | —                                                   | —                             |
| `volumeName`       | `Win32_Volume.Label`            | `Win32_LogicalDisk.VolumeName`              | —                                                   | —                             |
| `interfaceType`    | `MSFT_Disk.BusType → string`    | `Win32_DiskDrive.InterfaceType`             | —                                                   | —                             |
| `fileSystem`       | `Win32_Volume.FileSystem`       | `Win32_LogicalDisk.FileSystem`              | —                                                   | —                             |
| `sizeInBytes`      | `MSFT_Disk.Size`                | `Win32_DiskDrive.Size`                      | —                                                   | —                             |
| `freeSpaceInBytes` | `Win32_Volume.FreeSpace`        | `Win32_LogicalDisk.FreeSpace`               | —                                                   | —                             |

> **Why this order?** `MSFT_Disk` tends to surface **truer** hardware facts on newer Windows builds. `Win32_*` provides the broadest **legacy coverage**. PnP parsing is the **hail-mary** for vendor junky USB bridges.

---

## 6) Event flow & timing

1. **Watchers online**

   * Two `ManagementEventWatcher` instances subscribe to creation/deletion of `Win32_DiskDrive`.
   * WQL query interval `WITHIN 2` seconds to reduce jitter while keeping CPU near zero.

2. **Event fires**

   * We capture the `TargetInstance` and mark `eventType` = `Attached` or `Detached`.

3. **Rebind to object**

   * Re-open the instance as a `ManagementObject` to enable `ASSOCIATORS OF {…}` queries.

4. **Enrich**

   * Lookup `MSFT_Disk` for disk number (by `Win32_DiskDrive.Index`).
   * Fill each field per the fallback matrix.
   * If `eventType == Attached`, poll up to \~4s for `DriveLetter` to appear.

5. **Emit**

   * Build the PS object with **trimmed** strings. Return it to the caller.

---

## 7) Threading & reentrancy

* `GetNextEvent()` is **blocking** and **self-contained**. Each call:

  * Creates its own watchers.
  * Cleans them up (Stop/Dispose) before returning.
* You **can** run multiple concurrent calls (not typical). They won’t share state.

---

## 8) Error handling

* **Timeout:** returns `null` if you pass a `timeoutMilliseconds` and nothing happens.
* **Provider gaps:** if a source class/property is missing/unavailable, we continue down the **fallback chain**.
* **Association failures:** any exceptions during `ASSOCIATORS` or queries are caught; missing fields land as **empty string or zero**.
* **String normalization:** every string property passes through `.Trim()` to kill vendor padding and spurious whitespace.

---

## 9) Performance notes

* **Event delivery:** WMI intrinsic providers use sampling; with `WITHIN 2`, expect typical latency of **\~0–2s** from physical attach to event.
* **Mount discovery:** drive letters can lag device creation; the **≤4s** post-attach poll catches the usual case without blocking forever.
* **CPU/mem:** watchers are cheap; the heaviest cost is a couple of WMI queries and a handful of ASSOCIATORS around event time.

---

## 10) Security & rights

* **Regular user** context is usually enough to receive disk creation/deletion events and read most `Win32_*` properties.
* Certain fields (e.g., `Win32_DiskDrive.SerialNumber`) may require **elevated rights** or simply be blank depending on vendor/bridge.
* If `root\microsoft\windows\storage` is unavailable (older SKUs), `MSFT_Disk` fallbacks gracefully to `Win32_*`.

---

## 11) Build & packaging

### 11.1 Project

* **Target Framework:** `.NET Framework 4.8` (max compatibility with Windows PowerShell 5.1).
* **Refs:** `System.Management` + compile-time `Microsoft.PowerShell.5.ReferenceAssemblies` (runtime binds to system PowerShell).

### 11.2 Commands

* `dotnet build -c Release`
  Output: `.\bin\Release\net48\DskMon.dll`

### 11.3 Versioning

* Semantic versioning recommended (`1.0.0`, `1.1.0`, …).
* Expose `AssemblyInformationalVersion` if you want git SHA baked in.

---

## 12) Logging & diagnostics

* The library itself stays **silent** (by design).
* If you want trace logs, add a lightweight optional callback (future work), or wrap calls in PS and log `$evt` plus timestamps from your existing `Emit-TS.ps1`.

---

## 13) Limitations

* **Multiple volumes per disk:** we return the **first** mount we find. If you need **all**, expand the volume mapping to collect arrays.
* **Flaky USB bridges:** some hide/garble serial and firmware. We parse PnP as a last resort, but it’s not gospel.
* **No hotfixes for AutoPlay:** DskMon.dll doesn’t touch system settings; keep your existing `Set-SystemAutoPlay.ps1` for kiosk flows.

---

## 14) Testing strategy

### 14.1 Functional

* **Attach/Detach matrix:** SATA, USB2/3, NVMe (internal + external), SAS (if present).
* **File system coverage:** NTFS, exFAT, FAT32, unformatted/new media.
* **PnP variants:** different enclosure chipsets (ASMedia, JMicron, Realtek).

### 14.2 Reliability

* **Timeout path:** call with a small timeout and verify `null` return.
* **Mount delay:** attach media and immediately query to confirm 4s window captures letter.
* **Privilege variance:** run under standard vs elevated to ensure fallbacks behave (especially `SerialNumber`).

### 14.3 Regression

* Snapshot known devices and compare property fills after Windows updates.

---

## 15) Future work (optional hooks)

* **`Snapshot()`**: static method returning **all current disks** with the same schema.
* **Event stream API**: long-running subscription returning an `IObservable<PSObject>` (for C# consumers).
* **All mounts**: add `mountPoints` (array) and `volumes` (objects) when you care about multi-partition externals.
* **Pluggable logger**: optional `Action<string>` or ETW source for deep diagnostics.

---

## 16) Appendix

### 16.1 WQL snippets

```sql
-- Creation/Deletion of physical disks
SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_DiskDrive';
SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_DiskDrive';

-- Associate Disk→Partition→LogicalDisk
ASSOCIATORS OF {Win32_DiskDrive.DeviceID='\\.\PHYSICALDRIVE2'} WHERE AssocClass = Win32_DiskDriveToDiskPartition;
ASSOCIATORS OF {Win32_DiskPartition.DeviceID='Disk #2, Partition #0'} WHERE AssocClass = Win32_LogicalDiskToPartition;

-- Volume by DriveLetter
SELECT Label, FileSystem, FreeSpace, Capacity FROM Win32_Volume WHERE DriveLetter = 'D:';

-- MSFT_Disk by Number (CIM)
SELECT * FROM MSFT_Disk WHERE Number = 2;
```

### 16.2 BusType map (MSFT\_Disk → string)

```
1=SCSI, 2=ATAPI, 3=ATA, 4=IEEE1394, 5=SSA, 6=Fibre, 7=USB, 8=RAID,
9=iSCSI, 10=SAS, 11=SATA, 17=NVMe; others => ""
```

---

## 17) TL;DR

* One static call: **`[DskMon.DskMon]::GetNextEvent()`**.
* Returns a **trimmed PSCustomObject** with everything you care about.
* **CIM first, WMI second, PnP parse last**—so you don’t get stiffed by weird hardware.
* Clean, predictable, built for your kiosk/ingest loop.
