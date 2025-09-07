using System;
using System.Linq;
using System.Threading;
using System.Management;
using System.Management.Automation;
using System.Collections.Generic;

namespace DskMon
{
    /// <summary>
    /// DskMon.dll — Monitors physical disk attach/detach events and returns a PSCustomObject.
    /// Static usage from PowerShell: [DskMon]::GetNextEvent(TimeoutMs)
    /// Thread safety: Each GetNextEvent call is self-contained; concurrent calls create independent watchers.
    /// </summary>
    public static class DskMon
    {
        private static readonly Dictionary<string, string> _pnpCache = new Dictionary<string, string>();

        /// <summary>
        /// Blocks until a disk is attached or detached (or timeout).
        /// Returns a PSCustomObject with trimmed fields; null on timeout.
        /// </summary>
        /// <param name="timeoutMilliseconds">Wait time in ms; Timeout.Infinite for no timeout.</param>
        /// <param name="pollTimeoutSeconds">Max seconds to poll for volume assignment on attach (default: 4).</param>
        /// <param name="pollIntervalMs">Polling interval in ms for volume assignment (default: 250).</param>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if timeoutMilliseconds, pollTimeoutSeconds, or pollIntervalMs is negative.</exception>
        public static PSObject GetNextEvent(
            int timeoutMilliseconds = Timeout.Infinite,
            int pollTimeoutSeconds = 4,
            int pollIntervalMs = 250)
        {
            if (timeoutMilliseconds < Timeout.Infinite)
                throw new ArgumentOutOfRangeException(nameof(timeoutMilliseconds), "Timeout must be non-negative or Timeout.Infinite.");
            if (pollTimeoutSeconds < 0)
                throw new ArgumentOutOfRangeException(nameof(pollTimeoutSeconds), "Poll timeout must be non-negative.");
            if (pollIntervalMs < 0)
                throw new ArgumentOutOfRangeException(nameof(pollIntervalMs), "Poll interval must be non-negative.");

            var attachQ = new WqlEventQuery(
                "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_DiskDrive'");
            var detachQ = new WqlEventQuery(
                "SELECT * FROM __InstanceDeletionEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_DiskDrive'");

            using (var attachW = new ManagementEventWatcher("root\\CIMV2", attachQ.QueryString))
            using (var detachW = new ManagementEventWatcher("root\\CIMV2", detachQ.QueryString))
            {
                var signal = new AutoResetEvent(false);
                string eventType = null;
                ManagementBaseObject targetInstance = null;

                EventArrivedEventHandler onAttach = (s, e) =>
                {
                    eventType = "Attached";
                    targetInstance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
                    signal.Set();
                };
                EventArrivedEventHandler onDetach = (s, e) =>
                {
                    eventType = "Detached";
                    targetInstance = (ManagementBaseObject)e.NewEvent["TargetInstance"];
                    signal.Set();
                };

                attachW.EventArrived += onAttach;
                detachW.EventArrived += onDetach;

                try
                {
                    attachW.Start();
                    detachW.Start();

                    bool fired = signal.WaitOne(timeoutMilliseconds);
                    if (!fired || targetInstance == null)
                        return null;

                    // Rebind as ManagementObject for ASSOCIATORS queries
                    using (var disk = new ManagementObject((string)targetInstance["__PATH"]))
                    {
                        return BuildEventObject(disk, eventType, pollTimeoutSeconds, pollIntervalMs);
                    }
                }
                catch (ManagementException ex)
                {
                    // Log critical WMI errors but maintain fallback behavior
                    throw new InvalidOperationException($"WMI failure: {ex.Message}", ex);
                }
                finally
                {
                    attachW.EventArrived -= onAttach;
                    detachW.EventArrived -= onDetach;
                    attachW.Stop();
                    detachW.Stop();
                }
            }
        }

        // ───────────────────────────── Implementation ─────────────────────────────

        private static PSObject BuildEventObject(
            ManagementObject win32Disk,
            string eventType,
            int pollTimeoutSeconds,
            int pollIntervalMs)
        {
            using (var msftDisk = GetMsftDiskByWin32Index(win32Disk))
            {
                string deviceId = GetStr(win32Disk, "DeviceID");
                string pnpDeviceID = Coalesce(
                    GetStr(win32Disk, "PNPDeviceID"),
                    GetPnPFromWin32Disk(deviceId));

                var props = new Dictionary<string, object>
                {
                    { "eventType", eventType },
                    { "diskNumber", (int)CoalesceUL(GetUInt(win32Disk, "Index"), GetUInt(msftDisk, "Number")) },
                    { "name", deviceId },
                    { "serialNumber", Coalesce(GetStr(msftDisk, "SerialNumber"), GetStr(win32Disk, "SerialNumber"),
                        GetSerialFromPhysicalMedia(deviceId), ParseSerialFromPnP(pnpDeviceID)) },
                    { "model", Coalesce(GetStr(msftDisk, "Model"), GetStr(win32Disk, "Model"), GetStr(win32Disk, "Caption")) },
                    { "firmwareVersion", Coalesce(GetStr(msftDisk, "FirmwareVersion"), GetStr(win32Disk, "FirmwareRevision")) },
                    { "manufacturer", Coalesce(GetStr(win32Disk, "Manufacturer"), VendorFromPnP(pnpDeviceID), GetStr(msftDisk, "FriendlyName")) },
                    { "pnpDeviceID", pnpDeviceID },
                    { "friendlyName", Coalesce(GetStr(msftDisk, "FriendlyName"), GetStr(win32Disk, "FriendlyName"), GetStr(msftDisk, "Model")) },
                    { "caption", Coalesce(GetStr(win32Disk, "Caption"), GetStr(msftDisk, "Model")) },
                    { "hardwareID", GetHardwareIdFromPnP(pnpDeviceID) },
                    { "interfaceType", Coalesce(BusTypeToString(GetUInt(msftDisk, "BusType")), GetStr(win32Disk, "InterfaceType")) },
                    { "sizeInBytes", (long)CoalesceUL(GetUInt(msftDisk, "Size"), (ulong)GetLong(win32Disk, "Size")) }
                };

                // Volume mapping: collect all mount points and volumes
                var mountPoints = new List<string>();
                var volumeNames = new List<string>();
                var fileSystems = new List<string>();
                var freeBytesList = new List<long>();

                if (eventType == "Attached")
                {
                    var deadline = DateTime.UtcNow.AddSeconds(pollTimeoutSeconds);
                    while (DateTime.UtcNow < deadline && mountPoints.Count == 0)
                    {
                        PopulateVolumeFields(win32Disk, mountPoints, volumeNames, fileSystems, freeBytesList);
                        if (mountPoints.Count > 0) break;
                        Thread.Sleep(pollIntervalMs);
                    }
                }
                else
                {
                    PopulateVolumeFields(win32Disk, mountPoints, volumeNames, fileSystems, freeBytesList);
                }

                var o = new PSObject();
                foreach (var prop in props)
                    AddProp(o, prop.Key, prop.Value);

                // Add volume arrays (single or multiple)
                AddProp(o, "mountPoints", mountPoints.ToArray());
                AddProp(o, "volumeNames", volumeNames.ToArray());
                AddProp(o, "fileSystems", fileSystems.ToArray());
                AddProp(o, "freeSpaceInBytes", freeBytesList.ToArray());

                // Backward compatibility: single-value fields
                AddProp(o, "mountPoint", mountPoints.FirstOrDefault() ?? "");
                AddProp(o, "volumeName", volumeNames.FirstOrDefault() ?? "");
                AddProp(o, "fileSystem", fileSystems.FirstOrDefault() ?? "");
                AddProp(o, "freeSpaceInBytesSingle", freeBytesList.FirstOrDefault());

                return o;
            }
        }

        // ──────────────── Volume + Mount helpers ────────────────

        private static void PopulateVolumeFields(
            ManagementObject win32Disk,
            List<string> mountPoints,
            List<string> volumeNames,
            List<string> fileSystems,
            List<long> freeBytesList)
        {
            try
            {
                using (var parts = new ManagementObjectSearcher("root\\CIMV2",
                    $"ASSOCIATORS OF {{{win32Disk.Path.RelativePath}}} WHERE AssocClass = Win32_DiskDriveToDiskPartition"))
                {
                    foreach (ManagementObject part in parts.Get())
                    {
                        using (part)
                        using (var ldks = new ManagementObjectSearcher("root\\CIMV2",
                            $"ASSOCIATORS OF {{{part.Path.RelativePath}}} WHERE AssocClass = Win32_LogicalDiskToPartition"))
                        {
                            foreach (ManagementObject ldk in ldks.Get())
                            {
                                using (ldk)
                                {
                                    string candidate = GetStr(ldk, "DeviceID");
                                    if (!string.IsNullOrEmpty(candidate))
                                    {
                                        string volumeName = "";
                                        string fileSystem = "";
                                        long freeBytes = 0;

                                        if (!TryReadVolumeByDriveLetter(candidate, out volumeName, out fileSystem, out freeBytes))
                                        {
                                            fileSystem = Coalesce(fileSystem, GetStr(ldk, "FileSystem"));
                                            freeBytes = freeBytes == 0 ? GetLong(ldk, "FreeSpace") : freeBytes;
                                            volumeName = Coalesce(volumeName, GetStr(ldk, "VolumeName"));
                                        }

                                        mountPoints.Add(candidate);
                                        volumeNames.Add(volumeName);
                                        fileSystems.Add(fileSystem);
                                        freeBytesList.Add(freeBytes);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (ManagementException ex)
            {
                // Log but don't fail; empty lists are returned
                System.Diagnostics.Debug.WriteLine($"Volume query failed: {ex.Message}");
            }
        }

        private static bool TryReadVolumeByDriveLetter(
            string driveLetter,
            out string label,
            out string fs,
            out long free)
        {
            label = ""; fs = ""; free = 0;
            try
            {
                string esc = EscapeForWql(driveLetter);
                using (var vols = new ManagementObjectSearcher("root\\CIMV2",
                    $"SELECT Label, FileSystem, FreeSpace, Capacity FROM Win32_Volume WHERE DriveLetter = '{esc}'"))
                {
                    foreach (ManagementObject v in vols.Get())
                    {
                        using (v)
                        {
                            label = GetStr(v, "Label");
                            fs = GetStr(v, "FileSystem");
                            free = GetLong(v, "FreeSpace");
                            return true;
                        }
                    }
                }
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"Volume read failed for {driveLetter}: {ex.Message}");
            }
            return false;
        }

        // ──────────────── MSFT_Disk wiring ────────────────

        private static ManagementObject GetMsftDiskByWin32Index(ManagementObject win32Disk)
        {
            try
            {
                var idx = GetUInt(win32Disk, "Index");
                var scope = new ManagementScope(@"\\.\root\microsoft\windows\storage");
                {
                    scope.Connect();
                    using (var search = new ManagementObjectSearcher(scope,
                        new ObjectQuery($"SELECT * FROM MSFT_Disk WHERE Number = {idx}")))
                    {
                        foreach (ManagementObject mo in search.Get())
                            return mo; // Caller must dispose
                    }
                }
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"MSFT_Disk query failed: {ex.Message}");
            }
            return null;
        }

        private static string BusTypeToString(ulong busType)
        {
            switch ((int)busType)
            {
                case 1: return "SCSI";
                case 2: return "ATAPI";
                case 3: return "ATA";
                case 4: return "IEEE1394";
                case 5: return "SSA";
                case 6: return "Fibre";
                case 7: return "USB";
                case 8: return "RAID";
                case 9: return "iSCSI";
                case 10: return "SAS";
                case 11: return "SATA";
                case 17: return "NVMe";
                default: return "";
            }
        }

        // ──────────────── PnP + fallback helpers ────────────────

        private static string GetPnPFromWin32Disk(string deviceId)
        {
            if (string.IsNullOrEmpty(deviceId)) return "";
            if (_pnpCache.TryGetValue(deviceId, out string cached)) return cached;

            try
            {
                string esc = EscapeForWql(deviceId);
                using (var q = new ManagementObjectSearcher("root\\CIMV2",
                    $"SELECT PNPDeviceID FROM Win32_DiskDrive WHERE DeviceID = '{esc}'"))
                {
                    foreach (ManagementObject mo in q.Get())
                    {
                        using (mo)
                        {
                            string pnp = GetStr(mo, "PNPDeviceID");
                            _pnpCache[deviceId] = pnp;
                            return pnp;
                        }
                    }
                }
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"PnP query failed for {deviceId}: {ex.Message}");
            }
            return "";
        }

        private static string GetHardwareIdFromPnP(string pnpDeviceID)
        {
            if (string.IsNullOrEmpty(pnpDeviceID)) return "";
            try
            {
                string esc = EscapeForWql(pnpDeviceID);
                using (var srch = new ManagementObjectSearcher("root\\CIMV2",
                    $"SELECT HardwareID FROM Win32_PnPEntity WHERE PNPDeviceID = '{esc}'"))
                {
                    foreach (ManagementObject mo in srch.Get())
                    {
                        using (mo)
                        {
                            var arr = mo["HardwareID"] as string[];
                            return arr != null && arr.Length > 0 ? GetStr(mo, "HardwareID", arr[0]) : "";
                        }
                    }
                }
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"HardwareID query failed for {pnpDeviceID}: {ex.Message}");
            }
            return "";
        }

        private static string GetSerialFromPhysicalMedia(string deviceId)
        {
            if (string.IsNullOrEmpty(deviceId)) return "";
            try
            {
                string esc = EscapeForWql(deviceId);
                using (var q = new ManagementObjectSearcher("root\\CIMV2",
                    $"SELECT SerialNumber FROM Win32_PhysicalMedia WHERE Tag = '{esc}'"))
                {
                    foreach (ManagementObject mo in q.Get())
                    {
                        using (mo)
                            return GetStr(mo, "SerialNumber");
                    }
                }
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"PhysicalMedia query failed for {deviceId}: {ex.Message}");
            }
            return "";
        }

        private static string ParseSerialFromPnP(string pnp)
        {
            if (string.IsNullOrEmpty(pnp)) return "";
            try
            {
                var tail = pnp.Split('\\').LastOrDefault();
                if (string.IsNullOrEmpty(tail)) return "";
                int cut = tail.IndexOf('&');
                var core = (cut > 0) ? tail.Substring(0, cut) : tail;
                return GetStr(null, null, core); // Use GetStr for consistent trimming
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PnP serial parse failed: {ex.Message}");
                return "";
            }
        }

        private static string VendorFromPnP(string pnp)
        {
            if (string.IsNullOrEmpty(pnp)) return "";
            try
            {
                var up = pnp.ToUpperInvariant();
                var tag = "&VEN_";
                var i = up.IndexOf(tag);
                if (i < 0) return "";
                i += tag.Length;
                var j = up.IndexOf('&', i);
                if (j < 0) j = up.Length;
                return GetStr(null, null, pnp.Substring(i, j - i));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"PnP vendor parse failed: {ex.Message}");
                return "";
            }
        }

        // ──────────────── Mini util layer ────────────────

        private static string GetStr(ManagementBaseObject mo, string propName, string defaultValue = "")
        {
            try
            {
                if (mo == null || propName == null) return defaultValue?.Trim() ?? "";
                var p = mo.Properties[propName];
                return p?.Value?.ToString()?.Trim() ?? "";
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetStr failed for {propName}: {ex.Message}");
                return defaultValue?.Trim() ?? "";
            }
        }

        private static ulong GetUInt(ManagementBaseObject mo, string propName)
        {
            try
            {
                var v = mo?.Properties[propName]?.Value;
                return v == null ? 0 : Convert.ToUInt64(v);
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetUInt failed for {propName}: {ex.Message}");
                return 0;
            }
        }

        private static long GetLong(ManagementBaseObject mo, string propName)
        {
            try
            {
                var v = mo?.Properties[propName]?.Value;
                return v == null ? 0 : Convert.ToInt64(v);
            }
            catch (ManagementException ex)
            {
                System.Diagnostics.Debug.WriteLine($"GetLong failed for {propName}: {ex.Message}");
                return 0;
            }
        }

        private static string EscapeForWql(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            // Enhanced escaping for WQL: backslashes, single quotes, and control chars
            return s.Replace("\\", "\\\\")
                    .Replace("'", "\\'")
                    .Replace("\0", "")
                    .Replace("\r", "")
                    .Replace("\n", "");
        }

        private static string Coalesce(params string[] vals)
        {
            foreach (var v in vals)
            {
                var t = GetStr(null, null, v);
                if (!string.IsNullOrEmpty(t)) return t;
            }
            return "";
        }

        private static ulong CoalesceUL(params ulong[] vals)
        {
            foreach (var v in vals) if (v != 0) return v;
            return 0;
        }

        private static void AddProp(PSObject o, string name, object value)
        {
            if (value is string sv) o.Properties.Add(new PSNoteProperty(name, GetStr(null, null, sv)));
            else o.Properties.Add(new PSNoteProperty(name, value));
        }
    }
}