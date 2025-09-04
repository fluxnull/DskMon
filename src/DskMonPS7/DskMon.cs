using System;
using System.Linq;
using System.Threading;
using System.Management;
using System.Management.Automation; // PSCustomObject (PSObject with NoteProperties)
using System.Collections.Generic;

namespace DskMon
{
    /// <summary>
    /// DskMon.dll — fires when physical disks attach/detach and returns a PSCustomObject.
    /// Static usage from PowerShell:  [DskMon]::GetNextEvent(TimeoutMs)
    /// </summary>
    public static class DskMon
    {
        /// <summary>
        /// Blocks until a disk is attached or detached (or timeout).
        /// Returns a PSCustomObject with trimmed fields; null on timeout.
        /// </summary>
        /// <param name="timeoutMilliseconds">Wait time in ms; Timeout.Infinite for no timeout.</param>
        public static PSObject GetNextEvent(int timeoutMilliseconds = Timeout.Infinite)
        {
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

                attachW.Start();
                detachW.Start();

                bool fired = signal.WaitOne(timeoutMilliseconds);

                try { attachW.Stop(); } catch { }
                try { detachW.Stop(); } catch { }

                attachW.EventArrived -= onAttach;
                detachW.EventArrived -= onDetach;

                if (!fired || targetInstance == null) return null;

                // Rebind as ManagementObject to enable ASSOCIATORS queries
                using (var disk = new ManagementObject((string)targetInstance["__PATH"]))
                {
                    return BuildEventObject(disk, eventType);
                }
            }
        }

        // ───────────────────────────── Implementation ─────────────────────────────

        private static PSObject BuildEventObject(ManagementObject win32Disk, string eventType)
        {
            // Try to bind MSFT_Disk by Win32 index — best for serial/firmware/bus/size
            var msftDisk = GetMsftDiskByWin32Index(win32Disk); // may be null

            // Core identifiers (fallback chains)
            string deviceId = GetStr(win32Disk, "DeviceID"); // \\.\PHYSICALDRIVE#
            string pnpDeviceID = Coalesce(
                GetStr(win32Disk, "PNPDeviceID"),
                GetPnPFromWin32Disk(deviceId)
            );

            int diskNumber = (int)CoalesceUL(
                GetUInt(win32Disk, "Index"),
                GetUInt(msftDisk, "Number")
            );

            string serialNumber = Coalesce(
                GetStr(msftDisk, "SerialNumber"),
                GetStr(win32Disk, "SerialNumber"),
                GetSerialFromPhysicalMedia(deviceId),
                ParseSerialFromPnP(pnpDeviceID)
            );

            string firmware = Coalesce(
                GetStr(msftDisk, "FirmwareVersion"),
                GetStr(win32Disk, "FirmwareRevision")
            );

            string model = Coalesce(
                GetStr(msftDisk, "Model"),
                GetStr(win32Disk, "Model"),
                GetStr(win32Disk, "Caption")
            );

            string manufacturer = Coalesce(
                GetStr(win32Disk, "Manufacturer"),
                VendorFromPnP(pnpDeviceID),
                GetStr(msftDisk, "FriendlyName")
            );

            string friendlyName = Coalesce(
                GetStr(msftDisk, "FriendlyName"),
                TryGetStr(win32Disk, "FriendlyName"),
                model
            );

            string caption = Coalesce(
                GetStr(win32Disk, "Caption"),
                model
            );

            string hardwareID = GetHardwareIdFromPnP(pnpDeviceID);

            string interfaceType = Coalesce(
                BusTypeToString(GetUInt(msftDisk, "BusType")),
                GetStr(win32Disk, "InterfaceType")
            );

            long sizeBytes = (long)CoalesceUL(
                GetUInt(msftDisk, "Size"),
                (ulong)GetLong(win32Disk, "Size")
            );

            // Volume mapping — first good hit wins
            string mountPoint = "";
            string volumeName = "";
            string fileSystem = "";
            long freeBytes = 0;

            if (eventType == "Attached")
            {
                // Windows can raise creation before letter assignment; short poll to catch it
                var deadline = DateTime.UtcNow.AddSeconds(4);
                while (DateTime.UtcNow < deadline && string.IsNullOrEmpty(mountPoint))
                {
                    TryPopulateVolumeFields(win32Disk, out mountPoint, out volumeName, out fileSystem, out freeBytes);
                    if (!string.IsNullOrEmpty(mountPoint)) break;
                    Thread.Sleep(250);
                }
            }
            else
            {
                // Best-effort snapshot on detach (letter may be gone)
                TryPopulateVolumeFields(win32Disk, out mountPoint, out volumeName, out fileSystem, out freeBytes);
            }

            // Build PSCustomObject with trimmed NoteProperties (camelCase names)
            var o = new PSObject();
            AddProp(o, "eventType",        eventType);
            AddProp(o, "diskNumber",       diskNumber);
            AddProp(o, "name",             deviceId);
            AddProp(o, "serialNumber",     serialNumber);
            AddProp(o, "model",            model);
            AddProp(o, "firmwareVersion",  firmware);
            AddProp(o, "manufacturer",     manufacturer);
            AddProp(o, "pnpDeviceID",      pnpDeviceID);
            AddProp(o, "friendlyName",     friendlyName);
            AddProp(o, "caption",          caption);
            AddProp(o, "hardwareID",       hardwareID);
            AddProp(o, "mountPoint",       mountPoint);
            AddProp(o, "volumeName",       volumeName);
            AddProp(o, "interfaceType",    interfaceType);
            AddProp(o, "fileSystem",       fileSystem);
            AddProp(o, "sizeInBytes",      sizeBytes);
            AddProp(o, "freeSpaceInBytes", freeBytes);

            return o;
        }

        // ──────────────── Volume + Mount helpers ────────────────

        private static void TryPopulateVolumeFields(
            ManagementObject win32Disk,
            out string mountPoint,
            out string volumeName,
            out string fileSystem,
            out long freeBytes)
        {
            mountPoint = "";
            volumeName = "";
            fileSystem = "";
            freeBytes = 0;

            try
            {
                // Win32_DiskDrive → Win32_DiskPartition
                using (var parts = new ManagementObjectSearcher("root\\CIMV2",
                    $"ASSOCIATORS OF {{{win32Disk.Path.RelativePath}}} WHERE AssocClass = Win32_DiskDriveToDiskPartition"))
                {
                    foreach (ManagementObject part in parts.Get())
                    {
                        try
                        {
                            // Win32_DiskPartition → Win32_LogicalDisk
                            using (var ldks = new ManagementObjectSearcher("root\\CIMV2",
                                $"ASSOCIATORS OF {{{part.Path.RelativePath}}} WHERE AssocClass = Win32_LogicalDiskToPartition"))
                            {
                                foreach (ManagementObject ldk in ldks.Get())
                                {
                                    string candidate = Trim(GetStr(ldk, "DeviceID")); // e.g., "D:"
                                    if (!string.IsNullOrEmpty(candidate))
                                    {
                                        // Prefer Win32_Volume (FS + FreeSpace + Label)
                                        if (!TryReadVolumeByDriveLetter(candidate, out volumeName, out fileSystem, out freeBytes))
                                        {
                                            // Fallback to LogicalDisk for FS + Free
                                            fileSystem = Coalesce(fileSystem, Trim(GetStr(ldk, "FileSystem")));
                                            freeBytes = (freeBytes == 0) ? GetLong(ldk, "FreeSpace") : freeBytes;
                                            volumeName = Coalesce(volumeName, Trim(GetStr(ldk, "VolumeName")));
                                        }

                                        mountPoint = candidate;
                                        return; // first usable mapping wins
                                    }
                                }
                            }
                        }
                        finally { part.Dispose(); }
                    }
                }
            }
            catch
            {
                // swallow; leave empty/zero if not available
            }
        }

        private static bool TryReadVolumeByDriveLetter(string driveLetter,
            out string label, out string fs, out long free)
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
                        label = Trim(GetStr(v, "Label"));
                        fs = Trim(GetStr(v, "FileSystem"));
                        free = GetLong(v, "FreeSpace");
                        return true;
                    }
                }
            }
            catch { }
            return false;
        }

        // ──────────────── MSFT_Disk wiring ────────────────

        private static ManagementObject GetMsftDiskByWin32Index(ManagementObject win32Disk)
        {
            try
            {
                var idx = GetUInt(win32Disk, "Index");
                var scope = new ManagementScope(@"\\.\root\microsoft\windows\storage");
                scope.Connect();
                using (var search = new ManagementObjectSearcher(scope,
                    new ObjectQuery($"SELECT * FROM MSFT_Disk WHERE Number = {idx}")))
                {
                    foreach (ManagementObject mo in search.Get())
                        return mo;
                }
            }
            catch { }
            return null;
        }

        private static string BusTypeToString(ulong busType)
        {
            switch ((int)busType)
            {
                case 1:  return "SCSI";
                case 2:  return "ATAPI";
                case 3:  return "ATA";
                case 4:  return "IEEE1394";
                case 5:  return "SSA";
                case 6:  return "Fibre";
                case 7:  return "USB";
                case 8:  return "RAID";
                case 9:  return "iSCSI";
                case 10: return "SAS";
                case 11: return "SATA";
                case 17: return "NVMe";
                default: return ""; // Unknown/Other
            }
        }

        // ──────────────── PnP + fallback helpers ────────────────

        private static string GetPnPFromWin32Disk(string deviceId)
        {
            try
            {
                if (string.IsNullOrEmpty(deviceId)) return "";
                string esc = EscapeForWql(deviceId);
                using (var q = new ManagementObjectSearcher("root\\CIMV2",
                    $"SELECT PNPDeviceID FROM Win32_DiskDrive WHERE DeviceID = '{esc}'"))
                {
                    foreach (ManagementObject mo in q.Get())
                        return Trim(GetStr(mo, "PNPDeviceID"));
                }
            }
            catch { }
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
                        var arr = mo["HardwareID"] as string[];
                        if (arr != null && arr.Length > 0)
                            return Trim(arr[0]); // first is usually fine
                    }
                }
            }
            catch { }
            return "";
        }

        private static string GetSerialFromPhysicalMedia(string deviceId)
        {
            try
            {
                if (string.IsNullOrEmpty(deviceId)) return "";
                string esc = EscapeForWql(deviceId);
                using (var q = new ManagementObjectSearcher("root\\CIMV2",
                    $"SELECT SerialNumber FROM Win32_PhysicalMedia WHERE Tag = '{esc}'"))
                {
                    foreach (ManagementObject mo in q.Get())
                        return Trim(GetStr(mo, "SerialNumber"));
                }
            }
            catch { }
            return "";
        }

        private static string ParseSerialFromPnP(string pnp)
        {
            if (string.IsNullOrEmpty(pnp)) return "";
            try
            {
                var tail = pnp.Split('\\').LastOrDefault();
                if (string.IsNullOrEmpty(tail)) return "";

                if (pnp.StartsWith("USBSTOR"))
                {
                    int cut = tail.IndexOf('&');
                    if (cut > 0)
                    {
                        return Trim(tail.Substring(0, cut));
                    }
                }

                return Trim(tail);
            }
            catch { return ""; }
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
                return Trim(pnp.Substring(i, j - i));
            }
            catch { return ""; }
        }

        // ──────────────── Mini util layer ────────────────

        private static string Trim(string s) => s == null ? "" : s.Trim();

        private static string GetStr(ManagementBaseObject mo, string propName)
        {
            try
            {
                var p = mo?.Properties[propName];
                return Trim(p?.Value?.ToString());
            }
            catch { return ""; }
        }

        private static string TryGetStr(ManagementBaseObject mo, string propName)
        {
            var p = mo?.Properties[propName];
            if (p == null || p.Value == null) return "";
            return Trim(p.Value.ToString());
        }

        private static ulong GetUInt(ManagementBaseObject mo, string propName)
        {
            try
            {
                var v = mo?.Properties[propName]?.Value;
                if (v == null) return 0;
                return Convert.ToUInt64(v);
            }
            catch { return 0; }
        }

        private static long GetLong(ManagementBaseObject mo, string propName)
        {
            try
            {
                var v = mo?.Properties[propName]?.Value;
                if (v == null) return 0;
                return Convert.ToInt64(v);
            }
            catch { return 0; }
        }

        private static string EscapeForWql(string s)
        {
            if (s == null) return "";
            // WQL strings: escape backslashes and single quotes
            return s.Replace("\\", "\\\\").Replace("'", "\\'");
        }

        private static string Coalesce(params string[] vals)
        {
            foreach (var v in vals)
            {
                var t = Trim(v);
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
            if (value is string sv) o.Properties.Add(new PSNoteProperty(name, Trim(sv)));
            else o.Properties.Add(new PSNoteProperty(name, value));
        }
    }
}
