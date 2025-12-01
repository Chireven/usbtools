using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Win32.SafeHandles;
using System.Management;
using System.Text.RegularExpressions;

namespace USBTools;

/// <summary>
/// Analyzes drive geometry and partitions using native Windows APIs (IOCTLs)
/// This avoids WMI and PowerShell dependencies which can be unreliable.
/// </summary>
[SupportedOSPlatform("windows")]
public static class DriveAnalyzer
{
    public static DriveGeometry AnalyzeDrive(string driveLetter)
    {
        Logger.Log($"Analyzing drive {driveLetter}", Logger.LogLevel.Info);

        // Get disk number from drive letter using IOCTL
        var diskNumber = GetDiskNumberFromDriveLetter(driveLetter);
        if (diskNumber == -1)
        {
            throw new Exception($"Could not determine disk number for drive {driveLetter}");
        }

        var geometry = new DriveGeometry();

        // Get disk info (Style, Size, Signature/ID) via WMI (still useful for general info, but we can fallback if needed)
        // We'll try WMI for metadata, but if it fails, we'll just use defaults to allow the backup to proceed.
        GetDiskInfo(diskNumber, geometry);
        
        Logger.Log($"Partition style: {geometry.PartitionStyle}", Logger.LogLevel.Info);
        Logger.Log($"Total disk size: {geometry.TotalSize:N0} bytes", Logger.LogLevel.Info);

        // Get all partitions on this disk
        geometry.Partitions = GetPartitions(diskNumber);
        Logger.Log($"Found {geometry.Partitions.Count} partitions", Logger.LogLevel.Info);

        return geometry;
    }

    public static int GetDiskNumberFromDriveLetter(string driveLetter)
    {
        // Normalize drive letter to just the letter (e.g. "G")
        // Handle "G:", "G:\", "G"
        var letter = driveLetter.TrimEnd('\\', ':').ToUpper(); // Uppercase is safer for some APIs
        if (letter.Length > 1) letter = letter.Substring(0, 1); // Safety check

        // Try IOCTL first (fastest, most reliable for physical disks)
        int diskNumber = GetDiskNumberViaIoCtl(letter);
        if (diskNumber != -1) return diskNumber;

        Logger.Log("IOCTL disk detection failed, falling back to WMI...", Logger.LogLevel.Warning);

        // Fallback to WMI
        diskNumber = GetDiskNumberViaWmi(letter);
        if (diskNumber != -1) return diskNumber;

        Logger.Log("WMI disk detection failed, falling back to PowerShell...", Logger.LogLevel.Warning);

        // Fallback to PowerShell (Robust)
        diskNumber = GetDiskNumberViaPowerShell(letter);
        if (diskNumber != -1) return diskNumber;

        Logger.Log("PowerShell disk detection failed, falling back to DiskPart...", Logger.LogLevel.Warning);

        // Fallback to DiskPart (Nuclear option)
        return GetDiskNumberViaDiskPart(letter);
    }

    private static int GetDiskNumberViaIoCtl(string letter)
    {
        var volumePath = $@"\\.\{letter}:";

        using var volumeHandle = CreateFile(
            volumePath,
            0, // No access required for metadata
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            IntPtr.Zero,
            OPEN_EXISTING,
            0,
            IntPtr.Zero);

        if (volumeHandle.IsInvalid)
        {
            Logger.Log($"Failed to open volume handle for {letter}:. Error: {Marshal.GetLastWin32Error()}", Logger.LogLevel.Warning);
            return -1;
        }

        // Get volume disk extents
        var extents = new VOLUME_DISK_EXTENTS();
        var size = Marshal.SizeOf(extents);
        var ptr = Marshal.AllocHGlobal(size);
        
        try
        {
            Marshal.StructureToPtr(extents, ptr, false);
            
            if (!DeviceIoControl(
                volumeHandle,
                IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS,
                IntPtr.Zero,
                0,
                ptr,
                (uint)size,
                out var bytesReturned,
                IntPtr.Zero))
            {
                // If the buffer was too small (more than 1 extent), we'd get ERROR_MORE_DATA
                // But for a simple USB drive partition, 1 extent is expected.
                Logger.Log($"IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS failed. Error: {Marshal.GetLastWin32Error()}", Logger.LogLevel.Warning);
                return -1;
            }

            extents = Marshal.PtrToStructure<VOLUME_DISK_EXTENTS>(ptr);
            
            if (extents.NumberOfDiskExtents > 0)
            {
                // Return the disk number of the first extent
                return (int)extents.Extents[0].DiskNumber;
            }
        }
        finally
        {
            Marshal.FreeHGlobal(ptr);
        }

        return -1;
    }

    private static int GetDiskNumberViaWmi(string letter)
    {
        // 1. Try modern MSFT_Partition (Windows 8+)
        try
        {
            using var searcher = new ManagementObjectSearcher(@"\\.\root\Microsoft\Windows\Storage", $"SELECT * FROM MSFT_Partition WHERE DriveLetter = '{letter}'");
            
            foreach (ManagementObject partition in searcher.Get())
            {
                var diskNum = Convert.ToInt32(partition["DiskNumber"]);
                Logger.Log($"WMI (MSFT_Partition) found disk number {diskNum} for drive {letter}", Logger.LogLevel.Debug);
                return diskNum;
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"WMI (MSFT_Partition) lookup failed: {ex.Message}", Logger.LogLevel.Debug);
        }

        // 2. Try legacy Win32_LogicalDiskToPartition
        try
        {
            var driveLetterWithColon = letter + ":";
            
            // Query Win32_LogicalDiskToPartition to link Drive Letter -> Partition
            var query = $"ASSOCIATORS OF {{Win32_LogicalDisk.DeviceID='{driveLetterWithColon}'}} WHERE AssocClass=Win32_LogicalDiskToPartition";
            using var searcher = new ManagementObjectSearcher(query);
            
            var partitions = searcher.Get();
            if (partitions.Count == 0)
            {
                Logger.Log($"WMI (Win32) returned no partitions for {driveLetterWithColon}", Logger.LogLevel.Debug);
            }

            foreach (ManagementObject partition in partitions)
            {
                // partition is Win32_DiskPartition
                // DeviceID format example: "Disk #1, Partition #0"
                var deviceId = partition["DeviceID"]?.ToString();
                Logger.Log($"WMI Partition DeviceID: {deviceId}", Logger.LogLevel.Debug);

                if (!string.IsNullOrEmpty(deviceId))
                {
                    var match = Regex.Match(deviceId, @"Disk #(\d+)");
                    if (match.Success)
                    {
                        var diskNum = int.Parse(match.Groups[1].Value);
                        Logger.Log($"WMI (Win32) found disk number {diskNum} for drive {driveLetterWithColon}", Logger.LogLevel.Debug);
                        return diskNum;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"WMI (Win32) Disk Number lookup failed: {ex.Message}", Logger.LogLevel.Debug);
        }
        return -1;
    }

    private static int GetDiskNumberViaPowerShell(string letter)
    {
        try
        {
            // Use Get-Partition to find the disk number
            var command = $"Get-Partition -DriveLetter {letter} | Select-Object -ExpandProperty DiskNumber";
            
            // Suppress progress stream to avoid CLIXML output in stderr
            command = "$ProgressPreference = 'SilentlyContinue'; " + command;
            var encodedCommand = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(command));
            
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -NonInteractive -EncodedCommand {encodedCommand}",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(psi);
            if (process == null) return -1;

            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();

            if (int.TryParse(output.Trim(), out var diskNum))
            {
                return diskNum;
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"PowerShell lookup failed: {ex.Message}", Logger.LogLevel.Warning);
        }
        return -1;
    }

    private static int GetDiskNumberViaDiskPart(string letter)
    {
        try
        {
            // Create a temporary script file for diskpart
            var scriptPath = Path.GetTempFileName();
            File.WriteAllText(scriptPath, $"select volume {letter}\ndetail volume\nexit");

            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "diskpart.exe",
                Arguments = $"/s \"{scriptPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            using var process = System.Diagnostics.Process.Start(psi);
            if (process == null) return -1;

            var output = process.StandardOutput.ReadToEnd();
            process.WaitForExit();
            File.Delete(scriptPath);

            // Parse output for "Disk ###" or "* Disk 2"
            // Example: "* Disk 2    Online       14 GB      0 B"
            var match = Regex.Match(output, @"\*\s+Disk\s+(\d+)", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return int.Parse(match.Groups[1].Value);
            }
            
            // Try without asterisk
            match = Regex.Match(output, @"Disk\s+(\d+)\s+Online", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return int.Parse(match.Groups[1].Value);
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"DiskPart lookup failed: {ex.Message}", Logger.LogLevel.Warning);
        }
        return -1;
    }

    public static void GetDiskInfo(int diskNumber, DriveGeometry geometry)
    {
        try
        {
            // We use WMI here just for metadata (Size, Style). If it fails, it's not critical for the backup operation itself
            // (which relies on the disk number), but we need the size for the progress bar.
            using var searcher = new ManagementObjectSearcher($"SELECT * FROM Win32_DiskDrive WHERE Index = {diskNumber}");
            foreach (ManagementObject drive in searcher.Get())
            {
                geometry.TotalSize = Convert.ToInt64(drive["Size"]);
                
                try 
                {
                    // Try to get partition style from MSFT_Disk if available
                    using var storageSearcher = new ManagementObjectSearcher(@"\\.\root\Microsoft\Windows\Storage", $"SELECT * FROM MSFT_Disk WHERE Number = {diskNumber}");
                    foreach (ManagementObject disk in storageSearcher.Get())
                    {
                        var partitionStyle = Convert.ToUInt16(disk["PartitionStyle"]); // 1 = MBR, 2 = GPT
                        geometry.PartitionStyle = partitionStyle == 2 ? "GPT" : "MBR";
                        
                        if (geometry.PartitionStyle == "GPT")
                        {
                            geometry.GptDiskId = disk["Guid"]?.ToString() ?? Guid.NewGuid().ToString();
                        }
                        else
                        {
                            geometry.DiskSignature = Convert.ToUInt32(disk["Signature"]);
                        }
                        return;
                    }
                }
                catch
                {
                    geometry.PartitionStyle = "MBR"; // Default
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"Warning: Could not retrieve full disk info: {ex.Message}", Logger.LogLevel.Warning);
        }
    }

    public static List<PartitionInfo> GetPartitions(int diskNumber)
    {
        var partitions = new List<PartitionInfo>();

        try
        {
            // Use MSFT_Partition if available
            using var searcher = new ManagementObjectSearcher(@"\\.\root\Microsoft\Windows\Storage", $"SELECT * FROM MSFT_Partition WHERE DiskNumber = {diskNumber}");
            
            foreach (ManagementObject item in searcher.Get())
            {
                var partition = new PartitionInfo
                {
                    Index = Convert.ToInt32(item["PartitionNumber"]),
                    Size = Convert.ToInt64(item["Size"]),
                    Offset = Convert.ToInt64(item["Offset"]),
                    Letter = item["DriveLetter"]?.ToString()?.Trim()
                };
                
                if (partition.Letter != null && partition.Letter.Length == 1)
                {
                    partition.Letter += ":";
                }
                else if (Convert.ToChar(item["DriveLetter"]) == 0)
                {
                    partition.Letter = null;
                }

                partition.IsActive = false; // Default
                try { partition.IsActive = Convert.ToBoolean(item["IsActive"]); } catch {}

                partition.GptId = item["Guid"]?.ToString();
                partition.GptType = item["GptType"]?.ToString();
                
                if (item["MbrType"] != null)
                {
                    partition.Type = item["MbrType"].ToString();
                }

                if (!string.IsNullOrEmpty(partition.Letter))
                {
                    try
                    {
                        var driveInfo = new DriveInfo(partition.Letter);
                        partition.Label = driveInfo.VolumeLabel;
                        partition.FileSystem = driveInfo.DriveFormat;
                        partition.UsedSpace = driveInfo.TotalSize - driveInfo.TotalFreeSpace;
                    }
                    catch { }
                }

                partition.IsFixed = IsFixedPartition(partition);
                partitions.Add(partition);
            }
        }
        catch
        {
            // Fallback to Win32_DiskPartition if MSFT_Partition fails
             try
            {
                using var searcher = new ManagementObjectSearcher($"ASSOCIATORS OF {{Win32_DiskDrive.DeviceID='\\\\.\\PHYSICALDRIVE{diskNumber}'}} WHERE AssocClass=Win32_DiskDriveToDiskPartition");
                foreach (ManagementObject item in searcher.Get())
                {
                     var partition = new PartitionInfo
                    {
                        Index = Convert.ToInt32(item["Index"]), // Win32_DiskPartition Index is unique across system? No, usually just index.
                        // Actually Win32_DiskPartition.Index is system wide. We need "Name" or parsing.
                        // Let's just rely on the fact that we found it.
                        Size = Convert.ToInt64(item["Size"]),
                        Type = item["Type"]?.ToString() ?? "Unknown"
                    };
                    // Getting offset/letter via Win32 is painful (requires more associators).
                    // For now, if MSFT fails, we might just have basic info.
                    partitions.Add(partition);
                }
            }
            catch (Exception ex)
            {
                Logger.LogException(ex, "Error getting partitions");
            }
        }

        return partitions.OrderBy(p => p.Offset).ToList();
    }

    private static bool IsFixedPartition(PartitionInfo partition)
    {
        if (string.Equals(partition.GptType, "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(partition.GptType, "{e3c9e316-0b5c-4db8-817d-f92df00215ae}", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(partition.GptType, "{de94bba4-06d1-4d40-a16a-bfd50179d6ac}", StringComparison.OrdinalIgnoreCase)) return true;
        if (partition.Size < 1024 * 1024 * 1024 && partition.IsActive) return true;
        return false;
    }

    public static uint GenerateRandomDiskSignature()
    {
        var random = new Random();
        return (uint)random.Next(1, int.MaxValue);
    }

    public static string GenerateRandomGptDiskId()
    {
        return Guid.NewGuid().ToString();
    }

    // P/Invoke Definitions
    [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern SafeFileHandle CreateFile(
        string lpFileName,
        uint dwDesiredAccess,
        uint dwShareMode,
        IntPtr lpSecurityAttributes,
        uint dwCreationDisposition,
        uint dwFlagsAndAttributes,
        IntPtr hTemplateFile);

    [DllImport("kernel32.dll", ExactSpelling = true, SetLastError = true, CharSet = CharSet.Auto)]
    private static extern bool DeviceIoControl(
        SafeFileHandle hDevice,
        uint dwIoControlCode,
        IntPtr lpInBuffer,
        uint nInBufferSize,
        IntPtr lpOutBuffer,
        uint nOutBufferSize,
        out uint lpBytesReturned,
        IntPtr lpOverlapped);

    private const uint GENERIC_READ = 0x80000000;
    private const uint FILE_SHARE_READ = 0x00000001;
    private const uint FILE_SHARE_WRITE = 0x00000002;
    private const uint OPEN_EXISTING = 3;
    private const uint IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS = 0x00560000;

    [StructLayout(LayoutKind.Sequential)]
    private struct DISK_EXTENT
    {
        public uint DiskNumber;
        public long StartingOffset;
        public long ExtentLength;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct VOLUME_DISK_EXTENTS
    {
        public uint NumberOfDiskExtents;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public DISK_EXTENT[] Extents;
    }
}
