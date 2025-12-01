using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace USBTools;

public static class NativeDiskManager
{
    #region Constants
    private const uint GENERIC_READ = 0x80000000;
    private const uint GENERIC_WRITE = 0x40000000;
    private const uint OPEN_EXISTING = 3;
    private const uint FILE_SHARE_READ = 0x00000001;
    private const uint FILE_SHARE_WRITE = 0x00000002;
    
    private const uint FSCTL_LOCK_VOLUME = 0x00090018;
    private const uint FSCTL_DISMOUNT_VOLUME = 0x00090020;
    private const uint IOCTL_DISK_GET_DRIVE_GEOMETRY_EX = 0x000700A0;
    private const uint IOCTL_DISK_DELETE_DRIVE_LAYOUT = 0x0007C010;
    private const uint IOCTL_DISK_SET_DRIVE_LAYOUT_EX = 0x00070050;
    private const uint IOCTL_DISK_UPDATE_PROPERTIES = 0x00070140;
    
    private const uint PARTITION_STYLE_MBR = 0;
    private const uint PARTITION_STYLE_GPT = 1;
    #endregion

    #region P/Invoke Declarations
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
    #endregion

    #region Structures
    [StructLayout(LayoutKind.Sequential)]
    private struct DISK_GEOMETRY_EX
    {
        public DISK_GEOMETRY Geometry;
        public long DiskSize;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
        public byte[] Data;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DISK_GEOMETRY
    {
        public long Cylinders;
        public uint MediaType;
        public uint TracksPerCylinder;
        public uint SectorsPerTrack;
        public uint BytesPerSector;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct DRIVE_LAYOUT_INFORMATION_EX
    {
        [FieldOffset(0)]
        public uint PartitionStyle;
        [FieldOffset(4)]
        public uint PartitionCount;
        [FieldOffset(8)]
        public DRIVE_LAYOUT_INFORMATION_MBR Mbr;
        [FieldOffset(8)]
        public DRIVE_LAYOUT_INFORMATION_GPT Gpt;
        [FieldOffset(48)]
        public PARTITION_INFORMATION_EX PartitionEntry;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DRIVE_LAYOUT_INFORMATION_MBR
    {
        public uint Signature;
        public uint CheckSum;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct DRIVE_LAYOUT_INFORMATION_GPT
    {
        public Guid DiskId;
        public long StartingUsableOffset;
        public long UsableLength;
        public uint MaxPartitionCount;
    }

    [StructLayout(LayoutKind.Sequential, Size = 144)]
    private struct PARTITION_INFORMATION_EX
    {
        public uint PartitionStyle;
        public uint Padding1;
        public long StartingOffset;
        public long PartitionLength;
        public uint PartitionNumber;
        [MarshalAs(UnmanagedType.I1)]
        public bool RewritePartition;
        [MarshalAs(UnmanagedType.I1)]
        public bool IsServicePartition;
        public ushort Padding2;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 112)]
        public byte[] UnionData;
    }
    #endregion

    public static long GetDiskSize(int diskNumber)
    {
        string drivePath = $@"\\.\PHYSICALDRIVE{diskNumber}";
        
        using var handle = CreateFile(
            drivePath,
            GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            IntPtr.Zero,
            OPEN_EXISTING,
            0,
            IntPtr.Zero);

        if (handle.IsInvalid)
        {
            int error = Marshal.GetLastWin32Error();
            throw new IOException($"Failed to open disk {diskNumber}. Error: {error}");
        }

        int size = Marshal.SizeOf<DISK_GEOMETRY_EX>();
        IntPtr buffer = Marshal.AllocHGlobal(size);
        try
        {
            uint bytesReturned;
            bool result = DeviceIoControl(
                handle,
                IOCTL_DISK_GET_DRIVE_GEOMETRY_EX,
                IntPtr.Zero,
                0,
                buffer,
                (uint)size,
                out bytesReturned,
                IntPtr.Zero);

            if (!result)
            {
                int error = Marshal.GetLastWin32Error();
                throw new IOException($"Failed to get disk geometry. Error: {error}");
            }

            var geometry = Marshal.PtrToStructure<DISK_GEOMETRY_EX>(buffer);
            return geometry.DiskSize;
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    public static void SetPartitionLayout(int diskNumber, List<PartitionInfo> partitions, bool useGpt)
    {
        string drivePath = $@"\\.\PHYSICALDRIVE{diskNumber}";
        Logger.Log($"Opening {drivePath} for partition layout configuration...", Logger.LogLevel.Info);
        
        // PHASE 1: Delete existing partition table (then close handle to flush cache)
        using (var handle = CreateFile(
            drivePath,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            IntPtr.Zero,
            OPEN_EXISTING,
            0,
            IntPtr.Zero))
        {
            if (handle.IsInvalid)
            {
                int error = Marshal.GetLastWin32Error();
                throw new IOException($"Failed to open disk {diskNumber}. Error: {error}");
            }

            // Lock and dismount
            Logger.Log("Locking and dismounting volume...", Logger.LogLevel.Debug);
            uint bytesReturned;
            DeviceIoControl(handle, FSCTL_LOCK_VOLUME, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);
            DeviceIoControl(handle, FSCTL_DISMOUNT_VOLUME, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);

            // Delete existing partition table
            Logger.Log("Deleting existing partition table...", Logger.LogLevel.Info);
            bool deleteResult = DeviceIoControl(handle, IOCTL_DISK_DELETE_DRIVE_LAYOUT, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);
            if (!deleteResult)
            {
                int deleteError = Marshal.GetLastWin32Error();
                Logger.Log($"Warning: Failed to delete drive layout. Error: {deleteError}", Logger.LogLevel.Warning);
            }
            else
            {
                Logger.Log("Partition table deleted successfully.", Logger.LogLevel.Info);
            }

            // Update properties
            Logger.Log("Updating disk properties after deletion...", Logger.LogLevel.Debug);
            DeviceIoControl(handle, IOCTL_DISK_UPDATE_PROPERTIES, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);
        } // Handle closes here

        // CRITICAL: Ask user to physically remove and reinsert the drive
        Logger.Log("", Logger.LogLevel.Info);
        Logger.Log("================================================================================", Logger.LogLevel.Info);
        Logger.Log("IMPORTANT: To complete the disk reset, please:", Logger.LogLevel.Info);
        Logger.Log($"  1. Physically UNPLUG the USB drive from PhysicalDrive{diskNumber}", Logger.LogLevel.Info);
        Logger.Log("  2. Wait 5 seconds", Logger.LogLevel.Info);
        Logger.Log("  3. REPLUG the USB drive", Logger.LogLevel.Info);
        Logger.Log("  4. Press any key to continue...", Logger.LogLevel.Info);
        Logger.Log("================================================================================", Logger.LogLevel.Info);
        Logger.Log("", Logger.LogLevel.Info);
        
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("╔════════════════════════════════════════════════════════════════╗");
        Console.WriteLine("║                    ⚠️  ACTION REQUIRED  ⚠️                      ║");
        Console.WriteLine("╠════════════════════════════════════════════════════════════════╣");
        Console.WriteLine($"║  Please UNPLUG the USB drive (PhysicalDrive{diskNumber})                  ║");
        Console.WriteLine("║  Wait 5 seconds, then REPLUG it                               ║");
        Console.WriteLine("║                                                                ║");
        Console.WriteLine("║  This forces Windows to completely refresh the disk state.    ║");
        Console.WriteLine("╠════════════════════════════════════════════════════════════════╣");
        Console.WriteLine("║  Press ANY KEY when the drive has been reinserted...          ║");
        Console.WriteLine("╚════════════════════════════════════════════════════════════════╝");
        Console.ResetColor();
        Console.WriteLine();
        
        Console.ReadKey(true);
        Logger.Log("User confirmed drive has been reinserted. Continuing...", Logger.LogLevel.Info);

        // Wait a bit more for OS to fully recognize the drive
        Logger.Log("Waiting for OS to fully enumerate the drive...", Logger.LogLevel.Debug);
        Thread.Sleep(3000);

        // PHASE 2: Reopen with fresh handle and set new partition layout
        using (var handle = CreateFile(
            drivePath,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            IntPtr.Zero,
            OPEN_EXISTING,
            0,
            IntPtr.Zero))
        {
            if (handle.IsInvalid)
            {
                int error = Marshal.GetLastWin32Error();
                throw new IOException($"Failed to reopen disk {diskNumber}. Error: {error}");
            }

            // Construct layout structure
            int headerSize = 48;
            int partitionEntrySize = 144;
            int structSize = headerSize + (partitions.Count * partitionEntrySize);
            
            IntPtr layoutPtr = Marshal.AllocHGlobal(structSize);
            try
            {
                // Initialize to zeros
                for (int i = 0; i < structSize; i++)
                {
                    Marshal.WriteByte(layoutPtr, i, 0);
                }

                var layout = new DRIVE_LAYOUT_INFORMATION_EX
                {
                    PartitionStyle = useGpt ? PARTITION_STYLE_GPT : PARTITION_STYLE_MBR,
                    PartitionCount = (uint)partitions.Count
                };

                if (useGpt)
                {
                    layout.Gpt = new DRIVE_LAYOUT_INFORMATION_GPT
                    {
                        DiskId = Guid.NewGuid(),
                        StartingUsableOffset = 1024 * 1024,
                        MaxPartitionCount = 128
                    };
                }
                else
                {
                    layout.Mbr = new DRIVE_LAYOUT_INFORMATION_MBR
                    {
                        Signature = (uint)DateTime.Now.Ticks
                    };
                }

                Marshal.StructureToPtr(layout, layoutPtr, false);

                // Write partition entries
                IntPtr partitionPtr = IntPtr.Add(layoutPtr, 48);
                foreach (var partition in partitions)
                {
                    var partInfo = new PARTITION_INFORMATION_EX
                    {
                        PartitionStyle = useGpt ? PARTITION_STYLE_GPT : PARTITION_STYLE_MBR,
                        Padding1 = 0,
                        StartingOffset = partition.Offset,
                        PartitionLength = partition.Size,
                        PartitionNumber = (uint)partition.Index,
                        RewritePartition = true,
                        IsServicePartition = false,
                        Padding2 = 0,
                        UnionData = new byte[112]
                    };

                    if (useGpt)
                    {
                        var gptType = GetGptPartitionType(partition.FileSystem);
                        var gptId = Guid.NewGuid();
                        ulong attributes = partition.FileSystem.ToUpper() == "FAT32" ? 0x0000000000000001UL : 0UL;
                        
                        Array.Copy(gptType.ToByteArray(), 0, partInfo.UnionData, 0, 16);
                        Array.Copy(gptId.ToByteArray(), 0, partInfo.UnionData, 16, 16);
                        Array.Copy(BitConverter.GetBytes(attributes), 0, partInfo.UnionData, 32, 8);
                    }
                    else
                    {
                        partInfo.UnionData[0] = GetMbrPartitionType(partition.FileSystem);
                        partInfo.UnionData[1] = (byte)(partition.Index == 1 ? 1 : 0);
                        partInfo.UnionData[2] = 1;
                    }

                    Marshal.StructureToPtr(partInfo, partitionPtr, false);
                    partitionPtr = IntPtr.Add(partitionPtr, partitionEntrySize);
                }

                // Set layout
                Logger.Log($"Setting {(useGpt ? "GPT" : "MBR")} partition layout with {partitions.Count} partition(s)...", Logger.LogLevel.Info);
                uint bytesReturned;
                bool result = DeviceIoControl(
                    handle,
                    IOCTL_DISK_SET_DRIVE_LAYOUT_EX,
                    layoutPtr,
                    (uint)structSize,
                    IntPtr.Zero,
                    0,
                    out bytesReturned,
                    IntPtr.Zero);

                if (!result)
                {
                    int error = Marshal.GetLastWin32Error();
                    throw new IOException($"Failed to set drive layout. Error: {error}");
                }

                Logger.Log("Partition layout set successfully.", Logger.LogLevel.Info);

                // Update properties
                Logger.Log("Updating disk properties...", Logger.LogLevel.Debug);
                DeviceIoControl(handle, IOCTL_DISK_UPDATE_PROPERTIES, IntPtr.Zero, 0, IntPtr.Zero, 0, out bytesReturned, IntPtr.Zero);
            }
            finally
            {
                Marshal.FreeHGlobal(layoutPtr);
            }
        }
    }

    private static Guid GetGptPartitionType(string fileSystem)
    {
        return fileSystem.ToUpper() switch
        {
            "FAT32" => new Guid("EBD0A0A2-B9E5-4433-87C0-68B6B72699C7"),
            "NTFS" => new Guid("EBD0A0A2-B9E5-4433-87C0-68B6B72699C7"),
            _ => new Guid("EBD0A0A2-B9E5-4433-87C0-68B6B72699C7")
        };
    }

    private static byte GetMbrPartitionType(string fileSystem)
    {
        return fileSystem.ToUpper() switch
        {
            "FAT32" => 0x0C,
            "NTFS" => 0x07,
            _ => 0x07
        };
    }

    /// <summary>
    /// Test version of SetPartitionLayout that doesn't delete the partition table first
    /// Used by diagnostic command to isolate IOCTL issues
    /// </summary>
    public static void SetPartitionLayoutTest(int diskNumber, List<PartitionInfo> partitions, bool useGpt)
    {
        string drivePath = $@"\\.\PHYSICALDRIVE{diskNumber}";
        
       using var handle = CreateFile(
            drivePath,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            IntPtr.Zero,
            OPEN_EXISTING,
            0,
            IntPtr.Zero);

        if (handle.IsInvalid)
        {
            int error = Marshal.GetLastWin32Error();
            throw new IOException($"Failed to open disk {diskNumber}. Error: {error}");
        }

        int headerSize = 48;
        int partitionEntrySize = 144;
        int structSize = headerSize + (partitions.Count * partitionEntrySize);
        
        IntPtr layoutPtr = Marshal.AllocHGlobal(structSize);
        try
        {
            for (int i = 0; i < structSize; i++)
            {
                Marshal.WriteByte(layoutPtr, i, 0);
            }

            var layout = new DRIVE_LAYOUT_INFORMATION_EX
            {
                PartitionStyle = useGpt ? PARTITION_STYLE_GPT : PARTITION_STYLE_MBR,
                PartitionCount = (uint)partitions.Count
            };

            if (useGpt)
            {
                layout.Gpt = new DRIVE_LAYOUT_INFORMATION_GPT
                {
                    DiskId = Guid.NewGuid(),
                    StartingUsableOffset = 1024 * 1024,
                    MaxPartitionCount = 128
                };
            }
            else
            {
                layout.Mbr = new DRIVE_LAYOUT_INFORMATION_MBR
                {
                    Signature = (uint)DateTime.Now.Ticks
                };
            }

            Marshal.StructureToPtr(layout, layoutPtr, false);

            IntPtr partitionPtr = IntPtr.Add(layoutPtr, 48);
            foreach (var partition in partitions)
            {
                var partInfo = new PARTITION_INFORMATION_EX
                {
                    PartitionStyle = useGpt ? PARTITION_STYLE_GPT : PARTITION_STYLE_MBR,
                    Padding1 = 0,
                    StartingOffset = partition.Offset,
                    PartitionLength = partition.Size,
                    PartitionNumber = (uint)partition.Index,
                    RewritePartition = true,
                    IsServicePartition = false,
                    Padding2 = 0,
                    UnionData = new byte[112]
                };

                if (useGpt)
                {
                    var gptType = GetGptPartitionType(partition.FileSystem);
                    var gptId = Guid.NewGuid();
                    ulong attributes = partition.FileSystem.ToUpper() == "FAT32" ? 0x0000000000000001UL : 0UL;
                    
                    Array.Copy(gptType.ToByteArray(), 0, partInfo.UnionData, 0, 16);
                    Array.Copy(gptId.ToByteArray(), 0, partInfo.UnionData, 16, 16);
                    Array.Copy(BitConverter.GetBytes(attributes), 0, partInfo.UnionData, 32, 8);
                }
                else
                {
                    partInfo.UnionData[0] = GetMbrPartitionType(partition.FileSystem);
                    partInfo.UnionData[1] = (byte)(partition.Index == 1 ? 1 : 0);
                    partInfo.UnionData[2] = 1;
                }

                Marshal.StructureToPtr(partInfo, partitionPtr, false);
                partitionPtr = IntPtr.Add(partitionPtr, partitionEntrySize);
            }

            uint bytesReturned;
            bool result = DeviceIoControl(
                handle,
                IOCTL_DISK_SET_DRIVE_LAYOUT_EX,
                layoutPtr,
                (uint)structSize,
                IntPtr.Zero,
                0,
                out bytesReturned,
                IntPtr.Zero);

            if (!result)
            {
                int error = Marshal.GetLastWin32Error();
                throw new IOException($"Failed to set drive layout. Error: {error} (Buffer size: {structSize})");
            }
        }
        finally
        {
            Marshal.FreeHGlobal(layoutPtr);
        }
    }

    public class PartitionInfo
    {
        public int Index { get; set; }
        public long Offset { get; set; }
        public long Size { get; set; }
        public string FileSystem { get; set; } = "";
    }
}
