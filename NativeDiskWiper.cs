using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace USBTools;

public static class NativeDiskWiper
{
    private const uint GENERIC_READ = 0x80000000;
    private const uint GENERIC_WRITE = 0x40000000;
    private const uint OPEN_EXISTING = 3;
    private const uint FILE_ATTRIBUTE_NORMAL = 0x80;
    private const uint FILE_SHARE_READ = 0x00000001;
    private const uint FILE_SHARE_WRITE = 0x00000002;

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

    private const uint IOCTL_DISK_UPDATE_PROPERTIES = 0x00070050;

    public static void WipeHeader(int diskNumber)
    {
        string drivePath = $@"\\.\PHYSICALDRIVE{diskNumber}";
        Logger.Log($"Opening {drivePath} for header wiping...", Logger.LogLevel.Debug);

        using var handle = CreateFile(
            drivePath,
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            IntPtr.Zero,
            OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL,
            IntPtr.Zero);

        if (handle.IsInvalid)
        {
            int error = Marshal.GetLastWin32Error();
            throw new IOException($"Failed to open disk {diskNumber}. Error Code: {error}");
        }

        using var stream = new FileStream(handle, FileAccess.ReadWrite);
        
        // Wipe first 1MB to ensure MBR/GPT signatures and partition tables are gone
        int wipeSize = 1024 * 1024; 
        byte[] buffer = new byte[wipeSize]; // Zeros by default

        Logger.Log($"Writing {wipeSize} bytes of zeros to start of disk...", Logger.LogLevel.Info);
        stream.Write(buffer, 0, buffer.Length);
        stream.Flush();
        
        Logger.Log("Disk header wiped successfully.", Logger.LogLevel.Info);

        // Force OS to update partition table
        Logger.Log("Forcing disk property update...", Logger.LogLevel.Debug);
        uint bytesReturned;
        bool result = DeviceIoControl(
            handle,
            IOCTL_DISK_UPDATE_PROPERTIES,
            IntPtr.Zero,
            0,
            IntPtr.Zero,
            0,
            out bytesReturned,
            IntPtr.Zero);

        if (!result)
        {
            Logger.Log($"Warning: Failed to update disk properties. Error: {Marshal.GetLastWin32Error()}", Logger.LogLevel.Warning);
        }
        else
        {
            Logger.Log("Disk properties updated.", Logger.LogLevel.Info);
        }
    }
}
