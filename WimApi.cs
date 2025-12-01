using System.Runtime.InteropServices;
using System.Text;

namespace USBTools;

/// <summary>
/// P/Invoke wrapper for wimgapi.dll
/// </summary>
public static class WimApi
{
    private const string WimgapiDll = "wimgapi.dll";

    // Access Rights
    public const uint GENERIC_READ = 0x80000000;
    public const uint GENERIC_WRITE = 0x40000000;

    // WIM Creation Disposition Flags
    public const uint WIM_CREATE_NEW = 1;
    public const uint WIM_CREATE_ALWAYS = 2;
    public const uint WIM_OPEN_EXISTING = 3;
    public const uint WIM_OPEN_ALWAYS = 4;

    // WIM Flags
    public const uint WIM_FLAG_VERIFY = 0x00000002;
    public const uint WIM_FLAG_INDEX = 0x00000004;
    public const uint WIM_FLAG_NO_APPLY = 0x00000008;
    public const uint WIM_FLAG_NO_DIRACL = 0x00000010;
    public const uint WIM_FLAG_NO_FILEACL = 0x00000020;
    public const uint WIM_FLAG_SHARE_WRITE = 0x00000040;

    // Compression Types
    public const uint WIM_COMPRESS_NONE = 0;
    public const uint WIM_COMPRESS_XPRESS = 1;
    public const uint WIM_COMPRESS_LZX = 2;
    public const uint WIM_COMPRESS_LZMS = 3;

    // WIM Info Classes
    public const uint WIM_INFO_ATTRIBUTES = 1;
    public const uint WIM_INFO_COMPRESSION = 2;

    // WIM Message Constants
    public const uint WIM_MSG_TEXT = 0x9400;
    public const uint WIM_MSG_PROGRESS = 0x9401;
    public const uint WIM_MSG_PROCESS = 0x9402;
    public const uint WIM_MSG_SCANNING = 0x9403;
    public const uint WIM_MSG_SETRANGE = 0x9404;
    public const uint WIM_MSG_SETPOS = 0x9405;
    public const uint WIM_MSG_STEPIT = 0x9406;
    public const uint WIM_MSG_COMPRESS = 0x9407;
    public const uint WIM_MSG_ERROR = 0x9408;
    public const uint WIM_MSG_ALIGNMENT = 0x9409;
    public const uint WIM_MSG_RETRY = 0x940A;
    public const uint WIM_MSG_SPLIT = 0x940B;
    public const uint WIM_MSG_FILEINFO = 0x940C;
    public const uint WIM_MSG_INFO = 0x940D;
    public const uint WIM_MSG_WARNING = 0x940E;
    public const uint WIM_MSG_CHK_PROCESS = 0x940F;

    [DllImport(WimgapiDll, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr WIMCreateFile(
        string wimPath,
        uint desiredAccess,
        uint creationDisposition,
        uint flagsAndAttributes,
        uint compressionType,
        out uint creationResult);

    [DllImport(WimgapiDll, SetLastError = true)]
    public static extern bool WIMCloseHandle(IntPtr wimHandle);

    [DllImport(WimgapiDll, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool WIMSetReferenceFile(
        IntPtr wimHandle,
        string path,
        uint mode);

    [DllImport(WimgapiDll, SetLastError = true)]
    public static extern bool WIMGetImageInformation(
        IntPtr wimHandle,
        out IntPtr imageInfo,
        out uint sizeOfImageInfo);

    [DllImport(WimgapiDll, SetLastError = true)]
    public static extern IntPtr WIMLoadImage(
        IntPtr wimHandle,
        uint imageIndex);

    [DllImport(WimgapiDll, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool WIMApplyImage(
        IntPtr imageHandle,
        string path,
        uint applyFlags);

    [DllImport(WimgapiDll, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern IntPtr WIMCaptureImage(
        IntPtr wimHandle,
        string path,
        uint captureFlags);

    [DllImport(WimgapiDll, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool WIMSetImageInformation(
        IntPtr imageHandle,
        IntPtr imageInfo,
        uint sizeOfImageInfo);

    [DllImport(WimgapiDll, SetLastError = true)]
    public static extern uint WIMGetImageCount(IntPtr wimHandle);

    [DllImport(WimgapiDll, CharSet = CharSet.Unicode, SetLastError = true)]
    public static extern bool WIMSetTemporaryPath(
        IntPtr wimHandle,
        string tempPath);

    // Message callback for progress
    public delegate uint WIMMessageCallback(
        uint messageId,
        IntPtr wParam,
        IntPtr lParam,
        IntPtr userData);

    [DllImport(WimgapiDll, SetLastError = true)]
    public static extern uint WIMRegisterMessageCallback(
        IntPtr wimHandle,
        WIMMessageCallback callback,
        IntPtr userData);

    [DllImport(WimgapiDll, SetLastError = true)]
    public static extern bool WIMUnregisterMessageCallback(
        IntPtr wimHandle,
        WIMMessageCallback callback);

    public static bool IsAvailable()
    {
        try
        {
            // Try to load wimgapi.dll
            var ptr = LoadLibrary(WimgapiDll);
            if (ptr != IntPtr.Zero)
            {
                FreeLibrary(ptr);
                return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr LoadLibrary(string dllToLoad);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool FreeLibrary(IntPtr hModule);

    public static uint GetCompressionType(string compressionLevel)
    {
        return compressionLevel.ToLowerInvariant() switch
        {
            "none" => WIM_COMPRESS_NONE,
            "fast" => WIM_COMPRESS_XPRESS,
            "max" => WIM_COMPRESS_LZX,
            _ => WIM_COMPRESS_XPRESS
        };
    }
}
