namespace USBTools;

/// <summary>
/// Provides human-readable descriptions for Windows and WIM API error codes
/// </summary>
public static class ErrorCodeHelper
{
    public static string GetErrorDescription(int errorCode)
    {
        return errorCode switch
        {
            // Common WIM API errors
            1006 => "ERROR_INVALID_HANDLE - The handle is invalid. The WIM file may be corrupt or inaccessible.",
            1465 => "ERROR_BAD_PROFILE - Invalid data format. The XML structure may be malformed.",
            2 => "ERROR_FILE_NOT_FOUND - The system cannot find the file specified.",
            3 => "ERROR_PATH_NOT_FOUND - The system cannot find the path specified.",
            5 => "ERROR_ACCESS_DENIED - Access is denied. Try running as Administrator.",
            32 => "ERROR_SHARING_VIOLATION - The process cannot access the file because it is being used by another process.",
            80 => "ERROR_FILE_EXISTS - The file exists and cannot be overwritten.",
            87 => "ERROR_INVALID_PARAMETER - The parameter is incorrect.",
            112 => "ERROR_DISK_FULL - There is not enough space on the disk.",
            1314 => "ERROR_PRIVILEGE_NOT_HELD - A required privilege is not held by the client. Run as Administrator.",
            
            // WIM-specific errors
            4390 => "ERROR_WIM_NOT_BOOTABLE - The WIM file is not bootable.",
            4391 => "ERROR_WIM_HEADER_NOT_FOUND - The WIM header was not found.",
            4392 => "ERROR_WIM_INVALID_RESOURCE - The WIM resource is invalid.",
            
            // Generic
            _ => $"Windows Error Code {errorCode}. Check https://learn.microsoft.com/en-us/windows/win32/debug/system-error-codes for details."
        };
    }
}
