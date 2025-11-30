using System.Diagnostics;

namespace USBTools;

/// <summary>
/// Validates that a drive is a removable USB drive to prevent accidental overwrites
/// </summary>
public static class UsbDriveValidator
{
    public static bool IsUsbDrive(int diskNumber)
    {
        try
        {
            Logger.Log($"Validating disk {diskNumber} is a USB drive...", Logger.LogLevel.Info);

            var output = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty BusType");
            var busType = output.Trim();

            Logger.Log($"Disk {diskNumber} bus type: {busType}", Logger.LogLevel.Debug);

            var isUsb = busType.Equals("USB", StringComparison.OrdinalIgnoreCase);

            if (isUsb)
            {
                Logger.Log($"Disk {diskNumber} validated as USB drive", Logger.LogLevel.Info);
            }
            else
            {
                Logger.Log($"Disk {diskNumber} is NOT a USB drive (Bus Type: {busType})", Logger.LogLevel.Warning);
            }

            return isUsb;
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, $"Error validating disk {diskNumber}");
            return false;
        }
    }

    public static bool IsUsbDrive(string driveLetter)
    {
        try
        {
            var diskNumber = GetDiskNumberFromDriveLetter(driveLetter);
            if (diskNumber == -1)
            {
                Logger.Log($"Could not determine disk number for drive {driveLetter}", Logger.LogLevel.Error);
                return false;
            }

            return IsUsbDrive(diskNumber);
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, $"Error validating drive {driveLetter}");
            return false;
        }
    }

    private static int GetDiskNumberFromDriveLetter(string driveLetter)
    {
        try
        {
            var letter = driveLetter.TrimEnd(':').ToUpper();
            var output = ExecutePowerShell($"Get-Partition -DriveLetter {letter} | Select-Object -ExpandProperty DiskNumber");
            if (int.TryParse(output.Trim(), out var diskNum))
            {
                return diskNum;
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error getting disk number");
        }
        return -1;
    }

    public static string GetDiskInfo(int diskNumber)
    {
        try
        {
            var sizeOutput = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty Size");
            var friendlyNameOutput = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty FriendlyName");
            var busTypeOutput = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty BusType");

            long.TryParse(sizeOutput.Trim(), out var sizeBytes);
            var sizeGB = sizeBytes / (1024.0 * 1024 * 1024);

            return $"Disk {diskNumber}: {friendlyNameOutput.Trim()}\n" +
                   $"  Size: {sizeGB:F2} GB\n" +
                   $"  Bus Type: {busTypeOutput.Trim()}";
        }
        catch
        {
            return $"Disk {diskNumber}: Unable to retrieve details";
        }
    }

    private static string ExecutePowerShell(string command)
    {
        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -NonInteractive -Command \"{command}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi);
        if (process == null)
        {
            throw new Exception("Failed to start PowerShell");
        }

        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();

        if (process.ExitCode != 0 && !string.IsNullOrWhiteSpace(error))
        {
            throw new Exception($"PowerShell error: {error}");
        }

        return output;
    }
}
