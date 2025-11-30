using System.Diagnostics;
using System.Runtime.InteropServices;

namespace USBTools;

/// <summary>
/// Handles restore (apply) operations from WIM to USB
/// </summary>
public static class RestoreCommand
{
    public static async Task<int> ExecuteAsync(string sourceWim, string? targetDrive, string? diskNumberStr, bool autoYes, string provider)
    {
        try
        {
            Logger.Log($"Starting restore operation", Logger.LogLevel.Info);
            Logger.Log($"Source WIM: {sourceWim}", Logger.LogLevel.Info);
            Logger.Log($"Provider: {provider}", Logger.LogLevel.Info);

            if (!File.Exists(sourceWim))
            {
                Logger.Log($"WIM file not found: {sourceWim}", Logger.LogLevel.Error);
                return 1;
            }

            // Determine target disk number
            int diskNumber;
            if (!string.IsNullOrEmpty(diskNumberStr))
            {
                // User provided disk number directly
                if (!int.TryParse(diskNumberStr, out diskNumber))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error: Invalid disk number '{diskNumberStr}'");
                    Console.ResetColor();
                    return 1;
                }
                Logger.Log($"Target Disk Number: {diskNumber}", Logger.LogLevel.Info);
            }
            else if (!string.IsNullOrEmpty(targetDrive))
            {
                // User provided drive letter - get disk number from it
                targetDrive = targetDrive.TrimEnd(':', '\\');
                Logger.Log($"Target Drive: {targetDrive}", Logger.LogLevel.Info);
                
                diskNumber = GetDiskNumberFromDrive(targetDrive);
                if (diskNumber == -1)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error: Could not determine disk number for drive {targetDrive}:");
                    Console.ResetColor();
                    return 1;
                }
                Logger.Log($"Disk number for drive {targetDrive}: is {diskNumber}", Logger.LogLevel.Info);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error: No target specified");
                Console.ResetColor();
                return 1;
            }

            // VALIDATION: Ensure target is a USB drive
            if (!UsbDriveValidator.IsUsbDrive(diskNumber))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine();
                Console.WriteLine($"ERROR: Disk {diskNumber} is NOT a USB drive!");
                Console.WriteLine();
                Console.WriteLine("For safety, this tool only allows restoring to removable USB drives.");
                Console.WriteLine("This prevents accidental data loss on system or internal drives.");
                Console.ResetColor();
                Console.WriteLine();
                return 1;
            }

            // Get metadata from WIM
            var metadata = await ReadWimMetadataAsync(sourceWim);
            if (metadata == null)
            {
                Logger.Log("Failed to read WIM metadata", Logger.LogLevel.Error);
                return 1;
            }

            Logger.Log($"Source geometry: {metadata.PartitionStyle}, {metadata.Partitions.Count} partitions", Logger.LogLevel.Info);

            // Get target disk info
            var diskInfo = UsbDriveValidator.GetDiskInfo(diskNumber);
            Logger.Log($"Target disk info:\n{diskInfo}", Logger.LogLevel.Info);

            // CONFIRMATION: Ask user to confirm (unless --yes flag is set)
            var confirmMessage = $"You are about to restore to:\n\n{diskInfo}\n";
            if (!ConfirmationHelper.Confirm(confirmMessage, autoYes))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Operation cancelled by user.");
                Console.ResetColor();
                return 1;
            }

            // Get target size for geometry adaptation
            var targetSize = GetDiskSize(diskNumber);
            Logger.Log($"Target disk size: {targetSize:N0} bytes", Logger.LogLevel.Info);

            // Validate target size
            if (targetSize < metadata.TotalSize)
            {
                Logger.Log("WARNING: Target drive is smaller than source", Logger.LogLevel.Warning);
                
                // Calculate if we can fit by shrinking variable partitions
                var fixedSize = metadata.Partitions.Where(p => p.IsFixed).Sum(p => p.Size);
                if (targetSize < fixedSize)
                {
                    Logger.Log("Target drive is too small even with shrinking", Logger.LogLevel.Error);
                    return 1;
                }
            }

            // Prepare target disk (partition and format)
            await PrepareDiskAsync(diskNumber, metadata, targetSize);

            // Determine which provider to use
            bool useWimApi;
            if (provider == "wimapi")
            {
                useWimApi = true;
                if (!WimApi.IsAvailable())
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error: WIM API requested but wimgapi.dll is not available.");
                    Console.ResetColor();
                    return 1;
                }
            }
            else if (provider == "dism")
            {
                useWimApi = false;
            }
            else // "auto"
            {
                useWimApi = WimApi.IsAvailable();
            }

            Logger.Log($"Using provider: {(useWimApi ? "WIM API" : "DISM")}", Logger.LogLevel.Info);

            if (useWimApi)
            {
                await ApplyWithWimApiAsync(sourceWim, metadata, diskNumber);
            }
            else
            {
                await ApplyWithDismAsync(sourceWim, metadata, diskNumber);
            }

            // Make bootable
            await MakeBootableAsync(diskNumber, metadata);

            Logger.Log("Restore completed successfully", Logger.LogLevel.Info);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine("✓ Restore completed successfully!");
            Console.ResetColor();
            return 0;
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Restore operation failed");
            return 1;
        }
    }

    private static async Task<DriveGeometry?> ReadWimMetadataAsync(string wimPath)
    {
        try
        {
            IntPtr wimHandle = IntPtr.Zero;

            try
            {
                wimHandle = WimApi.WIMCreateFile(
                    wimPath,
                    WimApi.GENERIC_READ,
                    WimApi.WIM_OPEN_EXISTING,
                    0,
                    0,
                    out _);

                if (wimHandle == IntPtr.Zero)
                {
                    Logger.Log("Failed to open WIM with API, trying alternate method", Logger.LogLevel.Warning);
                    return null;
                }

                // Get image information
                if (WimApi.WIMGetImageInformation(wimHandle, out var infoPtr, out var infoSize))
                {
                    var xmlInfo = Marshal.PtrToStringUni(infoPtr, (int)infoSize / 2);
                    
                    // Extract JSON from description field
                    if (xmlInfo != null)
                    {
                        var descStart = xmlInfo.IndexOf("<DESCRIPTION>");
                        var descEnd = xmlInfo.IndexOf("</DESCRIPTION>");
                        
                        if (descStart != -1 && descEnd != -1)
                        {
                            descStart += "<DESCRIPTION>".Length;
                            var json = xmlInfo.Substring(descStart, descEnd - descStart);
                            return DriveGeometry.FromJson(json);
                        }
                    }
                }
            }
            finally
            {
                if (wimHandle != IntPtr.Zero)
                {
                    WimApi.WIMCloseHandle(wimHandle);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error reading WIM metadata");
        }

        return null;
    }

    private static int GetDiskNumberFromDrive(string driveLetter)
    {
        try
        {
            var output = ExecutePowerShell($"Get-Partition -DriveLetter {driveLetter} | Select-Object -ExpandProperty DiskNumber");
            if (int.TryParse(output.Trim(), out var diskNum))
            {
                return diskNum;
            }
        }
        catch { }
        return -1;
    }

    private static long GetDiskSize(int diskNumber)
    {
        try
        {
            var output = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty Size");
            if (long.TryParse(output.Trim(), out var size))
            {
                return size;
            }
        }
        catch { }
        return 0;
    }

    private static async Task PrepareDiskAsync(int diskNumber, DriveGeometry sourceGeometry, long targetSize)
    {
        Logger.Log($"Preparing disk {diskNumber} for restore", Logger.LogLevel.Info);

        // Clean disk
        Logger.Log("Cleaning disk...", Logger.LogLevel.Info);
        ExecutePowerShell($"Clear-Disk -Number {diskNumber} -RemoveData -Confirm:$false");

        // Initialize disk with appropriate partition style
        Logger.Log($"Initializing disk as {sourceGeometry.PartitionStyle}...", Logger.LogLevel.Info);
        ExecutePowerShell($"Initialize-Disk -Number {diskNumber} -PartitionStyle {sourceGeometry.PartitionStyle}");

        // Set new disk signature
        if (sourceGeometry.PartitionStyle == "MBR")
        {
            var newSignature = DriveAnalyzer.GenerateRandomDiskSignature();
            Logger.Log($"Setting new disk signature: {newSignature}", Logger.LogLevel.Info);
            ExecutePowerShell($"Set-Disk -Number {diskNumber} -Signature {newSignature}");
        }
        else
        {
            var newGuid = DriveAnalyzer.GenerateRandomGptDiskId();
            Logger.Log($"New GPT disk ID: {newGuid}", Logger.LogLevel.Info);
            ExecutePowerShell($"Set-Disk -Number {diskNumber} -Guid '{newGuid}'");
        }

        // Create partitions based on source geometry
        var offset = 1048576L; // Start at 1MB
        var partitionNumber = 1;

        foreach (var partition in sourceGeometry.Partitions.OrderBy(p => p.Offset))
        {
            var partitionSize = partition.Size;

            // Adjust variable partition if target is larger
            if (!partition.IsFixed && targetSize > sourceGeometry.TotalSize)
            {
                var extraSpace = targetSize - sourceGeometry.TotalSize;
                partitionSize += extraSpace;
                Logger.Log($"Expanding variable partition by {extraSpace:N0} bytes", Logger.LogLevel.Info);
            }

            Logger.Log($"Creating partition {partitionNumber} ({partitionSize:N0} bytes)...", Logger.LogLevel.Info);

            if (sourceGeometry.PartitionStyle == "GPT" && partition.GptType != null)
            {
                ExecutePowerShell($"New-Partition -DiskNumber {diskNumber} -Size {partitionSize} -GptType '{partition.GptType}'");
            }
            else
            {
                var isActive = partition.IsActive ? "-IsActive" : "";
                ExecutePowerShell($"New-Partition -DiskNumber {diskNumber} -Size {partitionSize} {isActive}");
            }

            partitionNumber++;
        }

        // Format partitions
        await FormatPartitionsAsync(diskNumber, sourceGeometry);
    }

    private static async Task FormatPartitionsAsync(int diskNumber, DriveGeometry geometry)
    {
        Logger.Log("Formatting partitions...", Logger.LogLevel.Info);

        var partitions = ExecutePowerShell($"Get-Partition -DiskNumber {diskNumber} | Where-Object {{ $_.Type -ne 'Reserved' }} | Select-Object -ExpandProperty PartitionNumber");
        var partNumbers = partitions.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

        var index = 0;
        foreach (var partNumStr in partNumbers)
        {
            if (int.TryParse(partNumStr.Trim(), out var partNum) && index < geometry.Partitions.Count)
            {
                var partition = geometry.Partitions[index];
                var fs = partition.FileSystem ?? "NTFS";
                var label = partition.Label ?? $"Volume{partNum}";

                Logger.Log($"Formatting partition {partNum} as {fs} (Label: {label})...", Logger.LogLevel.Info);
                
                ExecutePowerShell($@"
                    $part = Get-Partition -DiskNumber {diskNumber} -PartitionNumber {partNum}
                    $part | Format-Volume -FileSystem {fs} -NewFileSystemLabel '{label}' -Confirm:$false
                ");

                index++;
            }
        }
    }

    private static async Task ApplyWithWimApiAsync(string wimPath, DriveGeometry geometry, int diskNumber)
    {
        Logger.Log("Applying images with WIM API", Logger.LogLevel.Info);

        IntPtr wimHandle = IntPtr.Zero;

        try
        {
            wimHandle = WimApi.WIMCreateFile(
                wimPath,
                WimApi.GENERIC_READ,
                WimApi.WIM_OPEN_EXISTING,
                0,
                0,
                out _);

            if (wimHandle == IntPtr.Zero)
            {
                throw new Exception("Failed to open WIM file");
            }

            var imageCount = WimApi.WIMGetImageCount(wimHandle);
            Logger.Log($"WIM contains {imageCount} images", Logger.LogLevel.Info);

            // Get partition drive letters
            var driveLetters = GetPartitionDriveLetters(diskNumber);

            for (uint i = 1; i <= imageCount && i <= driveLetters.Count; i++)
            {
                Logger.Log($"Applying image {i} to {driveLetters[(int)i - 1]}:\\", Logger.LogLevel.Info);

                var imageHandle = WimApi.WIMLoadImage(wimHandle, i);
                if (imageHandle == IntPtr.Zero)
                {
                    Logger.Log($"Failed to load image {i}", Logger.LogLevel.Error);
                    continue;
                }

                var targetPath = $"{driveLetters[(int)i - 1]}:\\";
                var success = WimApi.WIMApplyImage(imageHandle, targetPath, 0);

                WimApi.WIMCloseHandle(imageHandle);

                if (!success)
                {
                    Logger.Log($"Failed to apply image {i}", Logger.LogLevel.Error);
                }
                else
                {
                    Logger.Log($"Image {i} applied successfully", Logger.LogLevel.Info);
                }
            }
        }
        finally
        {
            if (wimHandle != IntPtr.Zero)
            {
                WimApi.WIMCloseHandle(wimHandle);
            }
        }
    }

    private static async Task ApplyWithDismAsync(string wimPath, DriveGeometry geometry, int diskNumber)
    {
        Logger.Log("Applying images with DISM", Logger.LogLevel.Warning);

        var driveLetters = GetPartitionDriveLetters(diskNumber);
        
        for (int i = 1; i <= geometry.Partitions.Count && i <= driveLetters.Count; i++)
        {
            Logger.Log($"Applying image {i} to {driveLetters[i - 1]}:\\", Logger.LogLevel.Info);

            var targetPath = $"{driveLetters[i - 1]}:\\";
            var success = DismWrapper.ApplyImage(wimPath, i, targetPath);

            if (!success)
            {
                Logger.Log($"Failed to apply image {i} with DISM", Logger.LogLevel.Error);
            }
            else
            {
                Logger.Log($"Image {i} applied successfully", Logger.LogLevel.Info);
            }
        }
    }

    private static List<string> GetPartitionDriveLetters(int diskNumber)
    {
        var letters = new List<string>();
        try
        {
            var output = ExecutePowerShell($@"
                Get-Partition -DiskNumber {diskNumber} | 
                Where-Object {{ $_.DriveLetter -ne $null }} | 
                Select-Object -ExpandProperty DriveLetter
            ");

            foreach (var line in output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var letter = line.Trim();
                if (!string.IsNullOrEmpty(letter))
                {
                    letters.Add(letter);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error getting partition drive letters");
        }
        return letters;
    }

    private static async Task MakeBootableAsync(int diskNumber, DriveGeometry geometry)
    {
        Logger.Log("Configuring boot...", Logger.LogLevel.Info);

        try
        {
            var driveLetters = GetPartitionDriveLetters(diskNumber);
            if (driveLetters.Count == 0)
            {
                Logger.Log("No drive letters found for boot configuration", Logger.LogLevel.Warning);
                return;
            }

            // Find Windows partition (usually the largest or last partition)
            var windowsDrive = driveLetters.Last();
            
            if (geometry.PartitionStyle == "GPT")
            {
                // For GPT, find EFI partition and Windows partition
                var efiDrive = driveLetters.Count > 1 ? driveLetters.First() : windowsDrive;
                Logger.Log($"Running bcdboot {windowsDrive}:\\Windows /s {efiDrive}: /f UEFI", Logger.LogLevel.Info);
                ExecuteCommand("bcdboot", $"{windowsDrive}:\\Windows /s {efiDrive}: /f UEFI");
            }
            else
            {
                // For MBR
                Logger.Log($"Running bcdboot {windowsDrive}:\\Windows /s {windowsDrive}: /f BIOS", Logger.LogLevel.Info);
                ExecuteCommand("bcdboot", $"{windowsDrive}:\\Windows /s {windowsDrive}: /f BIOS");
                
                // Set active partition
                ExecutePowerShell($"Set-Partition -DiskNumber {diskNumber} -PartitionNumber 1 -IsActive $true");
            }

            Logger.Log("Boot configuration completed", Logger.LogLevel.Info);
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error configuring boot (may not be a bootable image)");
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
            Logger.Log($"PowerShell warning: {error}", Logger.LogLevel.Warning);
        }

        return output;
    }

    private static void ExecuteCommand(string command, string arguments)
    {
        var psi = new ProcessStartInfo
        {
            FileName = command,
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi);
        if (process == null)
        {
            throw new Exception($"Failed to start {command}");
        }

        process.WaitForExit();
        
        if (process.ExitCode != 0)
        {
            var error = process.StandardError.ReadToEnd();
            Logger.Log($"{command} warning: {error}", Logger.LogLevel.Warning);
        }
    }
}
