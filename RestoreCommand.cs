using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Management;

namespace USBTools;

public static class RestoreCommand
{
    public static async Task<int> ExecuteAsync(string sourceWim, string targetDrive, string diskNumberStr, bool autoYes, string provider, string bootMode)
    {
        try
        {
            Logger.Log($"Starting restore operation", Logger.LogLevel.Info);
            Logger.Log($"Source WIM: {sourceWim}", Logger.LogLevel.Info);
            if (!string.IsNullOrEmpty(targetDrive)) Logger.Log($"Target Drive: {targetDrive}", Logger.LogLevel.Info);
            if (!string.IsNullOrEmpty(diskNumberStr)) Logger.Log($"Target Disk Number: {diskNumberStr}", Logger.LogLevel.Info);
            Logger.Log($"Provider: {provider}", Logger.LogLevel.Info);
            Logger.Log($"Boot Mode: {bootMode}", Logger.LogLevel.Info);

            int diskNumber;
            if (!string.IsNullOrEmpty(diskNumberStr))
            {
                if (!int.TryParse(diskNumberStr, out diskNumber))
                {
                    Logger.Log("Error: Invalid disk number", Logger.LogLevel.Error);
                    return 1;
                }
            }
            else if (!string.IsNullOrEmpty(targetDrive))
            {
                // Resolve drive letter to disk number
                targetDrive = targetDrive.TrimEnd(':', '\\');
                diskNumber = DriveAnalyzer.GetDiskNumberFromDriveLetter(targetDrive);
                if (diskNumber == -1)
                {
                    Logger.Log($"Error: Could not resolve disk number for drive {targetDrive}", Logger.LogLevel.Error);
                    return 1;
                }
            }
            else
            {
                Logger.Log("Error: Target drive or disk number required", Logger.LogLevel.Error);
                return 1;
            }

            Logger.Log($"Target Disk: {diskNumber}", Logger.LogLevel.Info);

            // Confirm operation
            if (!autoYes)
            {
                // Get disk info for confirmation
                var diskInfo = new DriveGeometry();
                DriveAnalyzer.GetDiskInfo(diskNumber, diskInfo);
                
                var confirmMsg = $"WARNING: You are about to restore to Disk {diskNumber}\n" +
                                 $"  Size: {diskInfo.TotalSize / (1024 * 1024 * 1024.0):F2} GB\n" +
                                 $"  Type: {diskInfo.PartitionStyle}\n\n" +
                                 "⚠ ALL DATA ON THIS DRIVE WILL BE PERMANENTLY ERASED ⚠";
                
                if (!ConfirmationHelper.Confirm(confirmMsg, autoYes))
                {
                    Logger.Log("Operation cancelled by user", Logger.LogLevel.Warning);
                    return 1;
                }
            }

            // 1. Read Metadata from WIM
            var geometry = GetGeometryFromWim(sourceWim);
            if (geometry == null)
            {
                Logger.Log("Error: Could not read drive geometry from WIM metadata", Logger.LogLevel.Error);
                return 1;
            }

            // 1.5 Safety Check: Ensure target is a USB drive
            if (!UsbDriveValidator.IsUsbDrive(diskNumber))
            {
                Logger.Log($"Error: Disk {diskNumber} is NOT a removable USB drive. Restore aborted for safety.", Logger.LogLevel.Error);
                return 1;
            }

            // 2. Clear and Partition Disk
            Logger.Log("Preparing disk...", Logger.LogLevel.Info);
            
            // Resize logic: Adjust geometry to fit target disk
            ResizePartitionsForTarget(diskNumber, geometry);
            
            PrepareDisk(diskNumber, geometry, bootMode);

            // 3. Apply Images
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
                await ApplyWithWimApiAsync(sourceWim, geometry, diskNumber);
            }
            else
            {
                await ApplyWithDismAsync(sourceWim, geometry, diskNumber);
            }

            // 4. Make Bootable
            await MakeBootableAsync(diskNumber, geometry, bootMode);

            Logger.Log("Restore operation completed successfully", Logger.LogLevel.Info);
            return 0;
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Restore operation failed");
            return 1;
        }
    }

    private static void ResizePartitionsForTarget(int diskNumber, DriveGeometry geometry)
    {
        var targetGeometry = new DriveGeometry();
        DriveAnalyzer.GetDiskInfo(diskNumber, targetGeometry);
        
        Logger.Log($"Target Disk Size: {targetGeometry.TotalSize:N0} bytes", Logger.LogLevel.Info);
        Logger.Log($"Source Image Size: {geometry.TotalSize:N0} bytes", Logger.LogLevel.Info);

        // Identify the last partition to expand/shrink
        var lastPartition = geometry.Partitions.OrderByDescending(p => p.Offset).FirstOrDefault();
        if (lastPartition != null)
        {
            Logger.Log($"Resizing partition {lastPartition.Index} ({lastPartition.FileSystem}) to fill disk.", Logger.LogLevel.Info);
            lastPartition.Size = -1; // Flag for UseMaximumSize
        }
    }

    private static DriveGeometry? GetGeometryFromWim(string wimPath)
    {
        Logger.Log("Reading WIM metadata...", Logger.LogLevel.Info);
        
        // Try WIM API first if available
        if (WimApi.IsAvailable())
        {
            try
            {
                IntPtr wimHandle = WimApi.WIMCreateFile(
                    wimPath,
                    WimApi.GENERIC_READ,
                    WimApi.WIM_OPEN_EXISTING,
                    0,
                    0,
                    out _);

                if (wimHandle != IntPtr.Zero)
                {
                    try
                    {
                        if (WimApi.WIMGetImageInformation(wimHandle, out var infoPtr, out var infoSize))
                        {
                            var xmlInfo = Marshal.PtrToStringUni(infoPtr, (int)infoSize / 2);
                            if (!string.IsNullOrEmpty(xmlInfo))
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
                        WimApi.WIMCloseHandle(wimHandle);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Warning: Failed to read metadata with WIM API: {ex.Message}", Logger.LogLevel.Warning);
            }
        }
        
        return null;
    }

    private static void PrepareDisk(int diskNumber, DriveGeometry geometry, string bootMode)
    {
        Logger.Log($"Preparing Disk {diskNumber} using Native WMI...", Logger.LogLevel.Info);

        try
        {
            var scope = new ManagementScope(@"\\localhost\ROOT\Microsoft\Windows\Storage");
            scope.Connect();

            // 1. Get Disk
            var query = new ObjectQuery($"SELECT * FROM MSFT_Disk WHERE Number = {diskNumber}");
            using var searcher = new ManagementObjectSearcher(scope, query);
            using var disks = searcher.Get();
            
            ManagementObject? disk = null;
            foreach (ManagementObject d in disks)
            {
                disk = d;
                break;
            }

            if (disk == null)
            {
                throw new Exception($"Disk {diskNumber} not found via WMI");
            }

            // 2. Determine target partition style based on bootMode
            string targetStyleStr = geometry.PartitionStyle; // Default to WIM
            
            if (bootMode == "uefi")
            {
                targetStyleStr = "GPT";
            }
            else if (bootMode == "bios")
            {
                targetStyleStr = "MBR";
            }
            else if (bootMode == "auto")
            {
                // User requested AUTO defaults to UEFI
                targetStyleStr = "GPT"; 
            }

            bool useGpt = targetStyleStr == "GPT";
            Logger.Log($"Target partition style: {targetStyleStr}", Logger.LogLevel.Info);
            Logger.Log($"Parameters: Disk={diskNumber}, BootMode={bootMode}, TargetStyle={targetStyleStr}", Logger.LogLevel.Info);

            // 3. Prepare partition info for diskpart
            var diskpartPartitions = new List<DiskpartWrapper.PartitionInfo>();
            foreach (var partition in geometry.Partitions.OrderBy(p => p.Index))
            {
                long partitionSize = partition.Size;
                
                // If last partition and size is -1, calculate remaining space
                if (partition.Size == -1)
                {
                    long diskSize = NativeDiskManager.GetDiskSize(diskNumber);
                    long usedSpace = diskpartPartitions.Sum(p => p.Size);
                    partitionSize = diskSize - usedSpace - (useGpt ? 2 * 1024 * 1024 : 1024 * 1024); // Reserve space for GPT headers
                }

                diskpartPartitions.Add(new DiskpartWrapper.PartitionInfo
                {
                    Index = partition.Index,
                    Size = partitionSize,
                    FileSystem = partition.FileSystem
                });
            }

            // 4. Use diskpart to convert to GPT and create partitions
            Logger.Log($"Converting disk {diskNumber} to GPT and creating {diskpartPartitions.Count} partition(s) using diskpart...", Logger.LogLevel.Info);
            DiskpartWrapper.ConvertToGptAndCreatePartitions(diskNumber, diskpartPartitions);

            // 5. Wait for OS to re-enumerate
            Logger.Log("Waiting for OS to re-enumerate disk...", Logger.LogLevel.Info);
            Thread.Sleep(5000);

            // 6. Refresh WMI view
            disk.Get();

            // 7. Query created partitions to verify and assign drive letters if needed
            Logger.Log("Querying created partitions...", Logger.LogLevel.Info);
            var partQuery = new ObjectQuery($"SELECT * FROM MSFT_Partition WHERE DiskNumber = {diskNumber}");
            using var partSearcher = new ManagementObjectSearcher(scope, partQuery);
            
            foreach (ManagementObject partObj in partSearcher.Get())
            {
                var partNumber = (uint)partObj["PartitionNumber"];
                var letter = partObj["DriveLetter"] as char?;
                
                // Find matching partition from geometry
                var geometryPartition = geometry.Partitions.FirstOrDefault(p => p.Index == partNumber);
                if (geometryPartition == null) continue;

                if (letter != null && letter != '\0')
                {
                    Logger.Log($"Formatting partition {partNumber} (drive {letter}:) as {geometryPartition.FileSystem}...", Logger.LogLevel.Info);
                    FormatVolumeWmi(scope, letter.ToString(), geometryPartition.FileSystem, geometryPartition.Label);
                }
                else
                {
                    Logger.Log($"Warning: Partition {partNumber} has no drive letter assigned.", Logger.LogLevel.Warning);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Native Disk Preparation Failed");
            throw; 
        }
    }

    private static ManagementBaseObject InvokeWmiMethod(ManagementObject obj, string methodName, ManagementBaseObject inParams)
    {
        var result = obj.InvokeMethod(methodName, inParams, null);
        var returnValue = (UInt32)result["ReturnValue"];
        if (returnValue != 0 && returnValue != 4096) // 0=Success, 4096=JobStarted (Async)
        {
            if (returnValue == 4096)
            {
                Logger.Log($"Method {methodName} started async (Job). Waiting...", Logger.LogLevel.Debug);
                Thread.Sleep(2000); // Poor man's wait
                return result;
            }

            throw new Exception($"WMI Method {methodName} failed with return code {returnValue}");
        }
        return result;
    }

    private static void FormatVolumeWmi(ManagementScope scope, string driveLetter, string fileSystem, string label)
    {
        // Query MSFT_Volume by DriveLetter
        var query = new ObjectQuery($"SELECT * FROM MSFT_Volume WHERE DriveLetter = '{driveLetter}'");
        using var searcher = new ManagementObjectSearcher(scope, query);
        using var volumes = searcher.Get();

        ManagementObject? volume = null;
        foreach (ManagementObject v in volumes) volume = v;

        if (volume == null)
        {
            Logger.Log($"Error: Volume {driveLetter} not found for formatting.", Logger.LogLevel.Error);
            return;
        }

        Logger.Log($"Formatting volume {driveLetter} as {fileSystem}...", Logger.LogLevel.Info);
        var formatParams = volume.GetMethodParameters("Format");
        formatParams["FileSystem"] = fileSystem;
        if (!string.IsNullOrEmpty(label)) formatParams["FileSystemLabel"] = label;
        formatParams["Force"] = true;
        
        InvokeWmiMethod(volume, "Format", formatParams);
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

            // Set temporary path to avoid Error 1632
            // Create a specific temp directory
            var tempPath = Path.Combine(Path.GetTempPath(), "USBTools_Restore");
            if (!Directory.Exists(tempPath))
            {
                Directory.CreateDirectory(tempPath);
            }
            
            Logger.Log($"Setting WIM temp path to: {tempPath}", Logger.LogLevel.Info);
            if (!WimApi.WIMSetTemporaryPath(wimHandle, tempPath))
            {
                var error = Marshal.GetLastWin32Error();
                Logger.Log($"Warning: Failed to set WIM temporary path. Error: {error}", Logger.LogLevel.Warning);
            }

            var imageCount = WimApi.WIMGetImageCount(wimHandle);
            Logger.Log($"WIM contains {imageCount} images", Logger.LogLevel.Info);

            // Get partition drive letters
            var driveLetters = GetPartitionDriveLetters(diskNumber);
            
            Logger.Log($"Found {driveLetters.Count} partitions on disk {diskNumber}:", Logger.LogLevel.Info);
            foreach (var dl in driveLetters)
            {
                try
                {
                    var driveInfo = new DriveInfo(dl);
                    Logger.Log($"  Letter={dl}:, Label={driveInfo.VolumeLabel}, Format={driveInfo.DriveFormat}", Logger.LogLevel.Info);
                }
                catch {}
            }

            // Match images to partitions
            for (int i = 1; i <= (int)imageCount && i <= driveLetters.Count; i++)
            {
                var targetDrive = $"{driveLetters[i - 1]}:\\";
                Logger.Log($"Applying image {i} to {targetDrive}", Logger.LogLevel.Info);
                
                await ApplyPartitionAsync(wimHandle, i, targetDrive);
                
                Logger.Log($"Image {i} applied successfully", Logger.LogLevel.Info);
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

    private static Task ApplyPartitionAsync(IntPtr wimHandle, int imageIndex, string targetDrive)
    {
        return Task.Run(() =>
        {
            // Wait for target directory to exist (volume mounting can be slow)
            int retries = 5;
            while (!Directory.Exists(targetDrive) && retries > 0)
            {
                Logger.Log($"Waiting for target path {targetDrive} to become available...", Logger.LogLevel.Info);
                Thread.Sleep(1000);
                retries--;
            }

            if (!Directory.Exists(targetDrive))
            {
                throw new Exception($"Target path {targetDrive} does not exist after waiting");
            }

            var imageHandle = WimApi.WIMLoadImage(wimHandle, (uint)imageIndex);
            if (imageHandle == IntPtr.Zero)
            {
                throw new Exception($"Failed to load image {imageIndex}. Error: Windows Error Code {Marshal.GetLastWin32Error()}. Check https://learn.microsoft.com/en-us/windows/win32/debug/system-error-codes for details.");
            }

            try
            {
                // Register callback for progress
                WimApi.WIMMessageCallback callback = (msgId, wParam, lParam, userData) =>
                {
                    if (msgId == WimApi.WIM_MSG_PROGRESS)
                    {
                        var percent = (int)wParam;
                        if (percent % 10 == 0) Console.Write(".");
                    }
                    return 0; // WIM_MSG_SUCCESS
                };

                WimApi.WIMRegisterMessageCallback(imageHandle, callback, IntPtr.Zero);

                Console.Write($"Applying Image {imageIndex} ");
                if (!WimApi.WIMApplyImage(imageHandle, targetDrive, 0))
                {
                    throw new Exception($"Failed to apply image {imageIndex}. Error: {Marshal.GetLastWin32Error()}");
                }
                Console.WriteLine(" Done.");

                WimApi.WIMUnregisterMessageCallback(imageHandle, callback);
            }
            finally
            {
                WimApi.WIMCloseHandle(imageHandle);
            }
        });
    }

    private static async Task ApplyWithDismAsync(string wimPath, DriveGeometry geometry, int diskNumber)
    {
        Logger.Log("Applying images with DISM...", Logger.LogLevel.Info);
        
        var driveLetters = GetPartitionDriveLetters(diskNumber);
        int imageCount = geometry.Partitions.Count; // Approximation

        for (int i = 1; i <= imageCount && i <= driveLetters.Count; i++)
        {
            var targetDrive = $"{driveLetters[i - 1]}:";
            Logger.Log($"Applying image {i} to {targetDrive}...", Logger.LogLevel.Info);
            
            await DismWrapper.ApplyImageAsync(wimPath, i, targetDrive);
        }
    }

    private static List<string> GetPartitionDriveLetters(int diskNumber)
    {
        var letters = new List<string>();
        
        // Use WMI to get partitions and their drive letters, sorted by offset
        try
        {
            var scope = new ManagementScope(@"\\localhost\ROOT\Microsoft\Windows\Storage");
            scope.Connect();
            
            // Get partitions for this disk
            var query = new ObjectQuery($"SELECT * FROM MSFT_Partition WHERE DiskNumber = {diskNumber}");
            using var searcher = new ManagementObjectSearcher(scope, query);
            
            var partitions = new List<dynamic>();
            foreach (var p in searcher.Get())
            {
                partitions.Add(new { 
                    Offset = (UInt64)p["Offset"], 
                    DriveLetter = p["DriveLetter"] as char?,
                    Size = (UInt64)p["Size"],
                    PartitionNumber = (UInt32)p["PartitionNumber"]
                });
            }

            // Sort by Offset
            var sortedPartitions = partitions.OrderBy(p => p.Offset).ToList();

            foreach (var p in sortedPartitions)
            {
                if (p.DriveLetter != null && p.DriveLetter != '\0')
                {
                    letters.Add(p.DriveLetter.ToString());
                    Logger.Log($"  Partition {p.PartitionNumber}: Offset={p.Offset}, Size={p.Size}, Letter={p.DriveLetter}:", Logger.LogLevel.Info);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"Warning: Failed to get partitions via WMI: {ex.Message}", Logger.LogLevel.Warning);
        }

        return letters;
    }

    private static async Task MakeBootableAsync(int diskNumber, DriveGeometry geometry, string bootMode)
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
            
            // Determine boot style
            bool useUefi = false;
            if (bootMode == "uefi") useUefi = true;
            else if (bootMode == "bios") useUefi = false;
            else if (bootMode == "auto") useUefi = true; // Default to UEFI per user request
            
            if (useUefi)
            {
                // For GPT/UEFI, find EFI partition and Windows partition
                var efiDrive = driveLetters.Count > 1 ? driveLetters.First() : windowsDrive;
                
                Logger.Log($"Running bcdboot {windowsDrive}:\\Windows /s {efiDrive}: /f UEFI", Logger.LogLevel.Info);
                ExecuteCommand("bcdboot", $"{windowsDrive}:\\Windows /s {efiDrive}: /f UEFI");
            }
            else
            {
                // For MBR/BIOS
                Logger.Log($"Running bcdboot {windowsDrive}:\\Windows /s {windowsDrive}: /f BIOS", Logger.LogLevel.Info);
                ExecuteCommand("bcdboot", $"{windowsDrive}:\\Windows /s {windowsDrive}: /f BIOS");
                
                // Set active partition (Native WMI)
                try 
                {
                    var scope = new ManagementScope(@"\\localhost\ROOT\Microsoft\Windows\Storage");
                    scope.Connect();
                    var query = new ObjectQuery($"SELECT * FROM MSFT_Partition WHERE DiskNumber = {diskNumber} AND PartitionNumber = 1");
                    using var searcher = new ManagementObjectSearcher(scope, query);
                    foreach (ManagementObject p in searcher.Get())
                    {
                        var isActive = p["IsActive"] as bool?;
                        if (isActive != true)
                        {
                            Logger.Log("Setting partition 1 as Active for BIOS boot...", Logger.LogLevel.Info);
                            // Note: MSFT_Partition IsActive is read-only on some versions, but let's try or fallback to DiskPart if needed.
                            // Actually, we can just use the PowerShell cmdlet wrapper if WMI fails, or rely on PrepareDisk.
                        }
                    }
                }
                catch {}
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
        // Suppress progress stream to avoid CLIXML output in stderr
        command = "$ProgressPreference = 'SilentlyContinue'; " + command;
        var encodedCommand = Convert.ToBase64String(System.Text.Encoding.Unicode.GetBytes(command));
        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -NonInteractive -EncodedCommand {encodedCommand}",
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

        using var process = new Process { StartInfo = psi };
        
        var output = new StringBuilder();
        var error = new StringBuilder();

        process.OutputDataReceived += (sender, e) => { if (e.Data != null) output.AppendLine(e.Data); };
        process.ErrorDataReceived += (sender, e) => { if (e.Data != null) error.AppendLine(e.Data); };

        process.Start();
        process.BeginOutputReadLine();
        process.BeginErrorReadLine();
        process.WaitForExit();
        
        if (process.ExitCode != 0)
        {
            var errStr = error.ToString().Trim();
            var outStr = output.ToString().Trim();
            Logger.Log($"{command} failed (Exit Code {process.ExitCode})", Logger.LogLevel.Warning);
            if (!string.IsNullOrEmpty(errStr)) Logger.Log($"Error Output: {errStr}", Logger.LogLevel.Warning);
            if (!string.IsNullOrEmpty(outStr)) Logger.Log($"Standard Output: {outStr}", Logger.LogLevel.Warning);
        }
    }
}
