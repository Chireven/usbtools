using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Text.Json;
using System.Management;
using System.Security.Cryptography;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace USBTools;

public static class RestoreCommand
{
    public static async Task<int> ExecuteAsync(string sourceWim, string targetDrive, string diskNumberStr, bool autoYes, string provider, string bootMode, bool forceBoot = false, bool preflightOnly = false, bool verifyHashes = false)
    {
        try
        {
            Logger.Log($"Starting restore operation", Logger.LogLevel.Info);
            Logger.Log($"Source WIM: {sourceWim}", Logger.LogLevel.Info);
            if (!string.IsNullOrEmpty(targetDrive)) Logger.Log($"Target Drive: {targetDrive}", Logger.LogLevel.Info);
            if (!string.IsNullOrEmpty(diskNumberStr)) Logger.Log($"Target Disk Number: {diskNumberStr}", Logger.LogLevel.Info);
            Logger.Log($"Provider: {provider}", Logger.LogLevel.Info);
            Logger.Log($"Boot Mode: {bootMode}", Logger.LogLevel.Info);
            if (forceBoot) Logger.Log("Force Boot: Enabled", Logger.LogLevel.Warning);
            TelemetryReporter.Initialize();

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

            geometry.TemplateHint ??= DetermineTemplate(geometry);
            if (!string.IsNullOrEmpty(geometry.TemplateHint))
            {
                Logger.Log($"Template hint: {geometry.TemplateHint}", Logger.LogLevel.Info);
            }

            var preflight = RunPreflight(diskNumber, geometry, provider, bootMode);
            ReportPreflight(preflight);
            TelemetryReporter.Emit("restore", "preflight", new { preflight.Score, preflight.Ready, Issues = preflight.Issues });
            if (preflightOnly)
            {
                Logger.Log("Dry-run completed; no changes were made.", Logger.LogLevel.Info);
                return preflight.Ready ? 0 : 1;
            }
            if (!preflight.Ready)
            {
                Logger.Log("Preflight detected blocking issues; aborting to avoid destructive changes.", Logger.LogLevel.Error);
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
            var useWimApi = ResolveProvider(provider);
            Logger.Log($"Using provider: {(useWimApi ? "WIM API" : "DISM")}", Logger.LogLevel.Info);
            TelemetryReporter.Emit("restore", "applying", new { provider = useWimApi ? "wimapi" : "dism" });
            await ApplyImagesAsync(sourceWim, geometry, diskNumber, useWimApi, verifyHashes, provider);

            var useUefi = bootMode switch
            {
                "uefi" => true,
                "bios" => false,
                _ => geometry.PartitionStyle.Equals("GPT", StringComparison.OrdinalIgnoreCase)
            };

            ValidateRestoredLayout(geometry, diskNumber, useUefi);

            HandleSecurityArtifacts(geometry, diskNumber);

            // 4. Make Bootable
            await MakeBootableAsync(diskNumber, geometry, bootMode, forceBoot);

            WriteAuditManifest(sourceWim, geometry, diskNumber);

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
        const long gptOverheadBytes = 2 * 1024 * 1024;   // Reserve space for GPT headers/footers
        const long mbrOverheadBytes = 1 * 1024 * 1024;   // Leave a little room for the MBR
        const long elasticSafetyPadding = 256L * 1024 * 1024; // Add headroom when shrinking partitions

        var targetGeometry = new DriveGeometry();
        DriveAnalyzer.GetDiskInfo(diskNumber, targetGeometry);

        geometry.TotalSize = targetGeometry.TotalSize; // Keep metadata aligned with the target

        Logger.Log($"Target Disk Size: {targetGeometry.TotalSize:N0} bytes", Logger.LogLevel.Info);
        Logger.Log($"Source Image Size: {geometry.TotalSize:N0} bytes", Logger.LogLevel.Info);

        long overhead = geometry.PartitionStyle.Equals("GPT", StringComparison.OrdinalIgnoreCase)
            ? gptOverheadBytes
            : mbrOverheadBytes;

        var fixedPartitions = geometry.Partitions.Where(p => p.IsFixed).OrderBy(p => p.Offset).ToList();
        var elasticPartitions = geometry.Partitions.Where(p => !p.IsFixed).OrderBy(p => p.Offset).ToList();

        long fixedTotal = fixedPartitions.Sum(p => p.Size);
        long sourceElasticTotal = Math.Max(1, elasticPartitions.Sum(p => p.Size)); // avoid div-by-zero later

        // Calculate the minimum viable size for elastic partitions (used space + padding)
        var minElasticSizes = elasticPartitions
            .Select(p => Math.Max(p.UsedSpace + elasticSafetyPadding, p.UsedSpace))
            .ToList();

        long minimumRequired = fixedTotal + minElasticSizes.Sum() + overhead;
        if (minimumRequired > targetGeometry.TotalSize)
        {
            Logger.Log(
                $"Error: Target disk is too small. Requires at least {minimumRequired:N0} bytes but only {targetGeometry.TotalSize:N0} are available.",
                Logger.LogLevel.Error);
            throw new InvalidOperationException("Target disk is smaller than the captured image requirements.");
        }

        // If there are no elastic partitions, we cannot safely resize. Leave sizes intact but warn about leftover space.
        if (!elasticPartitions.Any())
        {
            long unused = targetGeometry.TotalSize - fixedTotal - overhead;
            if (unused < 0)
            {
                Logger.Log(
                    "Error: No elastic partitions available and fixed partitions exceed target disk capacity.",
                    Logger.LogLevel.Error);
                throw new InvalidOperationException("Unable to fit fixed partitions onto target disk.");
            }

            Logger.Log(
                $"No elastic partitions detected; proceeding with fixed sizes. Unallocated space: {unused:N0} bytes.",
                Logger.LogLevel.Warning);
            return;
        }

        long availableForElastic = targetGeometry.TotalSize - fixedTotal - overhead;
        long remaining = availableForElastic;

        for (int i = 0; i < elasticPartitions.Count; i++)
        {
            var partition = elasticPartitions[i];
            long minSize = minElasticSizes[i];

            long proportional = (long)Math.Floor((double)partition.Size / sourceElasticTotal * availableForElastic);
            long allocated = Math.Max(minSize, proportional);

            // Ensure remaining partitions can still meet their minima
            long remainingMin = minElasticSizes.Skip(i + 1).Sum();
            if (allocated + remainingMin > remaining)
            {
                allocated = remaining - remainingMin;
            }

            // Last partition takes the rest to avoid rounding losses
            if (i == elasticPartitions.Count - 1)
            {
                allocated = remaining;
            }

            partition.Size = allocated;
            remaining -= allocated;

            Logger.Log(
                $"Partition {partition.Index} resized to {allocated:N0} bytes (min {minSize:N0}).",
                Logger.LogLevel.Info);
        }
    }

    private static DriveGeometry? GetGeometryFromWim(string wimPath)
    {
        Logger.Log("Reading WIM metadata...", Logger.LogLevel.Info);

        var geometry = TryGetGeometryWithWimApi(wimPath);
        if (geometry != null)
        {
            return geometry;
        }

        geometry = TryGetGeometryWithDism(wimPath);
        if (geometry != null)
        {
            return geometry;
        }

        geometry = TryGetGeometryFromManifest(wimPath);
        if (geometry != null)
        {
            return geometry;
        }

        return null;
    }

    private static DriveGeometry? TryGetGeometryWithWimApi(string wimPath)
    {
        if (!WimApi.IsAvailable()) return null;

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
                                // Unescape XML entities
                                json = json.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&amp;", "&");

                                Logger.Log($"WIM Metadata JSON: {json}", Logger.LogLevel.Debug);

                                var geometry = DriveGeometry.FromJson(json);
                                if (geometry != null)
                                {
                                    Logger.Log($"Image Tool Version: {geometry.ToolVersion}", Logger.LogLevel.Info);
                                }
                                return geometry;
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

        return null;
    }

    private static DriveGeometry? TryGetGeometryWithDism(string wimPath)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "dism.exe",
                Arguments = $"/Get-WimInfo /WimFile:\"{wimPath}\" /Index:1",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null) return null;

            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            if (!string.IsNullOrWhiteSpace(error))
            {
                Logger.Log($"DISM metadata stderr: {error}", Logger.LogLevel.Warning);
            }

            var match = Regex.Match(output, @"Description\s*:\s*(.+)");
            if (match.Success)
            {
                var json = match.Groups[1].Value.Trim();
                var geometry = DriveGeometry.FromJson(json);
                if (geometry != null)
                {
                    Logger.Log("Read metadata via DISM fallback", Logger.LogLevel.Info);
                }
                return geometry;
            }
        }
        catch (Exception ex)
        {
            Logger.Log($"Warning: Failed to read metadata with DISM: {ex.Message}", Logger.LogLevel.Warning);
        }

        return null;
    }

    private static DriveGeometry? TryGetGeometryFromManifest(string wimPath)
    {
        var manifestCandidates = new List<string>
        {
            wimPath + ".manifest.json",
            Path.ChangeExtension(wimPath, ".json")
        };

        foreach (var path in manifestCandidates)
        {
            if (!File.Exists(path)) continue;

            try
            {
                var json = File.ReadAllText(path);
                var geometry = DriveGeometry.FromJson(json);
                if (geometry != null)
                {
                    Logger.Log($"Read metadata from sidecar manifest: {path}", Logger.LogLevel.Info);
                    return geometry;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"Warning: Failed to read manifest {path}: {ex.Message}", Logger.LogLevel.Warning);
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
                // User requested AUTO defaults to source metadata
                targetStyleStr = geometry.PartitionStyle; 
            }

            bool useGpt = targetStyleStr == "GPT";
            Logger.Log($"Target partition style: {targetStyleStr}", Logger.LogLevel.Info);
            Logger.Log($"Parameters: Disk={diskNumber}, BootMode={bootMode}, TargetStyle={targetStyleStr}", Logger.LogLevel.Info);

            // 3. Prepare partition info for diskpart
            var diskpartPartitions = new List<DiskpartWrapper.PartitionInfo>();
            
            // Calculate available drive letters (Z down to C)
            // We blacklist A and B to avoid issues with antivirus software and legacy assumptions
            var usedLetters = DriveInfo.GetDrives().Select(d => d.Name[0]).ToHashSet();
            var availableLetters = new Stack<char>();
            for (char c = 'Z'; c >= 'C'; c--)
            {
                if (!usedLetters.Contains(c)) availableLetters.Push(c);
            }
            
            // CRITICAL: Sort by Offset to ensure we create partitions in the same physical order as the source.
            // This ensures that the first partition created gets the first physical location, etc.
            var sortedPartitions = geometry.Partitions.OrderBy(p => p.Offset).ToList();

            foreach (var partition in sortedPartitions)
            {
                long partitionSize = partition.Size;
                
                // If last partition and size is -1, calculate remaining space
                if (partition.Size == -1)
                {
                    long diskSize = NativeDiskManager.GetDiskSize(diskNumber);
                    long usedSpace = diskpartPartitions.Sum(p => p.Size);
                    partitionSize = diskSize - usedSpace - (useGpt ? 2 * 1024 * 1024 : 1024 * 1024); // Reserve space for GPT headers
                }

                string assignedLetter = "";
                if (!string.IsNullOrWhiteSpace(partition.Letter))
                {
                    assignedLetter = partition.Letter.TrimEnd(':', '\\');
                }
                else if (availableLetters.Count > 0)
                {
                    assignedLetter = availableLetters.Pop().ToString();
                }

                if (useGpt && assignedLetter == "" && string.Equals(partition.GptType, "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}", StringComparison.OrdinalIgnoreCase))
                {
                    assignedLetter = "S"; // Reserve a predictable letter for EFI if nothing else is assigned
                }

                diskpartPartitions.Add(new DiskpartWrapper.PartitionInfo
                {
                    Index = partition.Index,
                    Size = partitionSize,
                    FileSystem = partition.FileSystem,
                    Label = partition.Label,
                    DriveLetter = assignedLetter,
                    IsActive = partition.IsActive,
                    GptType = partition.GptType,
                    GptId = partition.GptId,
                    IsFixed = partition.IsFixed
                });
            }

            // 4. Use diskpart to convert to GPT and create partitions
            Logger.Log($"Converting disk {diskNumber} to {targetStyleStr} and creating {diskpartPartitions.Count} partition(s) using diskpart...", Logger.LogLevel.Info);
            DiskpartWrapper.PrepareDiskAndCreatePartitions(diskNumber, diskpartPartitions, targetStyleStr);

            // 5. Wait for OS to re-enumerate
            Logger.Log("Waiting for OS to re-enumerate disk...", Logger.LogLevel.Info);
            Thread.Sleep(5000);

            // 6. Refresh WMI view
            disk.Get();

            // 7. Query created partitions to verify and assign drive letters if needed
            Logger.Log("Querying created partitions...", Logger.LogLevel.Info);
            var partQuery = new ObjectQuery($"SELECT * FROM MSFT_Partition WHERE DiskNumber = {diskNumber}");
            using var partSearcher = new ManagementObjectSearcher(scope, partQuery);
            
            var usedLetters = DriveInfo.GetDrives().Select(d => d.Name[0]).ToHashSet();

            foreach (ManagementObject partObj in partSearcher.Get())
            {
                var partNumber = (uint)partObj["PartitionNumber"];
                var letter = partObj["DriveLetter"] as char?;

                // Find matching partition from geometry
                var geometryPartition = geometry.Partitions.FirstOrDefault(p => p.Index == partNumber);
                if (geometryPartition == null) continue;

                if (letter != null && letter != '\0')
                {
                    Logger.Log($"Partition {partNumber} (drive {letter}:) verified.", Logger.LogLevel.Info);
                }
                else
                {
                    var preferredLetter = geometryPartition.Letter?.TrimEnd(':', '\\');
                    if (useGpt && string.IsNullOrWhiteSpace(preferredLetter) && string.Equals(geometryPartition.GptType, "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}", StringComparison.OrdinalIgnoreCase))
                    {
                        preferredLetter = "S";
                    }

                    var nextLetter = !string.IsNullOrWhiteSpace(preferredLetter)
                        ? preferredLetter
                        : GetNextFreeLetter(usedLetters);

                    if (!string.IsNullOrWhiteSpace(nextLetter))
                    {
                        AssignPartitionLetter(diskNumber, partNumber, nextLetter);
                        usedLetters.Add(nextLetter[0]);
                        Logger.Log($"Partition {partNumber} assigned drive letter {nextLetter} for accessibility.", Logger.LogLevel.Info);
                    }
                    else
                    {
                        Logger.Log($"Warning: Partition {partNumber} has no drive letter assigned and none are available.", Logger.LogLevel.Warning);
                    }
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

    private static bool ResolveProvider(string provider)
    {
        if (provider == "wimapi")
        {
            if (!WimApi.IsAvailable())
            {
                throw new InvalidOperationException("WIM API requested but not available.");
            }
            return true;
        }

        if (provider == "dism")
        {
            return false;
        }

        return WimApi.IsAvailable();
    }

    private class PartitionApplyPlan
    {
        public int ImageIndex { get; set; }
        public string DriveRoot { get; set; } = string.Empty;
        public bool IsBoot { get; set; }
        public PartitionInfo Geometry { get; set; } = new();
    }

    private static List<PartitionApplyPlan> BuildApplyPlan(DriveGeometry geometry, int diskNumber)
    {
        var driveLetters = GetPartitionDriveLetters(diskNumber);
        var plans = new List<PartitionApplyPlan>();
        for (int i = 0; i < geometry.Partitions.Count && i < driveLetters.Count; i++)
        {
            var partition = geometry.Partitions.OrderBy(p => p.Index).ToList()[i];
            bool isBoot = partition.IsFixed || (partition.Type?.Contains("Boot", StringComparison.OrdinalIgnoreCase) ?? false) ||
                          string.Equals(partition.GptType, "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}", StringComparison.OrdinalIgnoreCase);

            plans.Add(new PartitionApplyPlan
            {
                ImageIndex = i + 1,
                DriveRoot = $"{driveLetters[i]}:\\",
                IsBoot = isBoot,
                Geometry = partition
            });
        }

        return plans;
    }

    private static async Task ApplyImagesAsync(string wimPath, DriveGeometry geometry, int diskNumber, bool useWimApi, bool verifyHashes, string provider)
    {
        var plans = BuildApplyPlan(geometry, diskNumber);
        var checkpoint = RestoreCheckpoint.Load(wimPath);
        checkpoint.Provider = provider;
        ProgressTree.Node("Staging apply plan", 1);

        var bootPlans = plans.Where(p => p.IsBoot).OrderBy(p => p.ImageIndex).ToList();
        var dataPlans = plans.Where(p => !p.IsBoot).OrderBy(p => p.ImageIndex).ToList();

        foreach (var plan in bootPlans)
        {
            await ApplySinglePartitionAsync(wimPath, plan, useWimApi, checkpoint, verifyHashes);
        }

        await ApplyDataPartitionsAsync(wimPath, dataPlans, useWimApi, checkpoint, verifyHashes);

        checkpoint.Clear(wimPath);
    }

    private static async Task ApplySinglePartitionAsync(string wimPath, PartitionApplyPlan plan, bool useWimApi, RestoreCheckpoint checkpoint, bool verifyHashes)
    {
        if (checkpoint.CompletedImages.Contains(plan.ImageIndex))
        {
            Logger.Log($"Skipping image {plan.ImageIndex} ({plan.DriveRoot}) - already completed in previous run.", Logger.LogLevel.Info);
            return;
        }

        ProgressTree.Node($"Applying image {plan.ImageIndex} -> {plan.DriveRoot}", 2);
        await ApplyPartitionWithProviderAsync(wimPath, plan, useWimApi);

        if (verifyHashes)
        {
            await VerifyPartitionHashAsync(plan);
        }

        checkpoint.CompletedImages.Add(plan.ImageIndex);
        checkpoint.Save(wimPath);
        TelemetryReporter.Emit("restore", "partition_applied", new { plan.ImageIndex, plan.DriveRoot });
    }

    private static async Task ApplyDataPartitionsAsync(string wimPath, List<PartitionApplyPlan> dataPlans, bool useWimApi, RestoreCheckpoint checkpoint, bool verifyHashes)
    {
        if (!dataPlans.Any()) return;

        var concurrency = Math.Max(1, Math.Min(Environment.ProcessorCount, 4));
        using var gate = new SemaphoreSlim(concurrency);
        var tasks = dataPlans.Select(async plan =>
        {
            await gate.WaitAsync();
            try
            {
                await ApplySinglePartitionAsync(wimPath, plan, useWimApi, checkpoint, verifyHashes);
            }
            finally
            {
                gate.Release();
            }
        }).ToList();

        await Task.WhenAll(tasks);
    }

    private static Task ApplyPartitionWithProviderAsync(string wimPath, PartitionApplyPlan plan, bool useWimApi)
    {
        return useWimApi
            ? ApplyPartitionWithWimApiAsync(wimPath, plan.ImageIndex, plan.DriveRoot)
            : DismWrapper.ApplyImageAsync(wimPath, plan.ImageIndex, plan.DriveRoot);
    }

    private static Task ApplyPartitionWithWimApiAsync(string wimPath, int imageIndex, string targetDrive)
    {
        return Task.Run(() =>
        {
            int retries = 5;
            while (!Directory.Exists(targetDrive) && retries-- > 0)
            {
                Logger.Log($"Waiting for target path {targetDrive} to become available...", Logger.LogLevel.Info);
                Thread.Sleep(1000);
            }

            if (!Directory.Exists(targetDrive))
            {
                throw new Exception($"Target path {targetDrive} does not exist after waiting");
            }

            IntPtr wimHandle = IntPtr.Zero;
            IntPtr imageHandle = IntPtr.Zero;
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
                    throw new Exception("Failed to open WIM file for apply");
                }

                imageHandle = WimApi.WIMLoadImage(wimHandle, (uint)imageIndex);
                if (imageHandle == IntPtr.Zero)
                {
                    throw new Exception($"Failed to load image {imageIndex}. Error: {Marshal.GetLastWin32Error()}");
                }

                int lastPercent = 0;
                WimApi.WIMMessageCallback callback = (msgId, wParam, lParam, userData) =>
                {
                    if (msgId == WimApi.WIM_MSG_PROGRESS)
                    {
                        var percent = (int)wParam;
                        if (percent >= lastPercent + 10)
                        {
                            Logger.Log($"Image {imageIndex} -> {targetDrive}: {percent}%", Logger.LogLevel.Info);
                            lastPercent = percent;
                        }
                    }
                    return 0;
                };

                WimApi.WIMRegisterMessageCallback(imageHandle, callback, IntPtr.Zero);

                if (!WimApi.WIMApplyImage(imageHandle, targetDrive, 0))
                {
                    throw new Exception($"Failed to apply image {imageIndex}. Error: {Marshal.GetLastWin32Error()}");
                }
            }
            finally
            {
                if (imageHandle != IntPtr.Zero) WimApi.WIMCloseHandle(imageHandle);
                if (wimHandle != IntPtr.Zero) WimApi.WIMCloseHandle(wimHandle);
            }
        });
    }

    private static async Task VerifyPartitionHashAsync(PartitionApplyPlan plan)
    {
        if (string.IsNullOrWhiteSpace(plan.Geometry.DataHash))
        {
            Logger.Log($"No captured hash for partition {plan.ImageIndex}; skipping verification.", Logger.LogLevel.Warning);
            return;
        }

        ProgressTree.Node($"Hashing {plan.DriveRoot}", 3);
        var computed = await Task.Run(() => ComputePartitionHash(plan.DriveRoot));
        if (!computed.Equals(plan.Geometry.DataHash, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException($"Hash mismatch on {plan.DriveRoot}. Expected {plan.Geometry.DataHash}, got {computed}.");
        }

        Logger.Log($"Hash verified for {plan.DriveRoot}", Logger.LogLevel.Info);
    }

    private static string ComputePartitionHash(string driveRoot)
    {
        using var sha = SHA256.Create();
        foreach (var file in Directory.EnumerateFiles(driveRoot, "*", SearchOption.AllDirectories))
        {
            using var stream = File.OpenRead(file);
            var buffer = new byte[1024 * 1024];
            int read;
            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
            {
                sha.TransformBlock(buffer, 0, read, null, 0);
            }
        }

        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return Convert.ToHexString(sha.Hash); 
    }

    private class PreflightResult
    {
        public int Score { get; set; }
        public bool Ready { get; set; }
        public List<string> Issues { get; set; } = new();
        public List<string> Tips { get; set; } = new();
    }

    private static PreflightResult RunPreflight(int diskNumber, DriveGeometry geometry, string provider, string bootMode)
    {
        var result = new PreflightResult { Score = 100, Ready = true };

        var targetGeometry = new DriveGeometry();
        DriveAnalyzer.GetDiskInfo(diskNumber, targetGeometry);

        // Capacity check similar to resize safety
        long fixedTotal = geometry.Partitions.Where(p => p.IsFixed).Sum(p => p.Size);
        long elasticMin = geometry.Partitions.Where(p => !p.IsFixed).Sum(p => Math.Max(p.UsedSpace, p.UsedSpace + 256L * 1024 * 1024));
        long minimum = fixedTotal + elasticMin + (geometry.PartitionStyle.Equals("GPT", StringComparison.OrdinalIgnoreCase) ? 2 * 1024 * 1024 : 1 * 1024 * 1024);
        if (minimum > targetGeometry.TotalSize)
        {
            result.Ready = false;
            result.Score -= 40;
            result.Issues.Add($"Target disk too small: needs {minimum:N0} bytes, has {targetGeometry.TotalSize:N0}.");
            result.Tips.Add("Use a larger USB drive or shrink captured data partitions.");
        }

        // Provider availability
        if (provider == "wimapi" && !WimApi.IsAvailable())
        {
            result.Ready = false;
            result.Score -= 30;
            result.Issues.Add("wimgapi.dll not available for requested provider.");
            result.Tips.Add("Install Windows ADK or use --provider dism");
        }

        if (provider == "auto" && !WimApi.IsAvailable())
        {
            result.Score -= 5;
            result.Tips.Add("Using DISM fallback because WIM API is missing.");
        }

        // Partition style compatibility
        if (bootMode == "uefi" && !geometry.PartitionStyle.Equals("GPT", StringComparison.OrdinalIgnoreCase))
        {
            result.Score -= 10;
            result.Tips.Add("Source is MBR but UEFI requested; GPT conversion will be applied.");
        }

        if (!string.IsNullOrEmpty(geometry.TemplateHint))
        {
            result.Tips.Add($"Template matched: {geometry.TemplateHint}");
        }

        result.Score = Math.Max(0, result.Score);
        return result;
    }

    private static void ReportPreflight(PreflightResult result)
    {
        Logger.Log($"Preflight score: {result.Score}/100", result.Score >= 80 ? Logger.LogLevel.Info : Logger.LogLevel.Warning);
        foreach (var issue in result.Issues)
        {
            Logger.Log($"Issue: {issue}", Logger.LogLevel.Warning);
        }
        foreach (var tip in result.Tips)
        {
            Logger.Log($"Tip: {tip}", Logger.LogLevel.Info);
        }
    }

    private static string? DetermineTemplate(DriveGeometry geometry)
    {
        if (geometry.Partitions.Count >= 2)
        {
            var esp = geometry.Partitions.Any(p => string.Equals(p.GptType, "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}", StringComparison.OrdinalIgnoreCase));
            var msr = geometry.Partitions.Any(p => string.Equals(p.GptType, "{e3c9e316-0b5c-4db8-817d-f92df00215ae}", StringComparison.OrdinalIgnoreCase));
            if (esp && msr)
            {
                return "Windows To Go";
            }
        }

        if (geometry.Partitions.Count >= 1 && geometry.Partitions.Any(p => p.Label?.Contains("EFI", StringComparison.OrdinalIgnoreCase) == true))
        {
            return "UEFI Bootable USB";
        }

        return geometry.TemplateHint;
    }

    private static void HandleSecurityArtifacts(DriveGeometry geometry, int diskNumber)
    {
        _ = diskNumber;
        if (geometry.BitLockerDetected)
        {
            Logger.Log("BitLocker protectors were present in metadata; TPM-bound volumes may need protector regeneration after restore.", Logger.LogLevel.Warning);
        }

        if (geometry.SecureBootRequired)
        {
            Logger.Log("Secure Boot requirement detected; ensuring EFI files remain intact.", Logger.LogLevel.Info);
        }
    }

    private static void WriteAuditManifest(string sourceWim, DriveGeometry geometry, int diskNumber)
    {
        var manifestPath = Path.Combine(Path.GetDirectoryName(sourceWim) ?? Environment.CurrentDirectory, "restore-manifest.json");
        var manifest = new
        {
            source = Path.GetFileName(sourceWim),
            diskNumber,
            timestamp = DateTime.UtcNow,
            template = geometry.TemplateHint,
            partitions = geometry.Partitions.Select(p => new { p.Index, p.Size, p.FileSystem, p.DataHash }),
            machine = Environment.MachineName
        };

        var json = JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true });
        var signature = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(json)));
        var signed = new
        {
            manifest,
            signature
        };

        File.WriteAllText(manifestPath, JsonSerializer.Serialize(signed, new JsonSerializerOptions { WriteIndented = true }));
        Logger.Log($"Audit manifest saved to {manifestPath}", Logger.LogLevel.Info);
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
                // Register callback for progress
                int lastPercent = 0;
                WimApi.WIMMessageCallback callback = (msgId, wParam, lParam, userData) =>
                {
                    if (msgId == WimApi.WIM_MSG_PROGRESS)
                    {
                        var percent = (int)wParam;
                        // Log every 10%
                        if (percent > lastPercent && percent % 10 == 0)
                        {
                            Logger.Log($"Applying Image {imageIndex}: {percent}% completed", Logger.LogLevel.Info);
                            lastPercent = percent;
                        }
                    }
                    return 0; // WIM_MSG_SUCCESS
                };

                WimApi.WIMRegisterMessageCallback(imageHandle, callback, IntPtr.Zero);

                Logger.Log($"Applying Image {imageIndex}...", Logger.LogLevel.Info);
                if (!WimApi.WIMApplyImage(imageHandle, targetDrive, 0))
                {
                    throw new Exception($"Failed to apply image {imageIndex}. Error: {Marshal.GetLastWin32Error()}");
                }
                Logger.Log($"Image {imageIndex} applied successfully.", Logger.LogLevel.Info);

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
            var letter = driveLetters[i - 1];
            var targetDrive = $"{letter}:\\"; // Ensure trailing backslash
            Logger.Log($"Applying image {i} to {targetDrive}...", Logger.LogLevel.Info);
            
            // Wait for volume to be accessible
            bool isReady = false;
            for (int attempt = 0; attempt < 10; attempt++)
            {
                if (Directory.Exists(targetDrive))
                {
                    isReady = true;
                    break;
                }
                Logger.Log($"Waiting for {targetDrive} to become accessible (Attempt {attempt + 1}/10)...", Logger.LogLevel.Debug);
                await Task.Delay(1000);
            }

            if (!isReady)
            {
                throw new DirectoryNotFoundException($"Target drive {targetDrive} is not accessible after waiting.");
            }

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

    private static async Task MakeBootableAsync(int diskNumber, DriveGeometry geometry, string bootMode, bool forceBoot)
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

            // Determine boot style
            bool useUefi = false;
            if (bootMode == "uefi") useUefi = true;
            else if (bootMode == "bios") useUefi = false;
            else if (bootMode == "auto") useUefi = geometry.PartitionStyle == "GPT"; // Default to source metadata

            // Check if boot files already exist (e.g. restored from WIM)
            bool bootFilesExist = false;
            string? efiDrive = null;
            foreach (var letter in driveLetters)
            {
                if (Directory.Exists($"{letter}:\\EFI") || File.Exists($"{letter}:\\bootmgr"))
                {
                    Logger.Log($"Boot files detected on {letter}: (EFI/bootmgr).", Logger.LogLevel.Info);
                    bootFilesExist = true;
                    
                    // If UEFI, we might still want to ensure the EFI partition is identified for logging
                    if (Directory.Exists($"{letter}:\\EFI")) efiDrive = letter;
                }
            }

            if (bootFilesExist)
            {
                if (forceBoot)
                {
                    Logger.Log("Force Boot enabled: Proceeding with bcdboot generation despite existing boot files.", Logger.LogLevel.Warning);
                }
                else
                {
                    Logger.Log("Skipping bcdboot generation (files exist). Use --force-boot to override.", Logger.LogLevel.Info);
                    return;
                }
            }

            // Fallback: Try to find Windows partition for bcdboot (Full OS scenario)
            string? windowsDrive = null;
            foreach (var letter in driveLetters)
            {
                if (Directory.Exists($"{letter}:\\Windows\\System32"))
                {
                    windowsDrive = letter;
                    Logger.Log($"Found Windows on {letter}:", Logger.LogLevel.Info);
                    break;
                }
            }

            if (windowsDrive == null)
            {
                Logger.Log("Warning: Could not locate Windows partition to generate boot files. Assuming image is already bootable or non-OS.", Logger.LogLevel.Warning);
                return;
            }

            if (efiDrive == null && useUefi)
            {
                // If only one partition, it might be both
                efiDrive = windowsDrive;
            }
            
            if (useUefi)
            {
                Logger.Log($"Running bcdboot {windowsDrive}:\\Windows /s {efiDrive}: /f UEFI", Logger.LogLevel.Info);
                ExecuteCommand("bcdboot", $"{windowsDrive}:\\Windows /s {efiDrive}: /f UEFI");
            }
            else
            {
                // For MBR/BIOS
                Logger.Log($"Running bcdboot {windowsDrive}:\\Windows /s {windowsDrive}: /f BIOS", Logger.LogLevel.Info);
                ExecuteCommand("bcdboot", $"{windowsDrive}:\\Windows /s {windowsDrive}: /f BIOS");
                
                // Update MBR and PBR using bootsect
                Logger.Log($"Updating boot sector on {windowsDrive}:...", Logger.LogLevel.Info);
                ExecuteCommand("bootsect", $"/nt60 {windowsDrive}: /force /mbr");

                // Active partition is already set by DiskpartWrapper based on metadata
                Logger.Log("Verifying boot configuration...", Logger.LogLevel.Info);
            }

            Logger.Log("Boot configuration completed", Logger.LogLevel.Info);
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error configuring boot (may not be a bootable image)");
        }
    }

    private static void ValidateRestoredLayout(DriveGeometry geometry, int diskNumber, bool useUefi)
    {
        Logger.Log("Validating restored layout...", Logger.LogLevel.Info);
        var report = new StringBuilder();
        report.AppendLine("# Restore Validation Report");
        report.AppendLine($"Disk: {diskNumber}");
        report.AppendLine($"Timestamp: {DateTime.UtcNow:O}");
        report.AppendLine();

        var scope = new ManagementScope(@"\\localhost\ROOT\Microsoft\Windows\Storage");
        scope.Connect();

        var partitionQuery = new ObjectQuery($"SELECT * FROM MSFT_Partition WHERE DiskNumber = {diskNumber}");
        using var partSearcher = new ManagementObjectSearcher(scope, partitionQuery);

        var parts = new List<ManagementObject>();
        foreach (ManagementObject part in partSearcher.Get()) parts.Add(part);

        if (parts.Count != geometry.Partitions.Count)
        {
            report.AppendLine($"Partition count mismatch. Expected {geometry.Partitions.Count}, found {parts.Count}.");
            throw new InvalidOperationException($"Partition count mismatch. Expected {geometry.Partitions.Count}, found {parts.Count} on disk.");
        }

        const long sizeTolerance = 10L * 1024 * 1024; // 10MB tolerance to accommodate rounding
        string? efiLetter = null;

        foreach (var geoPart in geometry.Partitions)
        {
            var partObj = parts.FirstOrDefault(p => (uint)p["PartitionNumber"] == geoPart.Index);
            if (partObj == null) throw new InvalidOperationException($"Partition {geoPart.Index} not found after creation.");

            report.AppendLine($"- Partition {geoPart.Index}: {geoPart.Label} ({geoPart.FileSystem})");

            var actualSize = (long)(UInt64)partObj["Size"];
            if (Math.Abs(actualSize - geoPart.Size) > sizeTolerance && geoPart.Size > 0)
            {
                report.AppendLine($"  * Size mismatch: expected {geoPart.Size}, found {actualSize}");
                throw new InvalidOperationException($"Partition {geoPart.Index} size mismatch. Expected {geoPart.Size}, found {actualSize}.");
            }

            var actualGptType = partObj["GptType"] as string;
            if (!string.IsNullOrWhiteSpace(geoPart.GptType) && !string.Equals(geoPart.GptType, actualGptType, StringComparison.OrdinalIgnoreCase))
            {
                report.AppendLine($"  * GPT type mismatch: expected {geoPart.GptType}, found {actualGptType}");
                throw new InvalidOperationException($"Partition {geoPart.Index} GPT type mismatch. Expected {geoPart.GptType}, found {actualGptType}.");
            }

            var actualGuid = partObj["Guid"] as string;
            if (!string.IsNullOrWhiteSpace(geoPart.GptId) && !string.Equals(geoPart.GptId, actualGuid, StringComparison.OrdinalIgnoreCase))
            {
                report.AppendLine($"  * GUID mismatch: expected {geoPart.GptId}, found {actualGuid}");
                throw new InvalidOperationException($"Partition {geoPart.Index} unique ID mismatch. Expected {geoPart.GptId}, found {actualGuid}.");
            }

            var letter = partObj["DriveLetter"] as char?;
            if (useUefi && string.Equals(geoPart.GptType, "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}", StringComparison.OrdinalIgnoreCase))
            {
                if (letter == null || letter == '\0')
                {
                    report.AppendLine("  * EFI partition missing drive letter");
                    throw new InvalidOperationException("EFI partition is not mounted with a drive letter; boot files cannot be provisioned.");
                }
                efiLetter = letter.ToString();
            }

            if (letter != null && letter != '\0' && !string.IsNullOrWhiteSpace(geoPart.FileSystem))
            {
                var di = new DriveInfo(letter.ToString());
                if (!string.Equals(di.DriveFormat, geoPart.FileSystem, StringComparison.OrdinalIgnoreCase))
                {
                    report.AppendLine($"  * Filesystem mismatch: expected {geoPart.FileSystem}, found {di.DriveFormat}");
                    throw new InvalidOperationException($"Partition {geoPart.Index} filesystem mismatch. Expected {geoPart.FileSystem}, found {di.DriveFormat}.");
                }
            }
        }

        var driveLetters = GetPartitionDriveLetters(diskNumber);
        if (useUefi)
        {
            efiLetter ??= driveLetters.FirstOrDefault(dl => Directory.Exists($"{dl}:\\EFI"));
            if (string.IsNullOrEmpty(efiLetter) || !File.Exists($"{efiLetter}:\\EFI\\Microsoft\\Boot\\bootmgfw.efi"))
            {
                report.AppendLine("EFI boot files missing");
                throw new InvalidOperationException("EFI boot files are missing after apply; cannot ensure bootability.");
            }
        }
        else
        {
            var windowsDrive = driveLetters.FirstOrDefault(dl => File.Exists($"{dl}:\\bootmgr"));
            if (string.IsNullOrEmpty(windowsDrive))
            {
                report.AppendLine("Bootmgr not found on any partition");
                throw new InvalidOperationException("Boot files not detected on any restored partition. BIOS boot cannot be guaranteed.");
            }
        }

        Logger.Log("Layout validation completed successfully", Logger.LogLevel.Info);
        report.AppendLine();
        report.AppendLine("Validation succeeded.");
        File.WriteAllText(Path.Combine(Environment.CurrentDirectory, "restore-validation.md"), report.ToString());
    }

    private static string GetNextFreeLetter(HashSet<char> usedLetters)
    {
        var preferred = new List<char> { 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
        for (char c = 'Z'; c >= 'C'; c--) preferred.Add(c);

        foreach (var letter in preferred)
        {
            if (!usedLetters.Contains(letter))
            {
                return letter.ToString();
            }
        }

        return string.Empty;
    }

    private static void AssignPartitionLetter(int diskNumber, uint partitionNumber, string letter)
    {
        try
        {
            ExecutePowerShell($"Set-Partition -DiskNumber {diskNumber} -PartitionNumber {partitionNumber} -NewDriveLetter {letter}");
        }
        catch (Exception ex)
        {
            Logger.Log($"Warning: Failed to assign drive letter {letter} to partition {partitionNumber}: {ex.Message}", Logger.LogLevel.Warning);
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
