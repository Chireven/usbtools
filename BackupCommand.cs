using System.Runtime.InteropServices;

namespace USBTools;

/// <summary>
/// Handles backup (capture) operations from USB to WIM
/// </summary>
public static class BackupCommand
{
    public static async Task<int> ExecuteAsync(string sourceDrive, string destinationWim, string compression, string provider)
    {
        try
        {
            Logger.Log($"Starting backup operation", Logger.LogLevel.Info);
            Logger.Log($"Source Drive: {sourceDrive}", Logger.LogLevel.Info);
            Logger.Log($"Destination WIM: {destinationWim}", Logger.LogLevel.Info);
            Logger.Log($"Compression: {compression}", Logger.LogLevel.Info);
            Logger.Log($"Provider: {provider}", Logger.LogLevel.Info);

            // Normalize drive letter
            sourceDrive = sourceDrive.TrimEnd(':', '\\');

            // Analyze source drive
            var geometry = DriveAnalyzer.AnalyzeDrive(sourceDrive);
            var metadataJson = geometry.ToJson();
            Logger.Log($"Drive geometry captured: {geometry.Partitions.Count} partitions", Logger.LogLevel.Info);

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
                return await CaptureWithWimApiAsync(geometry, destinationWim, compression, metadataJson);
            }
            else
            {
                return await CaptureWithDismAsync(geometry, destinationWim, compression, metadataJson);
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Backup operation failed");
            return 1;
        }
    }

    private static async Task<int> CaptureWithWimApiAsync(DriveGeometry geometry, string wimPath, string compression, string metadata)
    {
        Logger.Log("Using WIM API for capture", Logger.LogLevel.Info);

        IntPtr wimHandle = IntPtr.Zero;
        var imageHandles = new List<IntPtr>();

        try
        {
            // Create WIM file
            var compressionType = WimApi.GetCompressionType(compression);
            wimHandle = WimApi.WIMCreateFile(
                wimPath,
                WimApi.GENERIC_WRITE,
                WimApi.WIM_CREATE_ALWAYS,
                0,
                compressionType,
                out _);

            if (wimHandle == IntPtr.Zero)
            {
                var error = Marshal.GetLastWin32Error();
                var errorDesc = ErrorCodeHelper.GetErrorDescription(error);
                Logger.Log($"WIM API Error {error}: {errorDesc}", Logger.LogLevel.Error);
                throw new Exception($"Failed to create WIM file. {errorDesc}");
            }

            Logger.Log($"WIM file created: {wimPath}", Logger.LogLevel.Info);

            // Set temporary path
            var tempPath = Path.Combine(Path.GetTempPath(), "usbtools_wim");
            Directory.CreateDirectory(tempPath);
            WimApi.WIMSetTemporaryPath(wimHandle, tempPath);

            // Capture each partition that has a drive letter
            var partitionIndex = 1;
            foreach (var partition in geometry.Partitions.Where(p => !string.IsNullOrEmpty(p.Letter)))
            {
                Logger.Log($"Capturing partition {partition.Index} ({partition.Letter}:) - {partition.Label ?? "No Label"}", Logger.LogLevel.Info);

                var sourcePath = $"{partition.Letter}:\\";
                
                // Capture image
                var imageHandle = WimApi.WIMCaptureImage(
                    wimHandle,
                    sourcePath,
                    WimApi.WIM_FLAG_VERIFY);

                if (imageHandle == IntPtr.Zero)
                {
                    var error = Marshal.GetLastWin32Error();
                    var errorDesc = ErrorCodeHelper.GetErrorDescription(error);
                    Logger.Log($"Failed to capture partition {partition.Letter}:", Logger.LogLevel.Error);
                    Logger.Log($"  WIM API Error {error}: {errorDesc}", Logger.LogLevel.Error);
                    Logger.Log($"  Partition: {partition.Letter}: ({partition.FileSystem}) - {partition.Size:N0} bytes", Logger.LogLevel.Error);
                    continue;
                }

                imageHandles.Add(imageHandle);

                // Set metadata for the first image
                if (partitionIndex == 1)
                {
                    Logger.Log("Storing drive geometry metadata in WIM...", Logger.LogLevel.Info);
                    
                    // Get the existing image information
                    if (WimApi.WIMGetImageInformation(wimHandle, out var existingInfoPtr, out var existingInfoSize))
                    {
                        var existingXml = Marshal.PtrToStringUni(existingInfoPtr, (int)existingInfoSize / 2);
                        
                        if (existingXml != null)
                        {
                            // Find the first IMAGE tag and insert DESCRIPTION after it
                            var imageTagPos = existingXml.IndexOf("<IMAGE INDEX=\"1\">");
                            if (imageTagPos == -1)
                            {
                                imageTagPos = existingXml.IndexOf("<IMAGE>");
                            }
                            
                            if (imageTagPos != -1)
                            {
                                // Find the end of the IMAGE tag
                                var imageTagEnd = existingXml.IndexOf(">", imageTagPos) + 1;
                                
                                // Insert the DESCRIPTION
                                var modifiedXml = existingXml.Substring(0, imageTagEnd) +
                                                 $"\n<DESCRIPTION>{metadata}</DESCRIPTION>" +
                                                 existingXml.Substring(imageTagEnd);
                                
                                // Set it back
                                if (!WimApi.WIMSetImageInformation(wimHandle, modifiedXml))
                                {
                                    var error = Marshal.GetLastWin32Error();
                                    var errorDesc = ErrorCodeHelper.GetErrorDescription(error);
                                    Logger.Log($"WARNING: Failed to store metadata - Error {error}: {errorDesc}", Logger.LogLevel.Warning);
                                    Logger.Log("The WIM file will be created but restore may not work properly", Logger.LogLevel.Warning);
                                }
                                else
                                {
                                    Logger.Log($"Successfully stored metadata ({metadata.Length} chars)", Logger.LogLevel.Info);
                                    Logger.Log("Metadata stored in first image DESCRIPTION field", Logger.LogLevel.Debug);
                                }
                            }
                            else
                            {
                                Logger.Log("WARNING: Could not find IMAGE tag in WIM XML", Logger.LogLevel.Warning);
                            }
                        }
                    }
                    else
                    {
                        Logger.Log("WARNING: Could not read existing WIM metadata structure", Logger.LogLevel.Warning);
                    }
                }

                Logger.Log($"Partition {partition.Letter}: captured successfully (Index {partitionIndex})", Logger.LogLevel.Info);
                partitionIndex++;
            }

            Logger.Log("All partitions captured successfully", Logger.LogLevel.Info);
            return 0;
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "WIM API capture failed");
            return 1;
        }
        finally
        {
            // Cleanup handles
            foreach (var handle in imageHandles)
            {
                if (handle != IntPtr.Zero)
                {
                    WimApi.WIMCloseHandle(handle);
                }
            }

            if (wimHandle != IntPtr.Zero)
            {
                WimApi.WIMCloseHandle(wimHandle);
            }
        }
    }

    private static async Task<int> CaptureWithDismAsync(DriveGeometry geometry, string wimPath, string compression, string metadata)
    {
        Logger.Log("Falling back to DISM for capture", Logger.LogLevel.Warning);

        try
        {
            var imageIndex = 1;
            foreach (var partition in geometry.Partitions.Where(p => !string.IsNullOrEmpty(p.Letter)))
            {
                Logger.Log($"Capturing partition {partition.Index} ({partition.Letter}:) - {partition.Label ?? "No Label"}", Logger.LogLevel.Info);

                var sourcePath = $"{partition.Letter}:\\";
                var imageName = $"Partition_{partition.Index}_{partition.Letter}";
                var description = imageIndex == 1 ? metadata : $"Partition {partition.Index}";

                var success = DismWrapper.CaptureImage(
                    sourcePath,
                    wimPath,
                    imageIndex,
                    imageName,
                    description,
                    compression);

                if (!success)
                {
                    Logger.Log($"Failed to capture partition {partition.Letter} with DISM", Logger.LogLevel.Error);
                    return 1;
                }

                Logger.Log($"Partition {partition.Letter}: captured successfully (Index {imageIndex})", Logger.LogLevel.Info);
                imageIndex++;
            }

            Logger.Log("All partitions captured successfully with DISM", Logger.LogLevel.Info);
            return 0;
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "DISM capture failed");
            return 1;
        }
    }
}
