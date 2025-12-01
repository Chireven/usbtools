using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

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

            // Verify source drive exists
            if (!Directory.Exists($"{sourceDrive}:\\"))
            {
                Logger.Log($"Error: Source drive {sourceDrive}: does not exist or is not accessible.", Logger.LogLevel.Error);
                return 1;
            }

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
                var result = await CaptureWithWimApiAsync(geometry, destinationWim, compression, metadataJson);
                if (result == 0)
                {
                    Logger.Log("Backup completed successfully. Verifying image...", Logger.LogLevel.Info);
                    return await TestCommand.ExecuteAsync(destinationWim);
                }
                return result;
            }
            else
            {
                var result = await CaptureWithDismAsync(geometry, destinationWim, compression, metadataJson);
                if (result == 0)
                {
                    Logger.Log("Backup completed successfully. Verifying image...", Logger.LogLevel.Info);
                    return await TestCommand.ExecuteAsync(destinationWim);
                }
                return result;
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
            WimApi.WIMSetTemporaryPath(wimHandle, tempPath);

            int imageIndex = 1;
            foreach (var partition in geometry.Partitions.Where(p => !string.IsNullOrEmpty(p.Letter)))
            {
                Logger.Log($"Capturing partition {partition.Index} ({partition.Letter}) - {partition.Label ?? "No Label"}", Logger.LogLevel.Info);

                // Ensure path ends with backslash
                var sourcePath = partition.Letter;
                if (!sourcePath.EndsWith("\\")) sourcePath += "\\";

                // Capture image
                // Note: WIM_FLAG_VERIFY is good practice
                var imageHandle = WimApi.WIMCaptureImage(wimHandle, sourcePath, WimApi.WIM_FLAG_VERIFY);
                
                if (imageHandle == IntPtr.Zero)
                {
                    var error = Marshal.GetLastWin32Error();
                    var errorDesc = ErrorCodeHelper.GetErrorDescription(error);
                    Logger.Log($"Failed to capture partition {partition.Letter}. Error {error}: {errorDesc}", Logger.LogLevel.Error);
                    throw new Exception($"Capture failed for {partition.Letter}. {errorDesc}");
                }

                // Set image info
                var imageName = $"Partition_{partition.Index}_{partition.Letter.TrimEnd(':', '\\')}";
                // We'll set the description for the first image to the metadata JSON later
                // For now, just set a placeholder or standard description
                var description = $"Partition {partition.Index} from {geometry.TotalSize} bytes disk";

                var xmlInfo = $"<IMAGE><NAME>{imageName}</NAME><DESCRIPTION>{description}</DESCRIPTION></IMAGE>";
                Logger.Log($"Setting Image Info for {imageName}: {xmlInfo}", Logger.LogLevel.Debug);
                
                // Manual marshaling with BOM to avoid Error 1465
                var preamble = Encoding.Unicode.GetPreamble();
                var bytes = Encoding.Unicode.GetBytes(xmlInfo);
                var bytesWithNull = new byte[preamble.Length + bytes.Length + 2];
                
                Array.Copy(preamble, 0, bytesWithNull, 0, preamble.Length);
                Array.Copy(bytes, 0, bytesWithNull, preamble.Length, bytes.Length);
                
                Logger.Log($"Image Info Buffer Size: {bytesWithNull.Length} bytes", Logger.LogLevel.Debug);
                
                var ptr = Marshal.AllocHGlobal(bytesWithNull.Length);
                try
                {
                    Marshal.Copy(bytesWithNull, 0, ptr, bytesWithNull.Length);
                    if (!WimApi.WIMSetImageInformation(imageHandle, ptr, (uint)bytesWithNull.Length))
                    {
                        Logger.Log($"Warning: Failed to set image information for {imageName}. Error: {Marshal.GetLastWin32Error()}", Logger.LogLevel.Warning);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(ptr);
                }

                WimApi.WIMCloseHandle(imageHandle);
                imageIndex++;
            }

            // Inject metadata into the first image description (if any images were captured)
            if (imageIndex > 1)
            {
                Logger.Log("Injecting drive geometry metadata...", Logger.LogLevel.Info);
                // Re-open the first image to update its description with the full JSON metadata
                var firstImageHandle = WimApi.WIMLoadImage(wimHandle, 1);
                if (firstImageHandle != IntPtr.Zero)
                {
                    // We need to get the current name to preserve it
                    // But WIMSetImageInformation takes both name and description.
                    // Let's just reconstruct the name we used: Partition_{Index}_{Letter}
                    // The first partition in our loop was the one at index 1.
                    var firstPartition = geometry.Partitions.Where(p => !string.IsNullOrEmpty(p.Letter)).First();
                    var imageName = $"Partition_{firstPartition.Index}_{firstPartition.Letter.TrimEnd(':', '\\')}";
                    
                    // Escape metadata for XML
                    var escapedMetadata = metadata.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
                    var xmlInfo = $"<IMAGE><NAME>{imageName}</NAME><DESCRIPTION>{escapedMetadata}</DESCRIPTION></IMAGE>";
                    Logger.Log($"Injecting Metadata for {imageName}: {xmlInfo}", Logger.LogLevel.Debug);
                    
                    var preamble = Encoding.Unicode.GetPreamble();
                    var bytes = Encoding.Unicode.GetBytes(xmlInfo);
                    var bytesWithNull = new byte[preamble.Length + bytes.Length + 2];
                    
                    Array.Copy(preamble, 0, bytesWithNull, 0, preamble.Length);
                    Array.Copy(bytes, 0, bytesWithNull, preamble.Length, bytes.Length);

                    Logger.Log($"Metadata Buffer Size: {bytesWithNull.Length} bytes", Logger.LogLevel.Debug);

                    var ptr = Marshal.AllocHGlobal(bytesWithNull.Length);
                    try
                    {
                        Marshal.Copy(bytesWithNull, 0, ptr, bytesWithNull.Length);
                        if (!WimApi.WIMSetImageInformation(firstImageHandle, ptr, (uint)bytesWithNull.Length))
                        {
                             Logger.Log($"Error: Failed to inject metadata. Error: {Marshal.GetLastWin32Error()}", Logger.LogLevel.Error);
                        }
                        else
                        {
                            Logger.Log("Metadata injected successfully.", Logger.LogLevel.Info);
                        }
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(ptr);
                    }
                    
                    WimApi.WIMCloseHandle(firstImageHandle);
                }
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
