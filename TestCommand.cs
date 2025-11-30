using System.Runtime.InteropServices;
using System.Text.Json;

namespace USBTools;

public static class TestCommand
{
    public static Task<int> ExecuteAsync(string sourcePath)
    {
        Logger.Log($"Testing image: {sourcePath}", Logger.LogLevel.Info);

        if (!File.Exists(sourcePath))
        {
            Logger.Log($"Error: File not found: {sourcePath}", Logger.LogLevel.Error);
            return Task.FromResult(1);
        }

        // 1. Check extension
        var extension = Path.GetExtension(sourcePath).ToLowerInvariant();
        if (extension != ".usbwim")
        {
            Logger.Log($"Warning: File extension is '{extension}', expected '.usbwim'", Logger.LogLevel.Warning);
            Logger.Log("This file may not have been created with USBTools or may be renamed.", Logger.LogLevel.Warning);
        }
        else
        {
            Logger.Log("File extension verified (.usbwim)", Logger.LogLevel.Info);
        }

        // 2. Open WIM and check validity
        IntPtr wimHandle = IntPtr.Zero;
        try
        {
            Logger.Log("Opening WIM file...", Logger.LogLevel.Info);
            wimHandle = WimApi.WIMCreateFile(
                sourcePath,
                WimApi.GENERIC_READ,
                WimApi.WIM_OPEN_EXISTING,
                0,
                0,
                out _);

            if (wimHandle == IntPtr.Zero)
            {
                var error = Marshal.GetLastWin32Error();
                var errorDesc = ErrorCodeHelper.GetErrorDescription(error);
                Logger.Log($"Error: Failed to open WIM file - {errorDesc}", Logger.LogLevel.Error);
                return Task.FromResult(1);
            }

            Logger.Log("WIM file is valid and accessible", Logger.LogLevel.Info);

            // 3. Check Image Count
            var imageCount = WimApi.WIMGetImageCount(wimHandle);
            Logger.Log($"WIM contains {imageCount} image(s)", Logger.LogLevel.Info);

            if (imageCount == 0)
            {
                Logger.Log("Error: WIM file contains no images", Logger.LogLevel.Error);
                return Task.FromResult(1);
            }

            // 4. Check Metadata
            Logger.Log("Checking for USBTools metadata...", Logger.LogLevel.Info);
            if (WimApi.WIMGetImageInformation(wimHandle, out var infoPtr, out var infoSize))
            {
                var xmlInfo = Marshal.PtrToStringUni(infoPtr, (int)infoSize / 2);
                
                if (string.IsNullOrEmpty(xmlInfo))
                {
                    Logger.Log("Error: Failed to read WIM XML information", Logger.LogLevel.Error);
                    return Task.FromResult(1);
                }

                // Look for description
                var descStart = xmlInfo.IndexOf("<DESCRIPTION>");
                var descEnd = xmlInfo.IndexOf("</DESCRIPTION>");

                if (descStart != -1 && descEnd != -1)
                {
                    descStart += "<DESCRIPTION>".Length;
                    var json = xmlInfo.Substring(descStart, descEnd - descStart);
                    
                    try
                    {
                        var geometry = DriveGeometry.FromJson(json);
                        if (geometry != null)
                        {
                            Logger.Log("✅ USBTools metadata found and valid", Logger.LogLevel.Info);
                            Logger.Log("Drive Geometry Details:", Logger.LogLevel.Info);
                            Logger.Log($"  Partition Style: {geometry.PartitionStyle}", Logger.LogLevel.Info);
                            Logger.Log($"  Total Size: {geometry.TotalSize:N0} bytes", Logger.LogLevel.Info);
                            Logger.Log($"  Partitions: {geometry.Partitions.Count}", Logger.LogLevel.Info);
                            
                            // Calculate minimum size required
                            // For fixed partitions, use full size. For variable (data) partitions, use UsedSpace + overhead.
                            // To be safe, we'll use UsedSpace + 20% buffer for variable partitions, or full size if fixed.
                            // But simpler approach requested: "use used space values".
                            // Let's use: Sum(Max(UsedSpace + 100MB, 10MB)) + 10MB partition table overhead
                            // We add 100MB buffer to used space to allow for filesystem structures/journaling.
                            
                            long minSize = 0;
                            foreach(var p in geometry.Partitions)
                            {
                                if (p.IsFixed)
                                {
                                    minSize += p.Size;
                                }
                                else
                                {
                                    // If UsedSpace is 0 (e.g. failed to get it), fallback to Size
                                    long effectiveUsed = p.UsedSpace > 0 ? p.UsedSpace : p.Size;
                                    // Add 100MB buffer for filesystem overhead on restore
                                    minSize += effectiveUsed + (100 * 1024 * 1024);
                                }
                            }
                            
                            // Add global overhead
                            minSize += (10 * 1024 * 1024);

                            Logger.Log($"  Minimum Drive Size: {minSize:N0} bytes ({minSize / (1024.0 * 1024 * 1024):F2} GB)", Logger.LogLevel.Info);

                            foreach(var p in geometry.Partitions)
                            {
                                var usedStr = p.UsedSpace > 0 ? $"Used: {p.UsedSpace:N0} bytes" : "Used: Unknown";
                                Logger.Log($"    - Partition {p.Index}: {p.FileSystem} ({p.Size:N0} bytes) [{usedStr}] {(p.IsActive ? "[Active]" : "")}", Logger.LogLevel.Info);
                            }

                            // Verify partition count matches image count (roughly)
                            // Note: MSR partitions might be in geometry but not captured as images, so exact match isn't always required,
                            // but usually for USBTools we capture everything relevant.
                            // Let's just log the comparison.
                            if (geometry.Partitions.Count(p => !string.IsNullOrEmpty(p.Letter)) != imageCount)
                            {
                                // This might be okay if we excluded some, but worth noting
                                Logger.Log($"Note: Metadata lists {geometry.Partitions.Count} partitions, WIM has {imageCount} images.", Logger.LogLevel.Debug);
                            }
                            
                            return Task.FromResult(0);
                        }
                        else
                        {
                            Logger.Log("Error: Metadata found but parsed as null", Logger.LogLevel.Error);
                            return Task.FromResult(1);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"Error: Metadata found but invalid JSON: {ex.Message}", Logger.LogLevel.Error);
                        Logger.Log($"Raw Metadata: {json}", Logger.LogLevel.Debug);
                        return Task.FromResult(1);
                    }
                }
                else
                {
                    Logger.Log("❌ Error: No USBTools metadata found in DESCRIPTION field", Logger.LogLevel.Error);
                    Logger.Log("This image cannot be restored using 'usbtools restore' with automatic geometry.", Logger.LogLevel.Error);
                    return Task.FromResult(1);
                }
            }
            else
            {
                Logger.Log("Error: Failed to retrieve image information from WIM", Logger.LogLevel.Error);
                return Task.FromResult(1);
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error testing image");
            return Task.FromResult(1);
        }
        finally
        {
            if (wimHandle != IntPtr.Zero)
            {
                WimApi.WIMCloseHandle(wimHandle);
            }
        }
    }
}
