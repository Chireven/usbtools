using System.Diagnostics;
using System.Text.RegularExpressions;

namespace USBTools;

/// <summary>
/// Analyzes drive geometry and partitions
/// </summary>
public static class DriveAnalyzer
{
    public static DriveGeometry AnalyzeDrive(string driveLetter)
    {
        Logger.Log($"Analyzing drive {driveLetter}", Logger.LogLevel.Info);

        // Get disk number from drive letter
        var diskNumber = GetDiskNumberFromDriveLetter(driveLetter);
        if (diskNumber == -1)
        {
            throw new Exception($"Could not determine disk number for drive {driveLetter}");
        }

        var geometry = new DriveGeometry();

        // Get partition style (MBR/GPT)
        geometry.PartitionStyle = GetPartitionStyle(diskNumber);
        Logger.Log($"Partition style: {geometry.PartitionStyle}", Logger.LogLevel.Info);

        // Get disk signature or GUID
        if (geometry.PartitionStyle == "MBR")
        {
            geometry.DiskSignature = GetDiskSignature(diskNumber);
        }
        else
        {
            geometry.GptDiskId = GetGptDiskId(diskNumber);
        }

        // Get total size
        geometry.TotalSize = GetDiskSize(diskNumber);
        Logger.Log($"Total disk size: {geometry.TotalSize:N0} bytes", Logger.LogLevel.Info);

        // Get all partitions on this disk
        geometry.Partitions = GetPartitions(diskNumber);
        Logger.Log($"Found {geometry.Partitions.Count} partitions", Logger.LogLevel.Info);

        return geometry;
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

    private static string GetPartitionStyle(int diskNumber)
    {
        try
        {
            var output = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty PartitionStyle");
            return output.Trim();
        }
        catch
        {
            return "MBR"; // Default fallback
        }
    }

    private static uint GetDiskSignature(int diskNumber)
    {
        try
        {
            var output = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty Signature");
            if (uint.TryParse(output.Trim(), out var signature))
            {
                return signature;
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error getting disk signature");
        }
        return 0;
    }

    private static string GetGptDiskId(int diskNumber)
    {
        try
        {
            var output = ExecutePowerShell($"Get-Disk -Number {diskNumber} | Select-Object -ExpandProperty Guid");
            return output.Trim();
        }
        catch
        {
            return Guid.NewGuid().ToString();
        }
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
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error getting disk size");
        }
        return 0;
    }

    private static List<PartitionInfo> GetPartitions(int diskNumber)
    {
        var partitions = new List<PartitionInfo>();

        try
        {
            // Get partition details using diskpart or PowerShell
            var output = ExecutePowerShell($@"
                Get-Partition -DiskNumber {diskNumber} | ForEach-Object {{
                    $obj = [PSCustomObject]@{{
                        PartitionNumber = $_.PartitionNumber
                        DriveLetter = $_.DriveLetter
                        Size = $_.Size
                        Offset = $_.Offset
                        Type = $_.Type
                        IsActive = $_.IsActive
                        GptType = $_.GptType
                        Guid = $_.Guid
                        UsedSpace = 0
                    }}
                    
                    # Get volume info for used space
                    if ($_.DriveLetter) {{
                        try {{
                            $vol = Get-Volume -DriveLetter $_.DriveLetter -ErrorAction SilentlyContinue
                            if ($vol) {{
                                $obj.UsedSpace = $vol.Size - $vol.SizeRemaining
                            }}
                        }} catch {{}}
                    }}
                    
                    $obj
                }} | ConvertTo-Json
            ");

            // Parse JSON output
            var partitionData = System.Text.Json.JsonSerializer.Deserialize<List<Dictionary<string, System.Text.Json.JsonElement>>>(output);
            
            if (partitionData != null)
            {
                foreach (var item in partitionData)
                {
                    var partition = new PartitionInfo
                    {
                        Index = item.ContainsKey("PartitionNumber") ? item["PartitionNumber"].GetInt32() : 0,
                        Size = item.ContainsKey("Size") ? item["Size"].GetInt64() : 0,
                        UsedSpace = item.ContainsKey("UsedSpace") ? item["UsedSpace"].GetInt64() : 0,
                        Offset = item.ContainsKey("Offset") ? item["Offset"].GetInt64() : 0,
                        Type = item.ContainsKey("Type") ? item["Type"].GetString() ?? "" : "",
                        IsActive = item.ContainsKey("IsActive") && item["IsActive"].GetBoolean(),
                        GptType = item.ContainsKey("GptType") ? item["GptType"].GetString() : null,
                        GptId = item.ContainsKey("Guid") ? item["Guid"].GetString() : null
                    };

                    // Get drive letter
                    if (item.ContainsKey("DriveLetter") && item["DriveLetter"].ValueKind != System.Text.Json.JsonValueKind.Null)
                    {
                        partition.Letter = item["DriveLetter"].GetString();
                    }

                    // Get volume info if there's a drive letter
                    if (!string.IsNullOrEmpty(partition.Letter))
                    {
                        try
                        {
                            var driveInfo = new DriveInfo(partition.Letter);
                            partition.Label = driveInfo.VolumeLabel;
                            partition.FileSystem = driveInfo.DriveFormat;
                        }
                        catch { }
                    }

                    // Determine if partition is "fixed" (EFI, Boot, Recovery)
                    partition.IsFixed = IsFixedPartition(partition);

                    partitions.Add(partition);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Error getting partitions");
        }

        return partitions;
    }

    private static bool IsFixedPartition(PartitionInfo partition)
    {
        // EFI System Partition
        if (partition.Type?.Contains("System", StringComparison.OrdinalIgnoreCase) == true)
            return true;

        // Microsoft Reserved Partition
        if (partition.Type?.Contains("Reserved", StringComparison.OrdinalIgnoreCase) == true)
            return true;

        // Recovery partitions
        if (partition.Type?.Contains("Recovery", StringComparison.OrdinalIgnoreCase) == true)
            return true;

        // Small boot partitions (< 1GB typically)
        if (partition.Size < 1024 * 1024 * 1024 && partition.IsActive)
            return true;

        return false;
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

    public static uint GenerateRandomDiskSignature()
    {
        var random = new Random();
        return (uint)random.Next(1, int.MaxValue);
    }

    public static string GenerateRandomGptDiskId()
    {
        return Guid.NewGuid().ToString();
    }
}
