using System.Diagnostics;
using System.Text;

namespace USBTools;

/// <summary>
/// Wrapper for diskpart commands to handle disk operations that are unreliable via WMI/IOCTL
/// </summary>
public static class DiskpartWrapper
{
    /// <summary>
    /// Prepares a disk (MBR/GPT) and creates partitions using diskpart
    /// </summary>
    public static void PrepareDiskAndCreatePartitions(int diskNumber, List<PartitionInfo> partitions, string partitionStyle)
    {
        Logger.Log($"Using diskpart to convert disk {diskNumber} to {partitionStyle}", Logger.LogLevel.Info);
        
        // Build diskpart script
        var script = new StringBuilder();
        script.AppendLine($"select disk {diskNumber}");
        script.AppendLine("clean");
        script.AppendLine($"convert {partitionStyle.ToLower()}");
        
        // Create partitions
        int currentPartitionIndex = 1;
        foreach (var partition in partitions)
        {
            Logger.Log($"Creating partition {currentPartitionIndex} (Source Index: {partition.Index}, {partition.FileSystem})", Logger.LogLevel.Info);
            
            long sizeMB = partition.Size / (1024 * 1024);
            
            // Check if this is an EFI partition
            bool isEfi = false;
            if (partitionStyle.Equals("GPT", StringComparison.OrdinalIgnoreCase))
            {
                // Check by GPT Type GUID or if it's a fixed partition with FAT32 (heuristic)
                if (string.Equals(partition.GptType, "{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}", StringComparison.OrdinalIgnoreCase) ||
                    (partition.IsFixed && partition.FileSystem.ToUpper() == "FAT32"))
                {
                    isEfi = true;
                }
            }

            if (isEfi)
            {
                script.AppendLine($"create partition efi size={sizeMB}");
            }
            else
            {
                script.AppendLine($"create partition primary size={sizeMB}");
            }

            if (!string.IsNullOrWhiteSpace(partition.GptType))
            {
                script.AppendLine($"set id={partition.GptType}");
            }

            if (!string.IsNullOrWhiteSpace(partition.GptId))
            {
                script.AppendLine($"uniqueid partition id={partition.GptId}");
            }

            // script.AppendLine($"select partition {currentPartitionIndex}"); // Removed to rely on implicit selection

            if (!string.IsNullOrEmpty(partition.DriveLetter))
            {
                script.AppendLine($"assign letter={partition.DriveLetter.TrimEnd(':')}");
            }
            else
            {
                script.AppendLine($"assign");
            }

            if (partition.IsActive && partitionStyle.Equals("MBR", StringComparison.OrdinalIgnoreCase))
            {
                script.AppendLine("active");
            }
            

            if (partition.FileSystem.ToUpper() == "FAT32")
            {
                var labelCmd = string.IsNullOrEmpty(partition.Label) ? "" : $" label=\"{partition.Label}\"";
                script.AppendLine($"format fs=fat32{labelCmd} quick");
            }
            else if (partition.FileSystem.ToUpper() == "NTFS")
            {
                var labelCmd = string.IsNullOrEmpty(partition.Label) ? "" : $" label=\"{partition.Label}\"";
                script.AppendLine($"format fs=ntfs{labelCmd} quick");
            }
            
            Logger.Log($"Partition {partition.Index}: {sizeMB}MB, {partition.FileSystem}, Label='{partition.Label}'{(isEfi ? " [EFI]" : "")}", Logger.LogLevel.Info);
            currentPartitionIndex++;
        }
        
        script.AppendLine("exit");
        
        // Write script to temp file
        string scriptPath = Path.Combine(Path.GetTempPath(), $"diskpart_{Guid.NewGuid()}.txt");
        try
        {
            File.WriteAllText(scriptPath, script.ToString());
            Logger.Log($"Diskpart script written to: {scriptPath}", Logger.LogLevel.Debug);
            Logger.Log("Script contents:", Logger.LogLevel.Debug);
            foreach (var line in script.ToString().Split('\n'))
            {
                if (!string.IsNullOrWhiteSpace(line))
                    Logger.Log($"  {line.Trim()}", Logger.LogLevel.Debug);
            }
            
            // Execute diskpart
            var psi = new ProcessStartInfo
            {
                FileName = "diskpart.exe",
                Arguments = $"/s \"{scriptPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            
            Logger.Log("Executing diskpart...", Logger.LogLevel.Info);
            using var process = Process.Start(psi);
            if (process == null)
            {
                throw new Exception("Failed to start diskpart process");
            }
            
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            process.WaitForExit();
            
            Logger.Log("Diskpart stdout:", Logger.LogLevel.Info);
            foreach (var line in output.Split('\n'))
            {
                if (!string.IsNullOrWhiteSpace(line))
                    Logger.Log($"  {line.Trim()}", Logger.LogLevel.Info);
            }
            
            if (!string.IsNullOrWhiteSpace(error))
            {
                Logger.Log("Diskpart stderr:", Logger.LogLevel.Warning);
                foreach (var line in error.Split('\n'))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        Logger.Log($"  {line.Trim()}", Logger.LogLevel.Warning);
                }
            }
            
            if (process.ExitCode != 0)
            {
                Logger.Log($"Diskpart failed with exit code {process.ExitCode}", Logger.LogLevel.Error);
                throw new Exception($"Diskpart failed with exit code {process.ExitCode}. Check log for details.");
            }
            
            Logger.Log("Diskpart completed successfully", Logger.LogLevel.Info);
        }
        finally
        {
            // Clean up temp script file
            try
            {
                if (File.Exists(scriptPath))
                {
                    File.Delete(scriptPath);
                }
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }
    
    public class PartitionInfo
    {
        public int Index { get; set; }
        public long Size { get; set; }
        public string FileSystem { get; set; } = "";
        public string Label { get; set; } = "";
        public string DriveLetter { get; set; } = "";
        public bool IsActive { get; set; }
        public string? GptType { get; set; }
        public string? GptId { get; set; }
        public bool IsFixed { get; set; }
    }
}
