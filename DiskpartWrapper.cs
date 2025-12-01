using System.Diagnostics;
using System.Text;

namespace USBTools;

/// <summary>
/// Wrapper for diskpart commands to handle disk operations that are unreliable via WMI/IOCTL
/// </summary>
public static class DiskpartWrapper
{
    /// <summary>
    /// Converts a disk to GPT and creates partitions using diskpart
    /// </summary>
    public static void ConvertToGptAndCreatePartitions(int diskNumber, List<PartitionInfo> partitions)
    {
        Logger.Log($"Using diskpart to convert disk {diskNumber} to GPT", Logger.LogLevel.Info);
        
        // Build diskpart script
        var script = new StringBuilder();
        script.AppendLine($"select disk {diskNumber}");
        script.AppendLine("clean");
        // script.AppendLine("convert gpt");
        
        // Create partitions
        foreach (var partition in partitions)
        {
            long sizeMB = partition.Size / (1024 * 1024);
            script.AppendLine($"create partition primary size={sizeMB}");
             script.AppendLine($"assign");

            if (partition.FileSystem.ToUpper() == "FAT32")
            {
                script.AppendLine("format fs=fat32 quick");
            }
            else if (partition.FileSystem.ToUpper() == "NTFS")
            {
                script.AppendLine("format fs=ntfs quick");
            }
            
            Logger.Log($"Partition {partition.Index}: {sizeMB}MB, {partition.FileSystem}", Logger.LogLevel.Info);
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
    }
}
