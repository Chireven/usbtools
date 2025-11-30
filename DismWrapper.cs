using System.Diagnostics;

namespace USBTools;

/// <summary>
/// Fallback wrapper for DISM.exe when wimgapi.dll is unavailable
/// </summary>
public static class DismWrapper
{
    public static bool CaptureImage(string sourcePath, string wimPath, int imageIndex, string imageName, string imageDescription, string compression)
    {
        Logger.Log("Using DISM fallback for capture operation", Logger.LogLevel.Warning);

        var compressionArg = compression.ToLowerInvariant() switch
        {
            "none" => "none",
            "fast" => "fast",
            "max" => "max",
            _ => "fast"
        };

        var arguments = $"/Capture-Image /ImageFile:\"{wimPath}\" /CaptureDir:\"{sourcePath}\" /Name:\"{imageName}\" /Description:\"{imageDescription}\" /Compress:{compressionArg}";

        if (imageIndex > 1)
        {
            // For subsequent images, we append
            arguments += " /Append";
        }

        return ExecuteDism(arguments);
    }

    public static bool ApplyImage(string wimPath, int imageIndex, string targetPath)
    {
        Logger.Log("Using DISM fallback for apply operation", Logger.LogLevel.Warning);
        var arguments = $"/Apply-Image /ImageFile:\"{wimPath}\" /Index:{imageIndex} /ApplyDir:\"{targetPath}\"";
        return ExecuteDism(arguments);
    }

    public static bool SetImageDescription(string wimPath, int imageIndex, string description)
    {
        Logger.Log("Using DISM fallback to set image description", Logger.LogLevel.Warning);
        
        // DISM doesn't directly support setting description, but we can set it via XML
        // This is a simplified approach - in real implementation you'd manipulate the WIM XML
        var arguments = $"/Set-ImageDescription /ImageFile:\"{wimPath}\" /Index:{imageIndex} /Description:\"{description}\"";
        return ExecuteDism(arguments);
    }

    private static bool ExecuteDism(string arguments)
    {
        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = "dism.exe",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            Logger.Log($"Executing: dism.exe {arguments}", Logger.LogLevel.Debug);

            using var process = Process.Start(psi);
            if (process == null)
            {
                Logger.Log("Failed to start DISM process", Logger.LogLevel.Error);
                return false;
            }

            // Read output
            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();

            process.WaitForExit();

            if (!string.IsNullOrWhiteSpace(output))
            {
                Logger.Log($"DISM Output: {output}", Logger.LogLevel.Debug);
            }

            if (!string.IsNullOrWhiteSpace(error))
            {
                Logger.Log($"DISM Error: {error}", Logger.LogLevel.Error);
            }

            if (process.ExitCode != 0)
            {
                Logger.Log($"DISM exited with code {process.ExitCode}", Logger.LogLevel.Error);
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "DISM execution failed");
            return false;
        }
    }
}
