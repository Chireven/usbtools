using System.Text.Json;

namespace USBTools;

public class RestoreCheckpoint
{
    public HashSet<int> CompletedImages { get; set; } = new();
    public string Provider { get; set; } = "auto";

    public static string GetPath(string wimPath)
    {
        var dir = Path.GetDirectoryName(wimPath) ?? Environment.CurrentDirectory;
        var name = Path.GetFileNameWithoutExtension(wimPath);
        return Path.Combine(dir, $"{name}.restore.chkpt.json");
    }

    public static RestoreCheckpoint Load(string wimPath)
    {
        var path = GetPath(wimPath);
        if (!File.Exists(path))
        {
            return new RestoreCheckpoint();
        }

        try
        {
            var json = File.ReadAllText(path);
            var loaded = JsonSerializer.Deserialize<RestoreCheckpoint>(json);
            return loaded ?? new RestoreCheckpoint();
        }
        catch
        {
            return new RestoreCheckpoint();
        }
    }

    public void Save(string wimPath)
    {
        var path = GetPath(wimPath);
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json);
    }

    public void Clear(string wimPath)
    {
        var path = GetPath(wimPath);
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
