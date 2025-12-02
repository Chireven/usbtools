using System.Text.Json;

namespace USBTools;

public static class TelemetryReporter
{
    private static readonly object _lock = new();
    private static string _path = Path.Combine(Environment.CurrentDirectory, "telemetry.jsonl");

    public static void Initialize(string? path = null)
    {
        if (!string.IsNullOrEmpty(path))
        {
            _path = path;
        }
    }

    public static void Emit(string stage, string status, object? payload = null)
    {
        var record = new
        {
            timestamp = DateTime.UtcNow,
            stage,
            status,
            payload
        };

        var json = JsonSerializer.Serialize(record);
        lock (_lock)
        {
            File.AppendAllText(_path, json + Environment.NewLine);
        }
    }
}

public static class ProgressTree
{
    public static bool Verbose { get; set; }

    public static void Node(string label, int depth = 0)
    {
        if (!Verbose) return;
        var prefix = new string(' ', depth * 2);
        Logger.Log($"{prefix}└─ {DateTime.Now:HH:mm:ss} {label}", Logger.LogLevel.Info);
    }
}
