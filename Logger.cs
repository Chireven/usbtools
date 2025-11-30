using System.Diagnostics;

namespace USBTools;

public static class Logger
{
    private static string? _logFilePath;
    private static readonly object _lock = new();
    public static bool DebugMode { get; set; } = false;

    public static void Initialize(string logFileName = "usbtools.log")
    {
        _logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, logFileName);
        Log("=== USB Tools Session Started ===", LogLevel.Info);
    }

    public enum LogLevel
    {
        Debug,
        Info,
        Warning,
        Error
    }

    public static void Log(string message, LogLevel level = LogLevel.Info)
    {
        // Skip debug messages unless debug mode is enabled
        if (level == LogLevel.Debug && !DebugMode)
        {
            return;
        }

        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        var logMessage = $"[{timestamp}] [{level}] {message}";

        // Console output with colors
        lock (_lock)
        {
            Console.ForegroundColor = level switch
            {
                LogLevel.Debug => ConsoleColor.DarkGray,
                LogLevel.Info => ConsoleColor.Green,
                LogLevel.Warning => ConsoleColor.Yellow,
                LogLevel.Error => ConsoleColor.Red,
                _ => ConsoleColor.White
            };
            Console.WriteLine(logMessage);
            Console.ResetColor();

            // File output
            if (_logFilePath != null)
            {
                try
                {
                    File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
                }
                catch
                {
                    // Ignore file write errors
                }
            }
        }
    }

    public static void LogException(Exception ex, string context = "")
    {
        var message = string.IsNullOrEmpty(context) 
            ? $"Exception: {ex.Message}\n{ex.StackTrace}" 
            : $"{context}: {ex.Message}\n{ex.StackTrace}";
        Log(message, LogLevel.Error);
    }
}
