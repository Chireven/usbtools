namespace USBTools;

/// <summary>
/// Helper for countdown warnings with abort capability
/// </summary>
public static class CountdownHelper
{
    /// <summary>
    /// Shows a countdown warning with abort capability
    /// </summary>
    /// <param name="message">Warning message to display</param>
    /// <param name="seconds">Number of seconds to count down</param>
    /// <returns>True if countdown completed, false if aborted</returns>
    public static bool ShowCountdown(string message, int seconds = 5)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine();
        Console.WriteLine($"⚠ WARNING: {message}");
        Console.WriteLine();
        Console.WriteLine($"Press Ctrl+C to abort, or any other key to continue immediately...");
        Console.ResetColor();

        for (int i = seconds; i > 0; i--)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"\rContinuing in {i} second{(i > 1 ? "s" : "")}...  ");
            Console.ResetColor();

            // Check if key is available (non-blocking)
            var startTime = DateTime.Now;
            while ((DateTime.Now - startTime).TotalMilliseconds < 1000)
            {
                if (Console.KeyAvailable)
                {
                    var key = Console.ReadKey(true);
                    
                    // Ctrl+C is handled by the system, but we can detect other keys
                    Console.WriteLine();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Continuing operation...");
                    Console.ResetColor();
                    Console.WriteLine();
                    return true;
                }
                Thread.Sleep(50);
            }
        }

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Continuing operation...");
        Console.ResetColor();
        Console.WriteLine();
        return true;
    }
}
