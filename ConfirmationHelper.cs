namespace USBTools;

/// <summary>
/// Helper class for user confirmations with color-coded warnings
/// </summary>
public static class ConfirmationHelper
{
    public static bool Confirm(string message, bool autoYes = false)
    {
        if (autoYes)
        {
            Logger.Log("Auto-confirming (--yes flag set)", Logger.LogLevel.Info);
            return true;
        }

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("⚠ WARNING ⚠");
        Console.ResetColor();
        Console.WriteLine(message);
        Console.WriteLine();

        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine("ALL DATA ON THIS DRIVE WILL BE PERMANENTLY ERASED");
        Console.ResetColor();
        Console.WriteLine();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.Write("Are you sure you want to continue? (y/N): ");
        Console.ResetColor();

        var response = Console.ReadLine();
        var confirmed = !string.IsNullOrEmpty(response) && response.ToLowerInvariant() == "y";

        if (confirmed)
        {
            Logger.Log("User confirmed operation", Logger.LogLevel.Info);
        }
        else
        {
            Logger.Log("User cancelled operation", Logger.LogLevel.Info);
        }

        return confirmed;
    }
}
