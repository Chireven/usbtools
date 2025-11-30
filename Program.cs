namespace USBTools;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Initialize logger
        Logger.Initialize();

        // Banner with colors
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("╔═══════════════════════════════════════╗");
        Console.WriteLine("║    USB Tools v0.1.0                   ║");
        Console.WriteLine("║  USB ↔ WIM Conversion Tool            ║");
        Console.WriteLine("╚═══════════════════════════════════════╝");
        Console.ResetColor();
        Console.WriteLine();

        if (args.Length == 0)
        {
            ShowHelp();
            return 0;
        }

        var command = args[0].ToLowerInvariant();

        try
        {
            return command switch
            {
                "backup" => await ExecuteBackupAsync(args),
                "restore" => await ExecuteRestoreAsync(args),
                "help" or "--help" or "-h" or "-?" => ShowHelp(),
                _ => ShowInvalidCommand(args[0])
            };
        }
        catch (Exception ex)
        {
            Logger.LogException(ex, "Fatal error");
            return 1;
        }
    }

    private static async Task<int> ExecuteBackupAsync(string[] args)
    {
        string? source = null;
        string? destination = null;
        string compression = "fast";

        // Parse arguments
        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i].ToLowerInvariant();
            
            if ((arg == "--source" || arg == "-s") && i + 1 < args.Length)
            {
                source = args[++i];
            }
            else if ((arg == "--destination" || arg == "-d") && i + 1 < args.Length)
            {
                destination = args[++i];
            }
            else if ((arg == "--compression" || arg == "-c") && i + 1 < args.Length)
            {
                compression = args[++i];
            }
            else if (arg == "--help" || arg == "-h")
            {
                ShowBackupHelp();
                return 0;
            }
        }

        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(destination))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: Both --source and --destination are required.");
            Console.ResetColor();
            Console.WriteLine();
            ShowBackupHelp();
            return 1;
        }

        return await BackupCommand.ExecuteAsync(source, destination, compression);
    }

    private static async Task<int> ExecuteRestoreAsync(string[] args)
    {
        string? source = null;
        string? target = null;
        string? diskNumber = null;
        bool autoYes = false;

        // Parse arguments
        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i].ToLowerInvariant();
            
            if ((arg == "--source" || arg == "-s") && i + 1 < args.Length)
            {
                source = args[++i];
            }
            else if ((arg == "--target" || arg == "-t") && i + 1 < args.Length)
            {
                target = args[++i];
            }
            else if ((arg == "--disk" || arg == "-k") && i + 1 < args.Length)
            {
                diskNumber = args[++i];
            }
            else if (arg == "--yes" || arg == "-y")
            {
                autoYes = true;
            }
            else if (arg == "--help" || arg == "-h")
            {
                ShowRestoreHelp();
                return 0;
            }
        }

        if (string.IsNullOrEmpty(source))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: --source is required.");
            Console.ResetColor();
            Console.WriteLine();
            ShowRestoreHelp();
            return 1;
        }

        // Must have either --target or --disk, but not both
        if (string.IsNullOrEmpty(target) && string.IsNullOrEmpty(diskNumber))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: Either --target or --disk is required.");
            Console.ResetColor();
            Console.WriteLine();
            ShowRestoreHelp();
            return 1;
        }

        if (!string.IsNullOrEmpty(target) && !string.IsNullOrEmpty(diskNumber))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: Cannot specify both --target and --disk.");
            Console.ResetColor();
            Console.WriteLine();
            ShowRestoreHelp();
            return 1;
        }

        return await RestoreCommand.ExecuteAsync(source, target, diskNumber, autoYes);
    }

    private static int ShowHelp()
    {
        Console.WriteLine("USB to WIM imaging tool - Backup and restore bootable USB drives");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Usage:");
        Console.ResetColor();
        Console.WriteLine("  usbtools <command> [options]");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Commands:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  backup   ");
        Console.ResetColor();
        Console.WriteLine("Capture USB drive to WIM image");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  restore  ");
        Console.ResetColor();
        Console.WriteLine("Restore WIM image to USB drive");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  help     ");
        Console.ResetColor();
        Console.WriteLine("Show this help message");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("Run 'usbtools <command> --help' for more information on a command.");
        Console.ResetColor();
        return 0;
    }

    private static void ShowBackupHelp()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.Write("Usage: ");
        Console.ResetColor();
        Console.WriteLine("usbtools backup --source <drive> --destination <wimfile> [options]");
        Console.WriteLine();
        Console.WriteLine("Capture a USB drive to a WIM image");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Options:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -s, --source <drive>           ");
        Console.ResetColor();
        Console.WriteLine("Source drive letter (e.g., E:)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -d, --destination <wimfile>    ");
        Console.ResetColor();
        Console.WriteLine("Destination WIM file path");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -c, --compression <level>      ");
        Console.ResetColor();
        Console.WriteLine("Compression: none, fast, max (default: fast)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -h, --help                     ");
        Console.ResetColor();
        Console.WriteLine("Show help");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Example:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("  usbtools backup -s E: -d C:\\Backups\\usb.wim --compression max");
        Console.ResetColor();
    }

    private static void ShowRestoreHelp()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.Write("Usage: ");
        Console.ResetColor();
        Console.WriteLine("usbtools restore --source <wimfile> --target <drive|disk> [options]");
        Console.WriteLine();
        Console.WriteLine("Restore a WIM image to a USB drive");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Options:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -s, --source <wimfile>    ");
        Console.ResetColor();
        Console.WriteLine("Source WIM file path");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -t, --target <drive>      ");
        Console.ResetColor();
        Console.WriteLine("Target drive letter (e.g., F:)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -k, --disk <number>       ");
        Console.ResetColor();
        Console.WriteLine("Target disk number (e.g., 2)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -y, --yes                 ");
        Console.ResetColor();
        Console.WriteLine("Auto-confirm without prompting");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -h, --help                ");
        Console.ResetColor();
        Console.WriteLine("Show help");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Examples:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("  usbtools restore -s C:\\Backups\\usb.wim -t F:");
        Console.WriteLine("  usbtools restore -s C:\\Backups\\usb.wim --disk 2 --yes");
        Console.ResetColor();
    }

    private static int ShowInvalidCommand(string command)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"Error: Unknown command '{command}'");
        Console.ResetColor();
        Console.WriteLine();
        ShowHelp();
        return 1;
    }
}

