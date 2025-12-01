using System.Diagnostics;
using System.Reflection;

namespace USBTools;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Parse log argument early
        string? logFile = null;
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--log" || args[i] == "-l")
            {
                if (i + 1 < args.Length && !args[i + 1].StartsWith("-"))
                {
                    logFile = args[i + 1];
                }
                break;
            }
        }

        // Banner with colors
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("╔════════════════════════════╗");
        Console.WriteLine("║    USB Tools v0.1.0        ║");
        Console.WriteLine("║  USB ↔ WIM Conversion Tool  ║");
        Console.WriteLine("╚════════════════════════════╝");
        
        // Show build date
        var buildDate = File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location);
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine($"Build Date: {buildDate:yyyy-MM-dd HH:mm:ss}");
        Console.ResetColor();
        Console.WriteLine();

        // Initialize logger
        Logger.Initialize(logFile);

        // Check for global debug flag
        if (args.Contains("--debug") || args.Contains("-v"))
        {
            Logger.DebugMode = true;
            Logger.Log("Debug mode enabled", Logger.LogLevel.Debug);
        }

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
                "test" => await ExecuteTestAsync(args),
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
        string provider = "auto";
        bool overwrite = false;

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
            else if ((arg == "--provider" || arg == "-p") && i + 1 < args.Length)
            {
                provider = args[++i].ToLowerInvariant();
            }
            else if (arg == "--overwrite" || arg == "-o")
            {
                overwrite = true;
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

        // Force .usbwim extension for backups
        var originalDestination = destination;
        var destExtension = Path.GetExtension(destination).ToLowerInvariant();
        
        if (destExtension != ".usbwim")
        {
            var destWithoutExt = Path.GetFileNameWithoutExtension(destination);
            var destDir = Path.GetDirectoryName(destination);
            destination = string.IsNullOrEmpty(destDir) 
                ? $"{destWithoutExt}.usbwim" 
                : Path.Combine(destDir, $"{destWithoutExt}.usbwim");
            
            if (!string.IsNullOrEmpty(destExtension))
            {
                Logger.Log($"Changed extension from '{destExtension}' to '.usbwim'", Logger.LogLevel.Warning);
                Logger.Log($"Original: {originalDestination}", Logger.LogLevel.Debug);
                Logger.Log($"Modified: {destination}", Logger.LogLevel.Debug);
            }
        }

        // Check if destination exists and overwrite not specified
        if (File.Exists(destination) && !overwrite)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: File '{destination}' already exists.");
            Console.WriteLine("Use --overwrite or -o to overwrite the existing file.");
            Console.ResetColor();
            return 1;
        }

        // Show countdown warning if overwriting
        if (File.Exists(destination) && overwrite)
        {
            if (!CountdownHelper.ShowCountdown($"File '{destination}' will be OVERWRITTEN", 5))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Operation cancelled by user.");
                Console.ResetColor();
                return 1;
            }
        }

        return await BackupCommand.ExecuteAsync(source, destination, compression, provider);
    }

    private static async Task<int> ExecuteRestoreAsync(string[] args)
    {
        string? source = null;
        string? target = null;
        string? diskNumber = null;
        bool autoYes = false;
        string provider = "auto";
        string bootMode = "auto";
        bool forceBoot = false;

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
            else if ((arg == "--provider" || arg == "-p") && i + 1 < args.Length)
            {
                provider = args[++i].ToLowerInvariant();
            }
            else if ((arg == "--boot-mode" || arg == "-b") && i + 1 < args.Length)
            {
                bootMode = args[++i].ToLowerInvariant();
                if (bootMode != "uefi" && bootMode != "bios" && bootMode != "auto")
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Error: Invalid boot mode '{bootMode}'. Valid values are: uefi, bios, auto.");
                    Console.ResetColor();
                    return 1;
                }
            }
            else if (arg == "--force-boot" || arg == "-f")
            {
                forceBoot = true;
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

        // Resolve source path (handle missing extension)
        var resolvedSource = ResolveSourcePath(source);
        if (resolvedSource == null)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: Source file '{source}' not found.");
            Console.ResetColor();
            return 1;
        }
        source = resolvedSource;

        return await RestoreCommand.ExecuteAsync(source, target, diskNumber, autoYes, provider, bootMode, forceBoot);
    }

    private static async Task<int> ExecuteTestAsync(string[] args)
    {
        string? source = null;

        for (int i = 1; i < args.Length; i++)
        {
            var arg = args[i].ToLowerInvariant();
            
            if ((arg == "--source" || arg == "-s") && i + 1 < args.Length)
            {
                source = args[++i];
            }
            else if (arg == "--debug" || arg == "-v")
            {
                // Already handled in Main
            }
            else if (arg == "--help" || arg == "-h")
            {
                ShowTestHelp();
                return 0;
            }
        }

        if (string.IsNullOrEmpty(source))
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error: --source is required.");
            Console.ResetColor();
            Console.WriteLine();
            ShowTestHelp();
            return 1;
        }

        // Resolve source path (handle missing extension)
        var resolvedSource = ResolveSourcePath(source);
        if (resolvedSource == null)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error: Source file '{source}' not found.");
            Console.ResetColor();
            return 1;
        }
        source = resolvedSource;

        return await TestCommand.ExecuteAsync(source);
    }

    private static string? ResolveSourcePath(string sourcePath)
    {
        if (File.Exists(sourcePath))
        {
            return sourcePath;
        }

        // Check if it's missing .usbwim extension
        if (string.IsNullOrEmpty(Path.GetExtension(sourcePath)))
        {
            var withExt = sourcePath + ".usbwim";
            if (File.Exists(withExt))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Note: Using '{withExt}'");
                Console.ResetColor();
                return withExt;
            }
        }

        // Check if user provided .wim but .usbwim exists
        if (Path.GetExtension(sourcePath).Equals(".wim", StringComparison.OrdinalIgnoreCase))
        {
            var usbwimPath = Path.ChangeExtension(sourcePath, ".usbwim");
            if (File.Exists(usbwimPath))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"Warning: '{sourcePath}' not found, but '{usbwimPath}' exists.");
                Console.Write("Do you want to use the .usbwim file instead? [Y/n] ");
                Console.ResetColor();
                
                var response = Console.ReadLine();
                if (string.IsNullOrEmpty(response) || response.Trim().ToLowerInvariant().StartsWith("y"))
                {
                    return usbwimPath;
                }
            }
        }

        return null;
    }

    private static int ShowHelp()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Usage: usbtools <command> [options]");
        Console.WriteLine();
        Console.WriteLine("Commands:");
        Console.ResetColor();
        Console.WriteLine("  backup     Capture a USB drive to a WIM file");
        Console.WriteLine("  restore    Restore a WIM file to a USB drive");
        Console.WriteLine("  test       Verify a WIM image");
        Console.WriteLine("  help       Show this help message");
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Run 'usbtools <command> --help' for command-specific help.");
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
        Console.WriteLine("Capture a USB drive partition structure and data to a WIM file");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Options:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -s, --source <drive>      ");
        Console.ResetColor();
        Console.WriteLine("Source drive letter (e.g. E:)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -d, --destination <file>  ");
        Console.ResetColor();
        Console.WriteLine("Destination WIM file path (.usbwim)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -c, --compression <type>  ");
        Console.ResetColor();
        Console.WriteLine("Compression: none, fast (default), max");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -o, --overwrite           ");
        Console.ResetColor();
        Console.WriteLine("Overwrite existing file");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -p, --provider <type>     ");
        Console.ResetColor();
        Console.WriteLine("Provider: auto (default), wimapi, dism");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -v, --debug               ");
        Console.ResetColor();
        Console.WriteLine("Enable verbose debug logging");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -h, --help                ");
        Console.ResetColor();
        Console.WriteLine("Show help");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Example:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("  usbtools backup -s E: -d C:\\Backups\\usb.usbwim");
        Console.WriteLine("  usbtools backup -s E: -d C:\\Backups\\usb.usbwim --overwrite");
        Console.ResetColor();
    }

    private static void ShowRestoreHelp()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.Write("Usage: ");
        Console.ResetColor();
        Console.WriteLine("usbtools restore --source <wimfile> --target <drive> [options]");
        Console.WriteLine();
        Console.WriteLine("Restore a WIM file to a USB drive, recreating partitions");
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
        Console.WriteLine("Target drive letter (e.g. F:)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -k, --disk <number>       ");
        Console.ResetColor();
        Console.WriteLine("Target disk number (alternative to target)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -p, --provider <type>     ");
        Console.ResetColor();
        Console.WriteLine("Provider: auto (default), wimapi, dism");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -y, --yes                 ");
        Console.ResetColor();
        Console.WriteLine("Confirm all prompts automatically");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -b, --boot-mode <mode>    ");
        Console.ResetColor();
        Console.WriteLine("Boot mode: uefi, bios, auto (default)");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -f, --force-boot          ");
        Console.ResetColor();
        Console.WriteLine("Force bcdboot generation even if boot files exist");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -v, --debug               ");
        Console.ResetColor();
        Console.WriteLine("Enable verbose debug logging");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -h, --help                ");
        Console.ResetColor();
        Console.WriteLine("Show help");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Example:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("  usbtools restore -s C:\\Backups\\usb.usbwim -t F:");
        Console.WriteLine("  usbtools restore -s C:\\Backups\\usb.usbwim --disk 2 --yes");
        Console.ResetColor();
    }

    private static void ShowTestHelp()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.Write("Usage: ");
        Console.ResetColor();
        Console.WriteLine("usbtools test --source <wimfile> [options]");
        Console.WriteLine();
        Console.WriteLine("Verify a WIM image has valid USBTools metadata");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Options:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -s, --source <wimfile>    ");
        Console.ResetColor();
        Console.WriteLine("Source WIM file path");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -v, --debug               ");
        Console.ResetColor();
        Console.WriteLine("Enable verbose debug logging");
        Console.ForegroundColor = ConsoleColor.Green;
        Console.Write("  -h, --help                ");
        Console.ResetColor();
        Console.WriteLine("Show help");
        Console.WriteLine();
        
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("Example:");
        Console.ResetColor();
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("  usbtools test -s C:\\Backups\\usb.usbwim");
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
