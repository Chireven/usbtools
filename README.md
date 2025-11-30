# USB Tools

A C# .NET console application for imaging bootable USB drives to WIM format and restoring them to drives of varying sizes.

## Description

This tool creates snapshots of bootable USB drives into a single `.wim` file and restores them to new drives while preserving boot capabilities. It intelligently handles drives of different sizes by analyzing partition geometry and adapting accordingly.

## Features

### Backup (USB → WIM)
- **Multi-partition capture**: All partitions stored as separate indexes in a single WIM
- **Metadata preservation**: Drive geometry (MBR/GPT, partition layout, labels) saved as JSON metadata
- **Compression options**: None, Fast (XPRESS), or Max (LZX)
- **Smart exclusions**: Automatically excludes `$Recycle.Bin` and `System Volume Information`
- **Dual API support**: Primary wimgapi.dll with DISM.exe fallback

### Restore (WIM → USB)
- **Smart geometry adaptation**: 
  - Larger targets: Expands variable partitions to use available space
  - Smaller targets: Attempts to shrink variable partitions (with validation)
- **Bootability preservation**: Automatic boot configuration for MBR and GPT
- **Collision prevention**: Generates new random disk signatures
- **Partition recreation**: Rebuilds partition table with proper types and flags

## Requirements

- .NET 8.0 or later
- Windows OS (wimgapi.dll or DISM.exe)
- **Administrator privileges required**

## Architecture

- **WIM API Wrapper**: P/Invoke for wimgapi.dll operations
- **DISM Fallback**: Automatic fallback to DISM.exe if wimgapi unavailable
- **Drive Analysis**: PowerShell-based geometry detection
- **Logging**: Dual console and file logging with color-coded output

## Installation

```bash
# Clone the repository
git clone https://github.com/Chireven/usbtools.git
cd usbtools

# Build the project
dotnet build -c Release

# The executable will be in: bin\Release\net8.0\usbtools.exe
```

## Usage

### Backup Command

Capture a USB drive to a WIM image:

```bash
# Basic backup with default (fast) compression
usbtools backup --source E: --destination "C:\Backups\usb_backup.wim"

# With maximum compression
usbtools backup -s E: -d "C:\Backups\usb_backup.wim" --compression max

# No compression
usbtools backup -s E: -d "C:\Backups\usb_backup.wim" -c none
```

**Parameters:**
- `--source` or `-s`: Source USB drive letter (e.g., `E:`)
- `--destination` or `-d`: Destination WIM file path
- `--compression` or `-c`: Compression level (`none`, `fast`, `max`) - Default: `fast`

### Restore Command

Restore a WIM image to a USB drive:

```bash
# Restore WIM to target drive
usbtools restore --source "C:\Backups\usb_backup.wim" --target F:

# Short form
usbtools restore -s "C:\Backups\usb_backup.wim" -t F:
```

**Parameters:**
- `--source` or `-s`: Source WIM file path
- `--target` or `-t`: Target USB drive letter (e.g., `F:`)

### Help

```bash
# General help
usbtools --help

# Command-specific help
usbtools backup --help
usbtools restore --help
```

## How It Works

### Backup Process
1. Analyzes source drive partition table (MBR/GPT)
2. Catalogs all partitions with metadata (size, type, flags, file system)
3. Identifies fixed (EFI, Boot, Recovery) vs. variable (Data) partitions
4. Captures each partition as a separate index in the WIM
5. Stores geometry metadata as JSON in the WIM description field

### Restore Process
1. Reads geometry metadata from WIM
2. Analyzes target drive capacity
3. Recreates partition table with new disk signature
4. Adapts partition sizes based on target capacity:
   - Fixed partitions: Restored at original size
   - Variable partitions: Expanded/shrunk to fit
5. Formats partitions with original file systems and labels
6. Applies WIM images to respective partitions
7. Configures boot loader (bcdboot) for MBR or UEFI

## Logging

All operations are logged to:
- **Console**: Color-coded by severity
- **File**: `usbtools.log` in the application directory

Log levels: Debug, Info, Warning, Error

## Technical Details

- **Target Framework**: .NET 8.0
- **CLI Library**: System.CommandLine
- **Admin Manifest**: Embedded for privilege elevation
- **P/Invoke**: Direct wimgapi.dll integration
- **Fallback**: DISM.exe wrapper for compatibility

## License

TBD

## Author

Chireven

## Notes

- Always run with Administrator privileges
- Target drive data will be completely erased during restore
- Ensure adequate disk space for WIM files (size varies by compression)
- Boot configuration assumes Windows-based bootable USBs
