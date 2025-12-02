# Restoration Logic Critique

## Partition Resizing and Geometry Handling
- **Elastic vs fixed detection is not used.** `ResizePartitionsForTarget` blindly marks only the last partition for expansion/shrinkage (`Size = -1`) without consulting `PartitionInfo.IsFixed` or `UsedSpace`, so recovery/EFI partitions could be resized and data partitions left unchanged depending on partition order. This ignores the intent of fixed vs. elastic metadata.
- **No guard for smaller targets.** The code does not compare the target disk size with the sum of required partition sizes; if the destination is smaller, partition creation will fail later with diskpart instead of performing proportional shrinking or aborting early.
- **Offset/size metadata is discarded.** Partition offsets from the manifest are not honored during recreation, which can break scenarios that depend on specific alignment/gap requirements.

## Partition Recreation and Bootability
- **Boot-critical identifiers are not preserved.** Disk GUIDs, per-partition GPT type/ID GUIDs, and MBR disk signatures from the manifest are not reapplied; diskpart creates new identifiers and there is no explicit write-back, which can break boot chaining or BitLocker/WinRE bindings.
- **EFI detection is heuristic.** GPT EFI creation is inferred from `GptType` or `IsFixed` + FAT32 instead of the manifest’s explicit type/size, and any partition marked `IsActive` is only honored for MBR. There is no handling of MSR, recovery GUID types, or hidden attributes.
- **Partition-to-image mapping is fragile.** Apply order is based on current drive letter enumeration rather than the captured partition metadata, so any drive-letter assignment differences at restore time can result in WIM images being applied to the wrong partitions.

## WIM Apply Flow
- **Serial WIMLoadImage/WIMApplyImage per partition.** Each image is opened and closed sequentially without streaming or reuse; acceptable for few partitions but could be optimized by validating image count vs. partition count and applying in metadata order.
- **Minimal error handling for temp path and callbacks.** Failure to set the temporary path only logs a warning, and callbacks are registered per image but not globally; failures from `WIMApplyImage` don’t attempt retry or more detailed context.

## Boot Configuration
- **UEFI/BIOS boot state is inferred post-apply.** The code searches for EFI/bootmgr files and then reruns `bcdboot`, but does not ensure the EFI partition size/label/mount point matches the captured layout, nor does it restore ESP GPT attributes. Legacy boot relies on `bootsect` without ensuring the correct partition is active when drive letters differ.

## Recommendations
- Use manifest flags (`IsFixed`, `UsedSpace`) to keep EFI/MSR/Recovery sizes constant and only stretch elastic data partitions. Abort early if the target capacity is insufficient rather than relying on diskpart failure.
- Preserve disk signature, GPT disk GUID, and per-partition GUID/type/attribute fields when recreating the layout. Honor captured offsets/alignment or at least align consistently and log deviations.
- Map WIM images to partitions using manifest indices (e.g., partition index/order) instead of enumerated drive letters to avoid misapplied images.
- Validate image count vs. partition count and ensure boot partitions receive the correct files before running `bcdboot`/`bootsect`.
