# Feature Opportunities for a Reliable, Appealing USB Restorer

The current implementation focuses on correctness and validation (see `RestoreCommand` and the supportability critique in `SUPPORTABILITY_ASSESSMENT.md`). To make the tool more attractive to operators and end users, consider the following additions.

## User Experience and Safety
- **Guided preflight with compatibility scoring:** Run the restore logic in a "dry run" mode that checks target disk capacity, partition-style compatibility, and available providers (WIM API vs. DISM) before any destructive steps, returning a single readiness score with actionable remediation tips.
- **Progressive disclosure UI/CLI:** Add a `--verbose` tree view showing partition creation, image apply progress per volume, and boot repair steps with timestamps, plus a `--quiet` mode for automation logs.
- **One-click rollback point:** Before wiping the disk, capture the existing partition table (GPT/MBR) to a small backup file so operators can revert if the restore is cancelled or fails early.

## Reliability and Speed
- **Resume-friendly WIM apply:** Track per-partition apply state in a checkpoint file so interrupted restores (power loss, USB removal) can resume at the last completed partition without re-imaging the entire drive.
- **Parallel data partition apply:** Allow data-only partitions to apply in parallel while keeping boot partitions serialized, improving throughput on USB 3.x/4.0 devices without risking boot corruption.
- **Hash-based verification:** After apply, compute block hashes for each partition and compare against hashes captured during backup, surfacing corruption early and providing confidence to auditors.

## Hardware and Layout Flexibility
- **Template-based partition mapping:** Offer predefined templates (Windows To Go, Linux live USB, dual-boot) that map captured partitions onto target disks with vendor-specific quirks (e.g., ESP alignment), reducing manual script edits.
- **Secure Boot and BitLocker handling:** Detect BitLocker/TPM binding in metadata; optionally back up and restore protectors or prompt to regenerate them, ensuring the restored device boots without user intervention.
- **Driver/package injection:** Permit injecting updated storage, USB, or network drivers and cumulative updates during apply for better hardware compatibility post-restore.

## Operations and Observability
- **Health metrics and telemetry hooks:** Emit structured logs/events (JSON/ETW) for each major stage (partitioning, apply, boot prep) so fleet operators can aggregate success rates and spot failing device models.
- **Rich post-restore validation report:** Extend `ValidateRestoredLayout` to generate a human-readable report (HTML/Markdown) summarizing partition table, GUIDs, file systems, boot files, and any mismatches, suitable for helpdesk attachments.
- **Pluggable storage provider abstraction:** Generalize the apply path to support future container formats (e.g., ESD, VHDX) without rewriting partition or boot logic, making the tool a long-term platform.

## Distribution and Security
- **Signed, self-contained package:** Bundle required binaries (`wimgapi.dll`, `dism`) and scripts in a signed installer to avoid dependency drift and provide integrity validation for support teams.
- **Configurable policy enforcement:** Allow administrators to define policies (minimum USB speed, required Secure Boot state, allowed partition templates) that block restores violating corporate standards.
- **Audit-friendly manifests:** Save a signed manifest alongside each restore detailing source image, hash, timestamp, operator, and target hardware IDs for compliance traceability.
