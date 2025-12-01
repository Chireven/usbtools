using System.Text.Json;
using System.Text.Json.Serialization;

namespace USBTools;

public class DriveGeometry
{
    [JsonPropertyName("toolVersion")]
    public string ToolVersion { get; set; } = "1.0.0";

    [JsonPropertyName("partitionStyle")]
    public string PartitionStyle { get; set; } = ""; // MBR or GPT

    [JsonPropertyName("diskSignature")]
    public uint DiskSignature { get; set; }

    [JsonPropertyName("gptDiskId")]
    public string? GptDiskId { get; set; }

    [JsonPropertyName("totalSize")]
    public long TotalSize { get; set; }

    [JsonPropertyName("partitions")]
    public List<PartitionInfo> Partitions { get; set; } = new();

    public string ToJson()
    {
        return JsonSerializer.Serialize(this, new JsonSerializerOptions 
        { 
            WriteIndented = false 
        });
    }

    public static DriveGeometry? FromJson(string json)
    {
        return JsonSerializer.Deserialize<DriveGeometry>(json);
    }
}

public class PartitionInfo
{
    [JsonPropertyName("index")]
    public int Index { get; set; }

    [JsonPropertyName("letter")]
    public string? Letter { get; set; }

    [JsonPropertyName("label")]
    public string? Label { get; set; }

    [JsonPropertyName("fileSystem")]
    public string? FileSystem { get; set; }

    [JsonPropertyName("size")]
    public long Size { get; set; }

    [JsonPropertyName("usedSpace")]
    public long UsedSpace { get; set; }

    [JsonPropertyName("offset")]
    public long Offset { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; } = ""; // System, Boot, Primary, etc.

    [JsonPropertyName("isActive")]
    public bool IsActive { get; set; }

    [JsonPropertyName("isFixed")]
    public bool IsFixed { get; set; } // True for EFI, Boot partitions

    [JsonPropertyName("gptType")]
    public string? GptType { get; set; }

    [JsonPropertyName("gptId")]
    public string? GptId { get; set; }
}
