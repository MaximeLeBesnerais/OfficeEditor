using System.Text.Json;

namespace PptxEditor.Core.Generation.Fixtures;

/// <summary>
/// Per-primitive RMSE thresholds, loaded from the committed thresholds file
/// (tools/visual-diff/baselines/gen/thresholds.json). The file mirrors
/// <see cref="FixtureCatalog"/>; a drift-guard test keeps the two in sync. Values are
/// absolute normalized-RMSE ceilings (0–1) applied per page and per document average —
/// unlike the REF-deck baselines, which are machine-dependent metrics.json snapshots,
/// thresholds are authored and machine-independent.
/// </summary>
public static class ParityThresholds
{
    /// <summary>Current thresholds file format version.</summary>
    public const int FileVersion = 1;

    /// <summary>Loads and validates a thresholds.json file.</summary>
    /// <exception cref="FileNotFoundException">The file does not exist (operational error, never a silent pass).</exception>
    /// <exception cref="InvalidOperationException">The file is malformed or has invalid values.</exception>
    public static IReadOnlyDictionary<string, double> Load(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        if (!File.Exists(path))
        {
            throw new FileNotFoundException(
                $"Parity thresholds not found: {path}\n"
                + "The file is committed at tools/visual-diff/baselines/gen/thresholds.json and mirrors "
                + "PptxEditor.Core/Generation/Fixtures/FixtureCatalog (drift-guarded by FixtureCatalogTests).",
                path);
        }

        ThresholdsFileDto? dto;
        try
        {
            dto = JsonSerializer.Deserialize<ThresholdsFileDto>(File.ReadAllText(path),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException($"'{path}' is not a valid parity thresholds file: {ex.Message}");
        }

        if (dto?.Thresholds is null)
        {
            throw new InvalidOperationException($"'{path}' does not contain a 'thresholds' object.");
        }

        if (dto.Version != FileVersion)
        {
            throw new InvalidOperationException($"'{path}' has version {dto.Version}; expected {FileVersion}.");
        }

        var thresholds = new Dictionary<string, double>(StringComparer.Ordinal);
        foreach ((string name, double value) in dto.Thresholds)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new InvalidOperationException($"'{path}' contains an empty fixture name.");
            }
            if (value is <= 0 or > 1)
            {
                throw new InvalidOperationException(
                    $"'{path}' threshold for '{name}' is {value}; normalized RMSE thresholds must be in (0, 1].");
            }
            thresholds[name] = value;
        }
        return thresholds;
    }

    /// <summary>Serializes the catalog thresholds to the committed file shape (used by regeneration + drift guard).</summary>
    public static string SerializeCatalog()
    {
        var dto = new ThresholdsFileDto
        {
            Version = FileVersion,
            Thresholds = FixtureCatalog.All.ToDictionary(f => f.Name, f => f.ThresholdRmse, StringComparer.Ordinal)
        };
        return JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });
    }

    private sealed class ThresholdsFileDto
    {
        public int Version { get; set; }
        public Dictionary<string, double>? Thresholds { get; set; }
    }
}
