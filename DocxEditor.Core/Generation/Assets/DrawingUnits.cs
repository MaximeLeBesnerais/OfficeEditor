namespace DocxEditor.Core.Generation.Assets;

/// <summary>
/// Drawing/Office metric conversions. OpenXML drawing coordinates are English Metric Units
/// (EMU): 914400 per inch, 12700 per point. Points are the vocabulary's display unit.
/// </summary>
public static class DrawingUnits
{
    /// <summary>EMU per point.</summary>
    public const long EmuPerPoint = 12700;

    /// <summary>EMU per inch.</summary>
    public const long EmuPerInch = 914400;

    /// <summary>Points per inch.</summary>
    public const double PointsPerInch = 72;

    /// <summary>Converts a point measure to EMU, rounding half away from zero.</summary>
    public static long ToEmu(double points) =>
        (long)Math.Round(points * EmuPerPoint, MidpointRounding.AwayFromZero);

    /// <summary>Converts an EMU measure to points.</summary>
    public static double ToPoints(long emu) => emu / (double)EmuPerPoint;
}
