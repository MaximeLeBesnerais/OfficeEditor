namespace XlsxEditor.Core.Builders;

/// <summary>
/// Calculation properties written to the workbook's &lt;calcPr&gt; element. Setting
/// <see cref="FullCalcOnLoad"/> tells Excel to recalculate formulas when the file opens
/// — the correct behaviour for formula cells written without a cached value.
/// </summary>
public sealed record WorkbookCalculationProperties
{
    /// <summary>
    /// Recount/recalculate all formulas on load (calcPr fullCalcOnLoad="1"). Defaults to
    /// true because formula cells are written without cached values and must not display
    /// stale or empty results.
    /// </summary>
    public bool FullCalcOnLoad { get; init; } = true;

    /// <summary>
    /// Forces a full recalculation of all formulas on load, ignoring dependencies
    /// (calcPr forceFullCalc="1").
    /// </summary>
    public bool? ForceFullCalc { get; init; }

    /// <summary>
    /// Recalculates dirty cells before saving (calcPr calcOnSave="1").
    /// </summary>
    public bool? CalcOnSave { get; init; }

    /// <summary>
    /// The calcId identifying the calc engine that last calculated the workbook. When
    /// null it is left unset; Excel applies full recalculation on open when
    /// <see cref="FullCalcOnLoad"/> is true regardless of the id.
    /// </summary>
    public uint? CalculationId { get; init; }
}
