namespace XlsxEditor.Core.Rendering.Models;

public enum XlsxValueKind
{
    Empty,
    Number,
    Text,
    Boolean,
    DateTime,
    Error
}

public enum XlsxReadIssueSeverity
{
    Info,
    Warning
}

public sealed record XlsxRenderValue
{
    public XlsxValueKind Kind { get; init; }

    public double? NumberValue { get; init; }

    public string? TextValue { get; init; }

    public bool? BooleanValue { get; init; }

    public DateTime? DateTimeValue { get; init; }

    public string? ErrorValue { get; init; }

    public static XlsxRenderValue Empty() => new() { Kind = XlsxValueKind.Empty };

    public static XlsxRenderValue Number(double value) => new() { Kind = XlsxValueKind.Number, NumberValue = value };

    public static XlsxRenderValue Text(string value) => new() { Kind = XlsxValueKind.Text, TextValue = value };

    public static XlsxRenderValue Boolean(bool value) => new() { Kind = XlsxValueKind.Boolean, BooleanValue = value };

    public static XlsxRenderValue DateTime(DateTime value, double serial) =>
        new() { Kind = XlsxValueKind.DateTime, DateTimeValue = value, NumberValue = serial };

    public static XlsxRenderValue Error(string value) => new() { Kind = XlsxValueKind.Error, ErrorValue = value };
}

public sealed record XlsxRenderBorderEdge
{
    public string? Style { get; init; }

    public string? ColorArgb { get; init; }
}

public sealed record XlsxRenderStyle
{
    public string? FontName { get; init; }

    public double? FontSizePt { get; init; }

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public string? FontColorArgb { get; init; }

    public string? FillArgb { get; init; }

    public XlsxRenderBorderEdge? BorderLeft { get; init; }

    public XlsxRenderBorderEdge? BorderRight { get; init; }

    public XlsxRenderBorderEdge? BorderTop { get; init; }

    public XlsxRenderBorderEdge? BorderBottom { get; init; }

    public string? HorizontalAlignment { get; init; }

    public string? VerticalAlignment { get; init; }

    public bool WrapText { get; init; }

    public string? NumberFormatCode { get; init; }
}

public sealed record XlsxRenderCell
{
    public int Row { get; init; }

    public int Column { get; init; }

    public XlsxRenderValue Value { get; init; } = XlsxRenderValue.Empty();

    public string? Formula { get; init; }

    public XlsxRenderStyle? Style { get; init; }

    public string? NumberFormatCode { get; init; }
}

public sealed record XlsxRenderRow
{
    public int Index { get; init; }

    public double HeightPt { get; init; }

    public bool Hidden { get; init; }

    public IReadOnlyList<XlsxRenderCell> Cells { get; init; } = Array.Empty<XlsxRenderCell>();
}

public sealed record XlsxRenderColumn
{
    public int Index { get; init; }

    public double WidthPt { get; init; }

    public bool Hidden { get; init; }
}

public sealed record XlsxRenderMerge
{
    public int FirstRow { get; init; }

    public int LastRow { get; init; }

    public int FirstCol { get; init; }

    public int LastCol { get; init; }
}

public sealed record XlsxRenderFreezePanes
{
    public int FrozenRows { get; init; }

    public int FrozenCols { get; init; }
}

public sealed record XlsxRenderTable
{
    public string Name { get; init; } = string.Empty;

    public int FirstRow { get; init; }

    public int LastRow { get; init; }

    public int FirstCol { get; init; }

    public int LastCol { get; init; }
}

public sealed record XlsxRenderSheet
{
    public string Name { get; init; } = string.Empty;

    public double PageWidthPt { get; init; }

    public double PageHeightPt { get; init; }

    public bool Landscape { get; init; }

    public IReadOnlyList<XlsxRenderColumn> Columns { get; init; } = Array.Empty<XlsxRenderColumn>();

    public IReadOnlyList<XlsxRenderRow> Rows { get; init; } = Array.Empty<XlsxRenderRow>();

    public IReadOnlyList<XlsxRenderMerge> Merges { get; init; } = Array.Empty<XlsxRenderMerge>();

    public XlsxRenderFreezePanes? FreezePanes { get; init; }

    public IReadOnlyList<XlsxRenderTable> Tables { get; init; } = Array.Empty<XlsxRenderTable>();

    public string? PrintArea { get; init; }

    public string? AutoFilterRange { get; init; }
}

public sealed record XlsxRenderWorkbook
{
    public string? Title { get; init; }

    public IReadOnlyList<XlsxRenderSheet> Sheets { get; init; } = Array.Empty<XlsxRenderSheet>();
}

public sealed record XlsxReadIssue
{
    public XlsxReadIssueSeverity Severity { get; init; }

    public string Message { get; init; } = string.Empty;

    public string? Location { get; init; }
}

public sealed record XlsxReadResult
{
    public XlsxRenderWorkbook Workbook { get; init; } = new();

    public IReadOnlyList<XlsxReadIssue> Issues { get; init; } = Array.Empty<XlsxReadIssue>();
}
