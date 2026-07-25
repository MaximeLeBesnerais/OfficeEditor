namespace PptxEditor.Core.Converters;

/// <summary>
/// Registry of PowerPoint <b>built-in</b> table styles. Built-in style definitions live in
/// the PowerPoint application itself, not in the file: <c>ppt/tableStyles.xml</c> typically
/// carries only the <c>def</c> GUID, so a table referencing a built-in
/// <c>a:tableStyleId</c> resolves to nothing. This registry supplies those definitions as
/// <c>a:tblStyle</c> outer XML so they flow through the exact same parsing pipeline
/// (scheme-color resolution, tint math, border extraction) as file-defined styles.
/// </summary>
/// <remarks>
/// Values for "Medium Style 2 - Accent 1" encode its published, widely-documented
/// definition (the ECMA-376 default table style, also the <c>def</c> of the Office
/// <c>tblStyleLst</c>), cross-checked against PowerPoint's observable rendering:
/// <list type="bullet">
/// <item><c>wholeTbl</c>: 1pt (12700 EMU) solid <c>lt1</c> (white) borders on every edge
/// including insideH/insideV; body text in <c>tx1</c>.</item>
/// <item><c>firstRow</c>: solid <c>accent1</c> fill, bold white (<c>lt1</c>) text; the
/// white wholeTbl borders become invisible against the accent fill.</item>
/// <item><c>lastRow</c>: bold <c>tx1</c> text, 3pt double white top rule. Our border
/// pipeline does not model <c>cmpd="dbl"</c> and renders it solid — visually equivalent
/// on a banded fill.</item>
/// <item><c>firstCol</c>/<c>lastCol</c>: bold <c>tx1</c> text.</item>
/// <item><c>band1H</c>: <c>accent1</c> tint 20% fill; <c>band2H</c>: tint 40% fill.
/// (DrawingML <c>a:tint</c> states how much of the source color is kept, the remainder
/// blending toward white — so tint 20% renders lighter than tint 40%.)</item>
/// </list>
/// Add further GUIDs by pasting the published <c>a:tblStyle</c> XML into
/// <see cref="OuterXmlById"/>; file-defined styles always take precedence over this registry.
/// </remarks>
internal static class BuiltInTableStyles
{
    /// <summary>"Medium Style 2 - Accent 1" — the ECMA-376 default table style GUID.</summary>
    public const string MediumStyle2Accent1Id = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}";

    private const string MediumStyle2Accent1Xml =
        "<a:tblStyle xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        $"styleId=\"{MediumStyle2Accent1Id}\" styleName=\"Medium Style 2 - Accent 1\">" +
        "<a:wholeTbl>" +
        "<a:tcTxStyle><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:tcTxStyle>" +
        "<a:tcStyle><a:tcBdr>" +
        "<a:left><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:left>" +
        "<a:right><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:right>" +
        "<a:top><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:top>" +
        "<a:bottom><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:bottom>" +
        "<a:insideH><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:insideH>" +
        "<a:insideV><a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln></a:insideV>" +
        "</a:tcBdr></a:tcStyle>" +
        "</a:wholeTbl>" +
        "<a:band1H><a:tcStyle>" +
        "<a:fill><a:solidFill><a:schemeClr val=\"accent1\"><a:tint val=\"20000\"/></a:schemeClr></a:solidFill></a:fill>" +
        "</a:tcStyle></a:band1H>" +
        "<a:band2H><a:tcStyle>" +
        "<a:fill><a:solidFill><a:schemeClr val=\"accent1\"><a:tint val=\"40000\"/></a:schemeClr></a:solidFill></a:fill>" +
        "</a:tcStyle></a:band2H>" +
        "<a:firstRow>" +
        "<a:tcTxStyle b=\"1\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:tcTxStyle>" +
        "<a:tcStyle><a:fill><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></a:fill></a:tcStyle>" +
        "</a:firstRow>" +
        "<a:lastRow>" +
        "<a:tcTxStyle b=\"1\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:tcTxStyle>" +
        "<a:tcStyle><a:tcBdr>" +
        "<a:top><a:ln w=\"38100\" cap=\"flat\" cmpd=\"dbl\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"lt1\"/></a:solidFill></a:ln></a:top>" +
        "</a:tcBdr></a:tcStyle>" +
        "</a:lastRow>" +
        "<a:firstCol><a:tcTxStyle b=\"1\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:tcTxStyle></a:firstCol>" +
        "<a:lastCol><a:tcTxStyle b=\"1\"><a:solidFill><a:schemeClr val=\"tx1\"/></a:solidFill></a:tcTxStyle></a:lastCol>" +
        "</a:tblStyle>";

    /// <summary>
    /// Built-in style GUID → <c>a:tblStyle</c> outer XML. Lookup is case-insensitive
    /// (GUID case varies between producers).
    /// </summary>
    public static IReadOnlyDictionary<string, string> OuterXmlById { get; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [MediumStyle2Accent1Id] = MediumStyle2Accent1Xml
        };
}
