using System.Globalization;
using System.Text;
using System.Text.Json;

namespace VisualDiff;

internal sealed record MetricsReport(DateTimeOffset GeneratedAt, int Dpi, string OutputDirectory, List<DocumentMetrics> Documents);
internal sealed record DocumentMetrics(string Name, string ReferencePdf, string GeneratedPdf, string OutputDirectory, int ReferencePageCount, int GeneratedPageCount, bool PageCountMismatch, List<PageMetrics> Pages);
internal sealed record PageMetrics(int PageNumber, string ReferenceImage, string GeneratedImage, string DiffImage, double? Rmse, double? NormalizedRmse, double? PercentRmse, int CompareExitCode, string RawMetric, string? Error);

/// <summary>
/// JSON (metrics.json) and HTML (index.html) report persistence, plus loading
/// metrics.json-shaped files back for threshold checks. The JSON shape is the
/// contract for baselines under tools/visual-diff/baselines/ — do not rename
/// record members without migrating baselines.
/// </summary>
internal static class ReportWriter
{
    public static void WriteJson<T>(string path, T value)
    {
        JsonSerializerOptions options = new() { WriteIndented = true };
        File.WriteAllText(path, JsonSerializer.Serialize(value, options));
    }

    public static MetricsReport LoadReport(string path)
    {
        string json;
        try
        {
            json = File.ReadAllText(path);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new InvalidOperationException($"Cannot read metrics file '{path}': {ex.Message}");
        }

        try
        {
            return JsonSerializer.Deserialize<MetricsReport>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
                ?? throw new InvalidOperationException($"Metrics file '{path}' deserialized to null.");
        }
        catch (JsonException ex)
        {
            throw new InvalidOperationException($"'{path}' is not a valid visual-diff metrics.json: {ex.Message}");
        }
    }

    public static void WriteHtml(string path, MetricsReport report)
    {
        StringBuilder html = new();
        html.AppendLine("<!doctype html>");
        html.AppendLine("<html lang=\"en\"><head><meta charset=\"utf-8\"><title>visual diff</title>");
        html.AppendLine("<style>body{font-family:sans-serif;margin:2rem}table{border-collapse:collapse;width:100%;margin-bottom:2rem}th,td{border:1px solid #ddd;padding:.5rem;vertical-align:top}img{max-width:220px;border:1px solid #ccc}.warn{color:#9a5b00;font-weight:bold}.err{color:#b00020;font-weight:bold}</style>");
        html.AppendLine("</head><body>");
        html.AppendLine("<h1>visual diff</h1>");
        html.AppendLine($"<p>Generated: {Escape(report.GeneratedAt.ToString("u", CultureInfo.InvariantCulture))} · DPI: {report.Dpi} (PDF inputs only)</p>");

        foreach (DocumentMetrics document in report.Documents)
        {
            html.AppendLine($"<h2>{Escape(document.Name)}</h2>");
            html.AppendLine($"<p>Reference pages: {document.ReferencePageCount}; Generated pages: {document.GeneratedPageCount}</p>");
            if (document.PageCountMismatch)
            {
                html.AppendLine("<p class=\"warn\">Page count mismatch: only common pages were compared.</p>");
            }

            html.AppendLine("<table><thead><tr><th>Page</th><th>Reference</th><th>Generated</th><th>Diff</th><th>Metrics</th></tr></thead><tbody>");
            foreach (PageMetrics page in document.Pages)
            {
                html.AppendLine("<tr>");
                html.AppendLine($"<td>{page.PageNumber}</td>");
                html.AppendLine(ImageCell(path, page.ReferenceImage));
                html.AppendLine(ImageCell(path, page.GeneratedImage));
                html.AppendLine(ImageCell(path, page.DiffImage));
                html.AppendLine("<td>");
                html.AppendLine($"RMSE: {FormatNullable(page.Rmse)}<br>");
                html.AppendLine($"Normalized: {FormatNullable(page.NormalizedRmse)}<br>");
                html.AppendLine($"Percent: {FormatNullable(page.PercentRmse)}%<br>");
                html.AppendLine($"Raw: {Escape(page.RawMetric)}");
                if (page.Error is not null)
                {
                    html.AppendLine($"<div class=\"err\">{Escape(page.Error)}</div>");
                }
                html.AppendLine("</td></tr>");
            }

            html.AppendLine("</tbody></table>");
        }

        html.AppendLine("</body></html>");
        File.WriteAllText(path, html.ToString());
    }

    private static string ImageCell(string htmlPath, string imagePath)
    {
        string relative = Path.GetRelativePath(Path.GetDirectoryName(htmlPath)!, imagePath).Replace(Path.DirectorySeparatorChar, '/');
        string escaped = Escape(relative);
        return $"<td><a href=\"{escaped}\"><img src=\"{escaped}\" alt=\"{escaped}\"></a></td>";
    }

    private static string FormatNullable(double? value) => value?.ToString("0.######", CultureInfo.InvariantCulture) ?? "n/a";

    private static string Escape(string value) => System.Net.WebUtility.HtmlEncode(value);
}
