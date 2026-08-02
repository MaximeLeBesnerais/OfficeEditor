namespace XlsxEditor.Core.Builders;

/// <summary>
/// Core (Dublin Core) package properties of a workbook: the authoring metadata stored
/// in the package's /docProps/core.xml part. All properties are optional; setting a
/// null property clears it (the set is a wholesale replacement, not a merge).
/// </summary>
public sealed record WorkbookCoreProperties
{
    /// <summary>The workbook's display title.</summary>
    public string? Title { get; init; }

    /// <summary>The workbook's subject/topic.</summary>
    public string? Subject { get; init; }

    /// <summary>The principal author (dc:creator).</summary>
    public string? Creator { get; init; }

    /// <summary>Keywords, space-separated (dc:subject-compatible list).</summary>
    public string? Keywords { get; init; }

    /// <summary>A free-text description of the workbook's contents.</summary>
    public string? Description { get; init; }

    /// <summary>The workbook's category.</summary>
    public string? Category { get; init; }

    /// <summary>The name of the last person to modify the workbook.</summary>
    public string? LastModifiedBy { get; init; }

    /// <summary>Creation timestamp (dcterms:created, UTC).</summary>
    public DateTime? Created { get; init; }

    /// <summary>Last-modified timestamp (dcterms:modified, UTC).</summary>
    public DateTime? Modified { get; init; }
}
