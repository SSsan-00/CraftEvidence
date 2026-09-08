using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Core.Services;

/// <summary>Resolves Case regions from numbering and Side headers; borders are optional diagnostics.</summary>
public sealed class CaseLayoutAnalyzer
{
  public LayoutAnalysisResult Analyze(SheetLayoutSignals signals)
  {
    ArgumentNullException.ThrowIfNull(signals);
    ArgumentNullException.ThrowIfNull(signals.Anchors);
    ArgumentNullException.ThrowIfNull(signals.VerticalBoundaries);
    ArgumentNullException.ThrowIfNull(signals.HorizontalBoundaries);
    ArgumentNullException.ThrowIfNull(signals.OldHeaderColumns);
    ArgumentNullException.ThrowIfNull(signals.NewHeaderColumns);

    if (signals.ActiveRow is < 1 or > ExcelWorksheetLimits.MaximumRow ||
      signals.RawUsedLastRow is < 1 or > ExcelWorksheetLimits.MaximumRow ||
      signals.LogicalEvidenceLastRow is < 1 or > ExcelWorksheetLimits.MaximumRow ||
      signals.LogicalEvidenceLastRow > signals.RawUsedLastRow ||
      signals.NewFirstColumn is < 1 or > ExcelWorksheetLimits.MaximumColumn ||
      signals.ObservedLastEvidenceColumn is < 1 or > ExcelWorksheetLimits.MaximumColumn)
    {
      return LayoutAnalysisResult.Unsafe("Worksheet coordinates exceed the Excel worksheet limits.");
    }

    if (signals.Anchors.Any(anchor => anchor.Row < 1 || anchor.Row > signals.LogicalEvidenceLastRow) ||
      signals.VerticalBoundaries.Any(boundary =>
        boundary.RightColumn is < 1 or > ExcelWorksheetLimits.MaximumColumn ||
        boundary.StartRow < 1 || boundary.EndRow < boundary.StartRow ||
        boundary.EndRow > ExcelWorksheetLimits.MaximumRow) ||
      signals.HorizontalBoundaries.Any(boundary =>
        boundary.Row is < 1 or > ExcelWorksheetLimits.MaximumRow ||
        boundary.FirstColumn < 1 || boundary.LastColumn < boundary.FirstColumn ||
        boundary.LastColumn > ExcelWorksheetLimits.MaximumColumn) ||
      signals.NewHeaderColumns.Concat(signals.OldHeaderColumns)
        .Any(column => column is < 1 or > ExcelWorksheetLimits.MaximumColumn))
    {
      return LayoutAnalysisResult.Unsafe("Layout signals contain invalid or out-of-range worksheet coordinates.");
    }

    var anchors = CaseAnchorNormalizer.ConfirmedAnchors(signals);
    var duplicate = anchors.GroupBy(anchor => $"{anchor.ColumnAValue}-{anchor.ColumnBValue}")
      .FirstOrDefault(group => group.Count() > 1);
    if (duplicate is not null)
    {
      return LayoutAnalysisResult.Unsafe(
        $"Case '{duplicate.Key}' is duplicated at rows {string.Join(", ", duplicate.Select(anchor => anchor.Row))}.");
    }

    if (anchors.GroupBy(anchor => anchor.Row).Any(group => group.Count() > 1))
    {
      return LayoutAnalysisResult.Unsafe("Multiple Case numbers occupy the same anchor row.");
    }

    var current = anchors.LastOrDefault(anchor => anchor.Row <= signals.ActiveRow);
    if (current is null)
    {
      return LayoutAnalysisResult.Unsafe("No Case anchor with numeric X-X numbering exists at or above the active row.");
    }

    var next = anchors.FirstOrDefault(anchor => anchor.Row > current.Row);
    var endRow = next is null ? signals.LogicalEvidenceLastRow : next.Row - 1;
    if (signals.ActiveRow > endRow)
    {
      return LayoutAnalysisResult.Unsafe("The active row is below the resolved Case boundary.");
    }

    var newHeaders = signals.NewHeaderColumns.Distinct().ToArray();
    var oldHeaders = signals.OldHeaderColumns.Distinct().ToArray();
    if (newHeaders.Length != 1 || newHeaders[0] != signals.NewFirstColumn || oldHeaders.Length > 1 ||
      oldHeaders.Any(column => column <= signals.NewFirstColumn))
    {
      return LayoutAnalysisResult.Unsafe("Side headers require one New header at the Evidence start and at most one Old header to its right.");
    }

    if (signals.ObservedLastEvidenceColumn is not { } lastColumn)
    {
      return LayoutAnalysisResult.Unsafe("The Evidence right edge cannot be determined safely; border right edges cannot replace it.");
    }

    var newRegion = new ColumnRange(signals.NewFirstColumn, oldHeaders.Length == 0 ? lastColumn : oldHeaders[0] - 1);
    ColumnRange? oldRegion = oldHeaders.Length == 0 ? null : new ColumnRange(oldHeaders[0], lastColumn);
    if (newRegion.Count < 2 || oldRegion is { Count: < 2 })
    {
      return LayoutAnalysisResult.Unsafe("Each available Side must contain at least two columns for the image inset.");
    }

    var reasons = new List<string>
    {
      $"Case {current.ColumnAValue}-{current.ColumnBValue} resolved at row {current.Row}.",
      next is null
        ? $"Final Case ends at logical Evidence row {endRow}; automatic trailing-row deletion is disabled."
        : $"Case end resolved from the next anchor at row {next.Row}.",
      oldRegion is null ? "Side headers identify New only." : "Side headers identify New/Old.",
    };
    if (signals.VerticalBoundaries.Any(boundary =>
      boundary.StartRow <= current.Row && boundary.EndRow >= current.Row &&
      (boundary.RightColumn != newRegion.LastColumn || boundary.EndRow < endRow)) ||
      signals.HorizontalBoundaries.Any(boundary =>
        boundary.Row > current.Row && boundary.Row <= endRow &&
        (boundary.Row != endRow || boundary.LastColumn != lastColumn)))
    {
      reasons.Add("Warning: border signals disagree with numbering or Side headers; numbering and headers take precedence.");
    }

    return new LayoutAnalysisResult(
      new EvidenceCaseLayout(current.Row, endRow, newRegion, oldRegion)
      {
        CanDeleteTrailingRows = next is not null,
      },
      LayoutConfidence.High,
      reasons);
  }
}
