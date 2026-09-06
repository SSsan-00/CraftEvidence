using CraftEvidence.Core.Models;

namespace CraftEvidence.Core.Services;

/// <summary>
/// Resolves the current Evidence Case without relying on a persisted row number or a fixed slot height.
/// </summary>
public sealed class CaseLayoutAnalyzer
{
  public LayoutAnalysisResult Analyze(SheetLayoutSignals signals)
  {
    ArgumentNullException.ThrowIfNull(signals);
    ArgumentNullException.ThrowIfNull(signals.Anchors);
    ArgumentNullException.ThrowIfNull(signals.VerticalBoundaries);
    ArgumentNullException.ThrowIfNull(signals.HorizontalBoundaries);
    ArgumentNullException.ThrowIfNull(signals.OldHeaderColumns);

    if (signals.ActiveRow is < 1 or > ExcelWorksheetLimits.MaximumRow ||
      signals.RawUsedLastRow is < 1 or > ExcelWorksheetLimits.MaximumRow ||
      signals.LogicalEvidenceLastRow is < 1 or > ExcelWorksheetLimits.MaximumRow ||
      signals.LogicalEvidenceLastRow > signals.RawUsedLastRow ||
      signals.NewFirstColumn is < 1 or > ExcelWorksheetLimits.MaximumColumn ||
      signals.ObservedLastEvidenceColumn is < 1 or > ExcelWorksheetLimits.MaximumColumn)
    {
      return LayoutAnalysisResult.Unsafe("Worksheet coordinates exceed the Excel worksheet limits.");
    }

    var anchors = NormalizeAnchors(signals.Anchors);
    if (anchors.Any(anchor => anchor.Row > signals.LogicalEvidenceLastRow) ||
      signals.VerticalBoundaries.Any(boundary =>
        boundary.RightColumn is < 1 or > ExcelWorksheetLimits.MaximumColumn ||
        boundary.StartRow < 1 ||
        boundary.EndRow < boundary.StartRow ||
        boundary.EndRow > ExcelWorksheetLimits.MaximumRow) ||
      signals.HorizontalBoundaries.Any(boundary =>
        boundary.Row is < 1 or > ExcelWorksheetLimits.MaximumRow ||
        boundary.FirstColumn < 1 ||
        boundary.LastColumn < boundary.FirstColumn ||
        boundary.LastColumn > ExcelWorksheetLimits.MaximumColumn) ||
      signals.OldHeaderColumns.Any(column => column is < 1 or > ExcelWorksheetLimits.MaximumColumn))
    {
      return LayoutAnalysisResult.Unsafe("Layout signals contain invalid or out-of-range worksheet coordinates.");
    }

    var currentAnchor = anchors.LastOrDefault(anchor => anchor.Row <= signals.ActiveRow);
    if (currentAnchor is null)
    {
      return LayoutAnalysisResult.Unsafe("No Case anchor exists at or above the active row.");
    }

    if (currentAnchor.Row != anchors[0].Row && !currentAnchor.HasTopBorder)
    {
      return LayoutAnalysisResult.Unsafe(
        $"The Case anchor at row {currentAnchor.Row} is not confirmed by a structural boundary.");
    }

    var reasons = new List<string> { $"Case anchor resolved at row {currentAnchor.Row}." };
    var confidence = currentAnchor.Row == anchors[0].Row || currentAnchor.HasTopBorder
      ? LayoutConfidence.High
      : LayoutConfidence.Medium;

    if (!currentAnchor.HasTopBorder && currentAnchor.Row != anchors[0].Row)
    {
      reasons.Add("The selected anchor has no top border; A/B values are the primary signal.");
    }

    var nextAnchor = anchors.FirstOrDefault(anchor => anchor.Row > currentAnchor.Row);
    if (nextAnchor is not null && !nextAnchor.HasTopBorder)
    {
      return LayoutAnalysisResult.Unsafe(
        reasons.Append($"The next Case anchor at row {nextAnchor.Row} is not structurally confirmed.").ToArray());
    }

    var endResolution = ResolveCaseEnd(signals, currentAnchor.Row, nextAnchor);
    if (endResolution.EndRow is null)
    {
      return LayoutAnalysisResult.Unsafe(reasons.Append(endResolution.Reason).ToArray());
    }

    var endRow = endResolution.EndRow.Value;
    reasons.Add(endResolution.Reason);
    confidence = Min(confidence, endResolution.Confidence);

    if (nextAnchor is null && endRow != signals.LogicalEvidenceLastRow)
    {
      return LayoutAnalysisResult.Unsafe(
        reasons.Append(
          $"The final Case end row {endRow} disagrees with logical Evidence row {signals.LogicalEvidenceLastRow}.").ToArray());
    }

    if (endRow > signals.LogicalEvidenceLastRow)
    {
      return LayoutAnalysisResult.Unsafe(
        reasons.Append(
          $"The Case end row {endRow} exceeds logical Evidence row {signals.LogicalEvidenceLastRow}.").ToArray());
    }

    if (signals.ActiveRow > endRow)
    {
      return LayoutAnalysisResult.Unsafe(
        reasons.Append("The active row is below the resolved Case boundary.").ToArray());
    }

    var sideResolution = ResolveSideBoundary(signals, currentAnchor.Row, endRow);
    if (sideResolution.NewRightColumn is null || sideResolution.LastEvidenceColumn is null)
    {
      return LayoutAnalysisResult.Unsafe(reasons.Append(sideResolution.Reason).ToArray());
    }

    var newRight = sideResolution.NewRightColumn.Value;
    if (newRight >= ExcelWorksheetLimits.MaximumColumn)
    {
      return LayoutAnalysisResult.Unsafe(
        reasons.Append("The New/Old boundary leaves no valid Excel column for the Old region.").ToArray());
    }

    var oldFirst = newRight + 1;
    var evidenceLast = sideResolution.LastEvidenceColumn.Value;
    if (newRight < signals.NewFirstColumn || evidenceLast < oldFirst)
    {
      return LayoutAnalysisResult.Unsafe(
        reasons.Append("Resolved New/Old column ranges are invalid.").ToArray());
    }

    reasons.Add(sideResolution.Reason);
    confidence = Min(confidence, sideResolution.Confidence);

    var layout = new EvidenceCaseLayout(
      currentAnchor.Row,
      endRow,
      new ColumnRange(signals.NewFirstColumn, newRight),
      new ColumnRange(oldFirst, evidenceLast));

    return new LayoutAnalysisResult(layout, confidence, reasons);
  }

  private static List<CaseAnchorSignal> NormalizeAnchors(IReadOnlyList<CaseAnchorSignal> anchors) =>
    anchors
      .Where(anchor => anchor.Row > 0 && anchor.HasCaseValue)
      .GroupBy(anchor => anchor.Row)
      .Select(group => new CaseAnchorSignal(
        group.Key,
        group.Any(anchor => anchor.HasValueInColumnA),
        group.Any(anchor => anchor.HasValueInColumnB),
        group.Any(anchor => anchor.HasTopBorder)))
      .OrderBy(anchor => anchor.Row)
      .ToList();

  private static EndResolution ResolveCaseEnd(
    SheetLayoutSignals signals,
    int startRow,
    CaseAnchorSignal? nextAnchor)
  {
    if (nextAnchor is not null)
    {
      return new EndResolution(
        nextAnchor.Row - 1,
        LayoutConfidence.High,
        $"Case end resolved from the next anchor at row {nextAnchor.Row}.");
    }

    var horizontalEnds = signals.HorizontalBoundaries
      .Where(boundary => boundary.Row > startRow && boundary.FirstColumn <= signals.NewFirstColumn)
      .Select(boundary => boundary.Row)
      .Distinct()
      .OrderBy(row => row)
      .ToArray();

    var verticalEnds = signals.VerticalBoundaries
      .Where(boundary => boundary.StartRow <= startRow && boundary.EndRow >= startRow)
      .Select(boundary => boundary.EndRow)
      .Distinct()
      .OrderBy(row => row)
      .ToArray();

    var horizontalEnd = horizontalEnds.Length == 1 ? horizontalEnds[0] : (int?)null;
    var verticalEnd = verticalEnds.Length == 1 ? verticalEnds[0] : (int?)null;

    if (horizontalEnds.Length > 1 || verticalEnds.Length > 1)
    {
      return new EndResolution(null, LayoutConfidence.Unsafe, "Multiple final Case boundaries conflict.");
    }

    if (horizontalEnd is not null && verticalEnd is not null)
    {
      return horizontalEnd == verticalEnd
        ? new EndResolution(
          horizontalEnd,
          LayoutConfidence.High,
          $"Final Case end row {horizontalEnd} is confirmed by horizontal and vertical borders.")
        : new EndResolution(null, LayoutConfidence.Unsafe, "Final Case border end signals disagree.");
    }

    if (horizontalEnd is not null || verticalEnd is not null)
    {
      var end = horizontalEnd ?? verticalEnd;
      return new EndResolution(
        end,
        LayoutConfidence.Medium,
        $"Final Case end row {end} is based on one border signal.");
    }

    return new EndResolution(null, LayoutConfidence.Unsafe, "The final Case end cannot be determined safely.");
  }

  private static SideResolution ResolveSideBoundary(
    SheetLayoutSignals signals,
    int startRow,
    int endRow)
  {
    var verticalColumns = signals.VerticalBoundaries
      .Where(boundary => boundary.StartRow <= startRow && boundary.EndRow >= endRow)
      .Select(boundary => boundary.RightColumn)
      .Where(column => column >= signals.NewFirstColumn)
      .Distinct()
      .OrderBy(column => column)
      .ToArray();

    var headerColumns = signals.OldHeaderColumns
      .Where(column => column > signals.NewFirstColumn)
      .Distinct()
      .OrderBy(column => column)
      .ToArray();

    if (verticalColumns.Length > 1 || headerColumns.Length > 1)
    {
      return new SideResolution(null, null, LayoutConfidence.Unsafe, "Multiple New/Old boundaries conflict.");
    }

    var verticalRight = verticalColumns.Length == 1 ? verticalColumns[0] : (int?)null;
    var headerRight = headerColumns.Length == 1 ? headerColumns[0] - 1 : (int?)null;
    if (verticalRight is not null && headerRight is not null && verticalRight != headerRight)
    {
      return new SideResolution(null, null, LayoutConfidence.Unsafe, "The vertical border and Old header disagree.");
    }

    if (verticalRight is null)
    {
      var reason = headerRight is null
        ? "The New/Old boundary cannot be determined safely."
        : "The Old header is not an independent structural signal for the New/Old boundary.";
      return new SideResolution(null, null, LayoutConfidence.Unsafe, reason);
    }


    var newRight = verticalRight.Value;

    var horizontalRightEdges = signals.HorizontalBoundaries
      .Where(boundary => boundary.Row == endRow)
      .Select(boundary => boundary.LastColumn)
      .Distinct()
      .ToArray();
    if (horizontalRightEdges.Length > 1)
    {
      return new SideResolution(null, null, LayoutConfidence.Unsafe, "Multiple Evidence right edges conflict.");
    }

    var observedRight = signals.ObservedLastEvidenceColumn;
    var horizontalRight = horizontalRightEdges.Length == 1 ? horizontalRightEdges[0] : (int?)null;
    if (observedRight is not null && horizontalRight is not null && observedRight != horizontalRight)
    {
      return new SideResolution(null, null, LayoutConfidence.Unsafe, "The observed and horizontal Evidence right edges disagree.");
    }

    var lastEvidenceColumn = observedRight ?? horizontalRight;
    if (lastEvidenceColumn == 0)
    {
      lastEvidenceColumn = null;
    }

    if (lastEvidenceColumn is null)
    {
      return new SideResolution(null, null, LayoutConfidence.Unsafe, "The Evidence right edge cannot be determined safely.");
    }

    return new SideResolution(
      newRight,
      lastEvidenceColumn,
      LayoutConfidence.High,
      $"New right edge resolved from vertical border column {newRight}.");
  }

  private static LayoutConfidence Min(LayoutConfidence left, LayoutConfidence right) =>
    (LayoutConfidence)Math.Min((int)left, (int)right);

  private sealed record EndResolution(int? EndRow, LayoutConfidence Confidence, string Reason);

  private sealed record SideResolution(
    int? NewRightColumn,
    int? LastEvidenceColumn,
    LayoutConfidence Confidence,
    string Reason);
}
