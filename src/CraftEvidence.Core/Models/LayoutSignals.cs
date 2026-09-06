namespace CraftEvidence.Core.Models;

/// <summary>Describes one possible Evidence Case anchor row.</summary>
public sealed record CaseAnchorSignal(
  int Row,
  bool HasValueInColumnA,
  bool HasValueInColumnB,
  bool HasTopBorder,
  string? ColumnAValue = null,
  string? ColumnBValue = null)
{
  public bool HasCaseValue => HasValueInColumnA || HasValueInColumnB;
}

/// <summary>Describes a vertical right-edge signal observed in a worksheet.</summary>
public sealed record VerticalBoundarySignal(int RightColumn, int StartRow, int EndRow);

/// <summary>Describes a full-width horizontal boundary observed in a worksheet.</summary>
public sealed record HorizontalBoundarySignal(int Row, int FirstColumn, int LastColumn);

/// <summary>
/// Immutable, Excel-free signals used to determine the active Evidence Case and New/Old regions.
/// </summary>
public sealed record SheetLayoutSignals(
  int ActiveRow,
  int RawUsedLastRow,
  int LogicalEvidenceLastRow,
  int NewFirstColumn,
  int? ObservedLastEvidenceColumn,
  IReadOnlyList<CaseAnchorSignal> Anchors,
  IReadOnlyList<VerticalBoundarySignal> VerticalBoundaries,
  IReadOnlyList<HorizontalBoundarySignal> HorizontalBoundaries,
  IReadOnlyList<int> OldHeaderColumns);
