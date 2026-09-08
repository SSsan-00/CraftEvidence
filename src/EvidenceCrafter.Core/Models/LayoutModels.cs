namespace EvidenceCrafter.Core.Models;

public enum EvidenceSide
{
  New,
  Old,
}

public enum SideLayoutKind
{
  NewOnly,
  Both,
}

public enum LayoutConfidence
{
  Unsafe = 0,
  Medium = 1,
  High = 2,
}

public readonly record struct ColumnRange(int FirstColumn, int LastColumn)
{
  public int Count => LastColumn - FirstColumn + 1;
}

public sealed record EvidenceCaseLayout(
  int StartRow,
  int EndRow,
  ColumnRange NewRegion,
  ColumnRange? OldRegion)
{
  public SideLayoutKind Kind => OldRegion is null ? SideLayoutKind.NewOnly : SideLayoutKind.Both;
  public int LastEvidenceColumn => OldRegion?.LastColumn ?? NewRegion.LastColumn;
  public bool CanDeleteTrailingRows { get; init; } = true;
  public bool SupportsSide(EvidenceSide side) =>
    side == EvidenceSide.New || side == EvidenceSide.Old && OldRegion is not null;
  public ColumnRange RegionFor(EvidenceSide side) => side switch
  {
    EvidenceSide.New => NewRegion,
    EvidenceSide.Old when OldRegion is { } old => old,
    _ => throw new InvalidOperationException($"The {side} side is unavailable in this Case layout."),
  };
}

/// <summary>Returns both a layout and the reasons used to accept or reject it.</summary>
public sealed record LayoutAnalysisResult(
  EvidenceCaseLayout? Layout,
  LayoutConfidence Confidence,
  IReadOnlyList<string> Reasons)
{
  public bool IsSafe => Layout is not null && Confidence is not LayoutConfidence.Unsafe;

  public static LayoutAnalysisResult Unsafe(params string[] reasons) =>
    new(null, LayoutConfidence.Unsafe, reasons);
}
