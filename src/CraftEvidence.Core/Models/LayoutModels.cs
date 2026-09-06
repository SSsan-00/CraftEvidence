namespace CraftEvidence.Core.Models;

public enum EvidenceSide
{
  New,
  Old,
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
  ColumnRange OldRegion);

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
