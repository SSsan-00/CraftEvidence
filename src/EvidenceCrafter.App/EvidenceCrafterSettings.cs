namespace EvidenceCrafter.App;

internal sealed record EvidenceCrafterSettings
{
  internal const double DefaultHorizontalMarginPoints = 6;
  public double HorizontalMarginPoints { get; init; } = DefaultHorizontalMarginPoints;

  public bool DiagnosticLoggingEnabled { get; init; } = true;

  public bool GlobalShortcutEnabled { get; init; } = true;

  public bool FollowExcelSelection { get; init; } = true;

  public bool AlwaysOnTop { get; init; }

  public PlacementAdvanceMode AdvanceMode { get; init; } = PlacementAdvanceMode.SameCaseThenNext;

  internal EvidenceCrafterSettings Normalize() => this with
  {
    HorizontalMarginPoints = double.IsFinite(HorizontalMarginPoints)
      ? Math.Clamp(HorizontalMarginPoints, 0, 72)
      : DefaultHorizontalMarginPoints,
  };
}

internal enum PlacementAdvanceMode
{
  SameCaseThenNext,
  NextCaseSameSide,
}
