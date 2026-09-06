namespace CraftEvidence.App;

internal sealed record CraftEvidenceSettings
{
  internal const double DefaultHorizontalMarginPoints = 6;
  public double HorizontalMarginPoints { get; init; } = DefaultHorizontalMarginPoints;

  public bool DiagnosticLoggingEnabled { get; init; } = true;

  public bool GlobalShortcutEnabled { get; init; } = true;

  internal CraftEvidenceSettings Normalize() => this with
  {
    HorizontalMarginPoints = double.IsFinite(HorizontalMarginPoints)
      ? Math.Clamp(HorizontalMarginPoints, 0, 72)
      : DefaultHorizontalMarginPoints,
  };
}
