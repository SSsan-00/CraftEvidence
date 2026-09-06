using CraftEvidence.Core.Models;

namespace CraftEvidence.Excel;

/// <summary>Moves Excel to a structurally confirmed neighbouring Case anchor.</summary>
public sealed class ExcelCaseNavigationService
{
  private readonly ExcelSheetSnapshotService snapshotService = new();
  private readonly IPlacementFocusService focusService = new ExcelPlacementFocusService();

  public CaseNavigationResult Navigate(
    WorkbookIdentity workbook,
    string worksheetName,
    CaseNavigationDirection direction)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return CaseNavigationResult.Failed("Case移動はSTAスレッドで実行する必要があります。");
    }

    var captured = snapshotService.Capture(workbook, worksheetName);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return CaseNavigationResult.Failed(captured.Message);
    }

    var snapshot = captured.Snapshot;
    var row = ResolveTargetRow(snapshot.LayoutSignals, direction);
    if (row is null)
    {
      return CaseNavigationResult.Failed(direction is CaseNavigationDirection.Previous
        ? "前のCaseはありません。"
        : "次のCaseはありません。");
    }

    var target = new CellReference(row.Value, snapshot.LayoutSignals.NewFirstColumn);
    var focused = focusService.FocusPlacedImage(workbook, snapshot.WorksheetName, target);
    return focused.Succeeded
      ? new CaseNavigationResult(true, snapshot.WorksheetName, target, $"{snapshot.WorksheetName}!R{row.Value} のCaseへ移動しました。")
      : CaseNavigationResult.Failed(focused.Message);
  }

  public static int? ResolveTargetRow(
    SheetLayoutSignals signals,
    CaseNavigationDirection direction)
  {
    ArgumentNullException.ThrowIfNull(signals);
    var candidates = signals.Anchors
      .Where(anchor => anchor.HasValueInColumnA || anchor.HasValueInColumnB)
      .OrderBy(anchor => anchor.Row)
      .ToArray();
    var firstAnchorRow = candidates.Length == 0 ? 0 : candidates[0].Row;
    var confirmed = candidates
      .Where(anchor => anchor.Row == firstAnchorRow || anchor.HasTopBorder)
      .Select(anchor => anchor.Row)
      .Distinct()
      .Order()
      .ToArray();
    var currentIndex = Array.FindLastIndex(confirmed, row => row <= signals.ActiveRow);
    var targetIndex = direction is CaseNavigationDirection.Previous ? currentIndex - 1 : currentIndex + 1;
    return targetIndex >= 0 && targetIndex < confirmed.Length ? confirmed[targetIndex] : null;
  }
}

public enum CaseNavigationDirection
{
  Previous,
  Next,
}

public sealed record CaseNavigationResult(
  bool Succeeded,
  string WorksheetName,
  CellReference Target,
  string Message)
{
  public static CaseNavigationResult Failed(string message) => new(false, string.Empty, default, message);
}
