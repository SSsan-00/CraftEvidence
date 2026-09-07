using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Moves Excel to a structurally confirmed neighbouring Case anchor.</summary>
public sealed class ExcelCaseNavigationService
{
  private readonly ExcelSheetSnapshotService snapshotService = new();
  private readonly IPlacementFocusService focusService = new ExcelPlacementFocusService();
  private readonly CaseLayoutAnalyzer layoutAnalyzer = new();

  public CaseNavigationResult Navigate(
    WorkbookIdentity workbook,
    string worksheetName,
    CaseNavigationDirection direction,
    string? currentCaseLabel = null,
    EvidenceSide currentSide = EvidenceSide.New,
    bool sameCaseThenNext = true)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return CaseNavigationResult.Failed("Case移動はSTAスレッドで実行する必要があります。");
    }
    if (!string.IsNullOrWhiteSpace(currentCaseLabel) &&
      ExcelAutomaticPlacementService.NormalizeCaseLabel(currentCaseLabel) is null)
    {
      return CaseNavigationResult.Failed("CaseはX-X形式で入力してください（例: 1-2）。");
    }

    var captured = snapshotService.CaptureForNavigation(workbook, worksheetName, includeWorksheetNames: true);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return CaseNavigationResult.Failed(captured.Message);
    }

    var snapshot = captured.Snapshot;
    var anchors = ExcelAutomaticPlacementService.ConfirmedAnchors(snapshot.LayoutSignals);
    var blocks = sameCaseThenNext
      ? anchors.SelectMany(anchor => new[]
        {
          new Block(anchor, EvidenceSide.New),
          new Block(anchor, EvidenceSide.Old),
        }).ToArray()
      : anchors.Select(anchor => new Block(anchor, currentSide)).ToArray();
    var currentIndex = FindCurrentIndex(blocks, snapshot, currentCaseLabel, currentSide);
    var step = direction is CaseNavigationDirection.Previous ? -1 : 1;
    Block? block = null;
    for (var index = currentIndex + step; index >= 0 && index < blocks.Length; index += step)
    {
      if (!IsOccupied(snapshot, blocks[index]))
      {
        block = blocks[index];
        break;
      }
    }
    if (block is null)
    {
      if (direction is CaseNavigationDirection.Next)
      {
        var nextSheet = NavigateNextWorksheet(workbook, snapshot, currentSide, sameCaseThenNext);
        if (nextSheet.Succeeded)
        {
          return nextSheet;
        }
      }

      return CaseNavigationResult.Failed(direction is CaseNavigationDirection.Previous
        ? "前に空いている配置先はありません。"
        : "次に空いている配置先はありません。");
    }

    var layout = layoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = block.Anchor.Row }).Layout;
    if (layout is null)
    {
      return CaseNavigationResult.Failed("次の配置先のレイアウトを安全に解析できません。");
    }
    var firstColumn = block.Side is EvidenceSide.New
      ? layout.NewRegion.FirstColumn
      : layout.OldRegion.FirstColumn;
    var target = new CellReference(layout.StartRow + 2, firstColumn + 1);
    var focused = focusService.FocusPlacedImage(workbook, snapshot.WorksheetName, target);
    var caseLabel = ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor);
    return focused.Succeeded
      ? new CaseNavigationResult(true, snapshot.WorksheetName, caseLabel, block.Side, target,
        $"{snapshot.WorksheetName} / Case {caseLabel} / {block.Side} へ移動しました。")
      : CaseNavigationResult.Failed(focused.Message);
  }

  private CaseNavigationResult NavigateNextWorksheet(
    WorkbookIdentity workbook,
    SheetSnapshot current,
    EvidenceSide currentSide,
    bool sameCaseThenNext)
  {
    if (!TryParseEvidenceSheetNumber(current.WorksheetName, out var currentNumber))
    {
      return CaseNavigationResult.Failed("次のEvidence Sheetを判定できません。");
    }

    var candidates = current.WorksheetNames
      .Select(name => (Name: name, Number: TryParseEvidenceSheetNumber(name, out var number) ? number : -1))
      .Where(item => item.Number > currentNumber)
      .OrderBy(item => item.Number);
    foreach (var candidate in candidates)
    {
      var activated = focusService.FocusPlacedImage(workbook, candidate.Name, new CellReference(1, 1));
      if (!activated.Succeeded)
      {
        continue;
      }

      var captured = snapshotService.CaptureForNavigation(workbook, candidate.Name);
      if (!captured.Succeeded || captured.Snapshot is null)
      {
        continue;
      }

      var snapshot = captured.Snapshot;
      var anchors = ExcelAutomaticPlacementService.ConfirmedAnchors(snapshot.LayoutSignals);
      var blocks = sameCaseThenNext
        ? anchors.SelectMany(anchor => new[] { new Block(anchor, EvidenceSide.New), new Block(anchor, EvidenceSide.Old) })
        : anchors.Select(anchor => new Block(anchor, currentSide));
      foreach (var block in blocks)
      {
        if (IsOccupied(snapshot, block))
        {
          continue;
        }

        var layout = layoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = block.Anchor.Row }).Layout;
        if (layout is null)
        {
          continue;
        }

        var region = block.Side is EvidenceSide.New ? layout.NewRegion : layout.OldRegion;
        var target = new CellReference(layout.StartRow + 2, region.FirstColumn + 1);
        var focused = focusService.FocusPlacedImage(workbook, candidate.Name, target);
        if (focused.Succeeded)
        {
          var caseLabel = ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor);
          return new CaseNavigationResult(true, candidate.Name, caseLabel, block.Side, target,
            $"{candidate.Name} / Case {caseLabel} / {block.Side} へ移動しました。");
        }
      }
    }

    _ = focusService.FocusPlacedImage(workbook, current.WorksheetName, current.ActiveCell);
    return CaseNavigationResult.Failed("次のEvidence Sheetに空いている配置先はありません。");
  }

  private static bool TryParseEvidenceSheetNumber(string name, out int number)
  {
    number = -1;
    return name.Length > 1 && (name[0] is 'B' or 'b') && int.TryParse(name.AsSpan(1), out number);
  }

  private static int FindCurrentIndex(
    IReadOnlyList<Block> blocks,
    SheetSnapshot snapshot,
    string? currentCaseLabel,
    EvidenceSide currentSide)
  {
    var normalizedCaseLabel = string.IsNullOrWhiteSpace(currentCaseLabel)
      ? null
      : ExcelAutomaticPlacementService.NormalizeCaseLabel(currentCaseLabel);
    var index = normalizedCaseLabel is not null
      ? Array.FindIndex(blocks.ToArray(), block =>
        string.Equals(
          ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor),
          normalizedCaseLabel,
          StringComparison.OrdinalIgnoreCase) &&
        block.Side == currentSide)
      : -1;
    if (index >= 0)
    {
      return index;
    }

    index = Array.FindLastIndex(blocks.ToArray(), block => block.Anchor.Row <= snapshot.ActiveCell.Row && block.Side == currentSide);
    return index >= 0 ? index : 0;
  }

  private bool IsOccupied(SheetSnapshot snapshot, Block block)
  {
    var layout = layoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = block.Anchor.Row }).Layout;
    if (layout is null)
    {
      return true;
    }
    var region = block.Side is EvidenceSide.New ? layout.NewRegion : layout.OldRegion;
    return snapshot.Shapes.Any(shape =>
      shape.StartRow <= layout.EndRow && shape.EndRow >= layout.StartRow + 1 &&
      shape.StartColumn <= region.LastColumn && shape.EndColumn >= region.FirstColumn + 1);
  }

  private sealed record Block(CaseAnchorSignal Anchor, EvidenceSide Side);

  public static int? ResolveTargetRow(
    SheetLayoutSignals signals,
    CaseNavigationDirection direction)
  {
    ArgumentNullException.ThrowIfNull(signals);
    var confirmed = ExcelAutomaticPlacementService.ConfirmedAnchors(signals)
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
  string CaseLabel,
  EvidenceSide Side,
  CellReference Target,
  string Message)
{
  public static CaseNavigationResult Failed(string message) =>
    new(false, string.Empty, string.Empty, EvidenceSide.New, default, message);
}
