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
    var blocks = AvailableBlocks(snapshot, currentSide, sameCaseThenNext);
    if (blocks.Length == 0)
    {
      return CaseNavigationResult.Failed("Case番号・新／旧ヘッダーから配置先を解析できません。");
    }
    if (blocks[0].Layout.Kind == SideLayoutKind.NewOnly) currentSide = EvidenceSide.New;
    if (!string.IsNullOrWhiteSpace(currentCaseLabel) && !blocks.Any(block =>
      ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor) == ExcelAutomaticPlacementService.NormalizeCaseLabel(currentCaseLabel)))
    {
      return CaseNavigationResult.Failed("指定Caseが見つかりません。配置先を確認してください。");
    }
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

    var layout = block.Layout;
    var firstColumn = layout.RegionFor(block.Side).FirstColumn;
    var target = new CellReference(layout.StartRow + 2, firstColumn + 1);
    var focused = focusService.FocusPlacedImage(workbook, snapshot.WorksheetName, target);
    var caseLabel = ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor);
    return focused.Succeeded
      ? new CaseNavigationResult(true, snapshot.WorksheetName, caseLabel, block.Side, target,
        $"{snapshot.WorksheetName} / Case {caseLabel} / {block.Side} へ移動しました。")
        { LayoutSignals = snapshot.LayoutSignals, LayoutKind = layout.Kind }
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
      var blocks = AvailableBlocks(snapshot, currentSide, sameCaseThenNext);
      foreach (var block in blocks)
      {
        if (IsOccupied(snapshot, block))
        {
          continue;
        }

        var layout = block.Layout;
        var region = layout.RegionFor(block.Side);
        var target = new CellReference(layout.StartRow + 2, region.FirstColumn + 1);
        var focused = focusService.FocusPlacedImage(workbook, candidate.Name, target);
        if (focused.Succeeded)
        {
          var caseLabel = ExcelAutomaticPlacementService.FormatCaseLabel(block.Anchor);
          return new CaseNavigationResult(true, candidate.Name, caseLabel, block.Side, target,
            $"{candidate.Name} / Case {caseLabel} / {block.Side} へ移動しました。")
            { LayoutSignals = snapshot.LayoutSignals, LayoutKind = layout.Kind };
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

  private static bool IsOccupied(SheetSnapshot snapshot, Block block)
  {
    var layout = block.Layout;
    var region = layout.RegionFor(block.Side);
    return snapshot.Shapes.Any(shape =>
      shape.StartRow <= layout.EndRow && shape.EndRow >= layout.StartRow + 1 &&
      shape.StartColumn <= region.LastColumn && shape.EndColumn >= region.FirstColumn + 1);
  }

  private Block[] AvailableBlocks(SheetSnapshot snapshot, EvidenceSide side, bool sameCaseThenNext)
  {
    var result = new List<Block>();
    foreach (var anchor in ExcelAutomaticPlacementService.ConfirmedAnchors(snapshot.LayoutSignals))
    {
      var layout = layoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = anchor.Row }).Layout;
      if (layout is null) continue;
      if (sameCaseThenNext || layout.Kind == SideLayoutKind.NewOnly)
      {
        result.Add(new Block(anchor, EvidenceSide.New, layout));
        if (sameCaseThenNext && layout.SupportsSide(EvidenceSide.Old))
          result.Add(new Block(anchor, EvidenceSide.Old, layout));
      }
      else result.Add(new Block(anchor, side, layout));
    }
    return result.ToArray();
  }

  private sealed record Block(CaseAnchorSignal Anchor, EvidenceSide Side, EvidenceCaseLayout Layout);

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
  public SideLayoutKind LayoutKind { get; init; }
  public SheetLayoutSignals? LayoutSignals { get; init; }
  public static CaseNavigationResult Failed(string message) =>
    new(false, string.Empty, string.Empty, EvidenceSide.New, default, message);
}
