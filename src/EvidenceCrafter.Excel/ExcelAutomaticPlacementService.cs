using System.Security.Cryptography;
using System.Text;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Excel;

/// <summary>Connects a live worksheet snapshot to Case analysis, placement planning and verified Excel mutations.</summary>
public sealed class ExcelAutomaticPlacementService
{
  private readonly ExcelSheetSnapshotService snapshotService;
  private readonly CaseLayoutAnalyzer layoutAnalyzer;
  private readonly ContentOccupancyAnalyzer occupancyAnalyzer;
  private readonly PlacementPlanner placementPlanner;
  private readonly ExcelRowMutationService rowMutationService;
  private readonly ExcelImagePlacementService imagePlacementService;

  public ExcelAutomaticPlacementService(
    ExcelSheetSnapshotService? snapshotService = null,
    CaseLayoutAnalyzer? layoutAnalyzer = null,
    ContentOccupancyAnalyzer? occupancyAnalyzer = null,
    PlacementPlanner? placementPlanner = null,
    ExcelRowMutationService? rowMutationService = null,
    ExcelImagePlacementService? imagePlacementService = null)
  {
    this.snapshotService = snapshotService ?? new ExcelSheetSnapshotService();
    this.layoutAnalyzer = layoutAnalyzer ?? new CaseLayoutAnalyzer();
    this.occupancyAnalyzer = occupancyAnalyzer ?? new ContentOccupancyAnalyzer();
    this.placementPlanner = placementPlanner ?? new PlacementPlanner(new ImageSizingService());
    this.rowMutationService = rowMutationService ?? new ExcelRowMutationService();
    this.imagePlacementService = imagePlacementService ?? new ExcelImagePlacementService();
  }

  /// <summary>Builds a COM-free plan for every image without changing Excel.</summary>
  public AutomaticPlacementAnalysisResult Analyze(
    WorkbookIdentity workbook,
    string worksheetName,
    EvidenceSide side,
    IReadOnlyList<AutomaticPlacementImage> images,
    bool preferActiveGap = true,
    double horizontalMarginPoints = 6,
    string? requestedCaseLabel = null,
    bool autoDetectSide = false)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ArgumentNullException.ThrowIfNull(images);
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return AutomaticPlacementAnalysisResult.Failed("自動配置の解析はSTAスレッドで実行する必要があります。");
    }

    var validation = ValidateImages(images, requireFiles: false);
    if (validation is not null)
    {
      return AutomaticPlacementAnalysisResult.Failed(validation);
    }

    var captured = snapshotService.Capture(workbook, worksheetName);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return AutomaticPlacementAnalysisResult.Failed(captured.Message);
    }

    var snapshot = captured.Snapshot;
    var selectedCase = ResolveRequestedCase(snapshot, requestedCaseLabel);
    if (!selectedCase.Succeeded)
    {
      return AutomaticPlacementAnalysisResult.Failed(selectedCase.Message);
    }

    if (selectedCase.Row != snapshot.ActiveCell.Row)
    {
      captured = snapshotService.Capture(workbook, snapshot.WorksheetName, selectedCase.Row);
      if (!captured.Succeeded || captured.Snapshot is null)
      {
        return AutomaticPlacementAnalysisResult.Failed(captured.Message);
      }

      snapshot = captured.Snapshot;
    }

    snapshot = snapshot with
    {
      ActiveCell = new CellReference(selectedCase.Row, snapshot.ActiveCell.Column),
      LayoutSignals = snapshot.LayoutSignals with { ActiveRow = selectedCase.Row },
    };
    if (autoDetectSide)
    {
      var currentLayout = layoutAnalyzer.Analyze(snapshot.LayoutSignals).Layout;
      if (currentLayout is not null)
      {
        side = SideForColumn(snapshot.ActiveCell.Column, currentLayout, side);
      }
    }
    if (snapshot.IsReadOnly || snapshot.IsProtected)
    {
      return AutomaticPlacementAnalysisResult.Failed(
        snapshot.IsReadOnly ? "対象Workbookは読み取り専用です。" : $"シート {snapshot.WorksheetName} は保護されています。");
    }

    return AnalyzeSnapshot(snapshot, side, images, preferActiveGap, horizontalMarginPoints);
  }

  /// <summary>Builds the same plan from an already captured immutable snapshot.</summary>
  public AutomaticPlacementAnalysisResult AnalyzeSnapshot(
    SheetSnapshot snapshot,
    EvidenceSide side,
    IReadOnlyList<AutomaticPlacementImage> images,
    bool preferActiveGap = true,
    double horizontalMarginPoints = 6,
    string? requestedCaseLabel = null)
  {
    ArgumentNullException.ThrowIfNull(snapshot);
    ArgumentNullException.ThrowIfNull(images);
    var validation = ValidateImages(images, requireFiles: false);
    if (validation is not null)
    {
      return AutomaticPlacementAnalysisResult.Failed(validation);
    }

    if (snapshot.IsReadOnly || snapshot.IsProtected)
    {
      return AutomaticPlacementAnalysisResult.Failed(
        snapshot.IsReadOnly ? "対象Workbookは読み取り専用です。" : $"シート {snapshot.WorksheetName} は保護されています。");
    }

    var selectedCase = ResolveRequestedCase(snapshot, requestedCaseLabel);
    if (!selectedCase.Succeeded)
    {
      return AutomaticPlacementAnalysisResult.Failed(selectedCase.Message);
    }

    snapshot = snapshot with
    {
      ActiveCell = new CellReference(selectedCase.Row, snapshot.ActiveCell.Column),
      LayoutSignals = snapshot.LayoutSignals with { ActiveRow = selectedCase.Row },
    };

    var analyzed = layoutAnalyzer.Analyze(snapshot.LayoutSignals);
    if (!analyzed.IsSafe || analyzed.Layout is null)
    {
      return AutomaticPlacementAnalysisResult.Failed(string.Join(" ", analyzed.Reasons));
    }

    try
    {
      var layout = analyzed.Layout;
      var contents = occupancyAnalyzer.Analyze(snapshot, layout).ToList();
      var rowHeights = snapshot.RowHeights.ToDictionary(pair => pair.Key, pair => pair.Value);
      var plans = new List<AutomaticPlacementStep>(images.Count);
      for (var index = 0; index < images.Count; index++)
      {
        var sideColumns = side is EvidenceSide.New ? layout.NewRegion : layout.OldRegion;
        var width = AvailableWidth(snapshot, sideColumns, horizontalMarginPoints);
        var plan = placementPlanner.Plan(new PlacementRequest(
          layout,
          side,
          images[index].Dimensions,
          width,
          index == 0 ? snapshot.ActiveCell.Row : layout.EndRow,
          index == 0 && preferActiveGap,
          contents,
          rowHeights));
        plans.Add(new AutomaticPlacementStep(index, images[index], plan, width));
        ApplyPlanToModel(plan, side, ref layout, contents, rowHeights);
      }

      return new AutomaticPlacementAnalysisResult(
        true,
        snapshot.WorksheetName,
        analyzed,
        plans,
        Fingerprint(snapshot),
        ResolveCaseLabel(snapshot, analyzed.Layout!.StartRow),
        snapshot.ActiveCell.Row,
        side,
        $"{snapshot.WorksheetName} の{side}側へ{plans.Count}件を配置する計画を作成しました。")
      {
        LayoutSignals = snapshot.LayoutSignals,
      };
    }
    catch (Exception exception) when (exception is ArgumentException or InvalidOperationException or OverflowException)
    {
      return AutomaticPlacementAnalysisResult.Failed($"配置計画を作成できません: {exception.Message}");
    }
  }

  /// <summary>Analyzes and applies all images. A failed step compensates prior Shapes and inserted rows in reverse order.</summary>
  public AutomaticPlacementResult PlaceImages(
    WorkbookIdentity workbook,
    string worksheetName,
    EvidenceSide side,
    IReadOnlyList<AutomaticPlacementImage> images,
    bool preferActiveGap = true,
    double horizontalMarginPoints = 6,
    AutomaticPlacementAnalysisResult? preparedAnalysis = null,
    string? requestedCaseLabel = null)
  {
    var validation = ValidateImages(images, requireFiles: true);
    if (validation is not null)
    {
      return AutomaticPlacementResult.Failed(validation);
    }

    var initialAnalysis = preparedAnalysis ?? Analyze(
      workbook,
      worksheetName,
      side,
      images,
      preferActiveGap,
      horizontalMarginPoints,
      requestedCaseLabel);
    if (!initialAnalysis.Succeeded)
    {
      return AutomaticPlacementResult.Failed(initialAnalysis.Message, initialAnalysis);
    }

    if (!AnalysisMatchesRequest(initialAnalysis, worksheetName, images) || !SnapshotStillMatches(workbook, initialAnalysis))
    {
      return AutomaticPlacementResult.Failed(
        "プレビュー後にWorksheetが変更されました。再度キャプチャを確認してください。",
        initialAnalysis);
    }

    var appliedRows = new List<AppliedRowInsertion>();
    var placed = new List<AutomaticPlacedImage>();
    var executedSteps = new List<AutomaticPlacementStep>();
    for (var index = 0; index < images.Count; index++)
    {
      var step = initialAnalysis.Steps[index];
      var currentCaseEnd = initialAnalysis.LayoutAnalysis!.Layout!.EndRow + appliedRows.Sum(row => row.Count);
      foreach (var insertion in step.Plan.Insertions)
      {
        var appliedInsertion = ResolveAppliedInsertion(currentCaseEnd, insertion);
        var mutation = rowMutationService.InsertRows(
          workbook,
          initialAnalysis.WorksheetName,
          appliedInsertion);
        if (!mutation.Succeeded || !mutation.Changed)
        {
          return Compensate(workbook, initialAnalysis, placed, appliedRows, mutation.Message);
        }

        appliedRows.Add(new AppliedRowInsertion(mutation.WorksheetName, mutation.StartRow, mutation.Count, insertion.Reason));
        currentCaseEnd = checked(currentCaseEnd + mutation.Count);
      }

      var placement = imagePlacementService.PlaceImage(
        workbook,
        initialAnalysis.WorksheetName,
        step.Plan.FocusCell,
        side,
        step.Image.ImagePath,
        step.Image.Dimensions,
        step.AvailableWidthPoints,
        horizontalMarginPoints);
      if (!placement.Succeeded)
      {
        return Compensate(workbook, initialAnalysis, placed, appliedRows, placement.Message);
      }

      placed.Add(new AutomaticPlacedImage(
        step.Index,
        placement.ShapeName,
        placement.WorksheetName,
        placement.FocusCell,
        placement.FocusSucceeded,
        placement.Target!,
        step.Plan));
      executedSteps.Add(step);
    }

    var executedAnalysis = initialAnalysis with { Steps = executedSteps };

    return new AutomaticPlacementResult(
      true,
      true,
      executedAnalysis,
      placed,
      appliedRows,
      [],
      $"{placed.Count}件の画像を配置しました（未保存）。");
  }

  private bool SnapshotStillMatches(
    WorkbookIdentity workbook,
    AutomaticPlacementAnalysisResult expected)
  {
    var current = snapshotService.Capture(workbook, expected.WorksheetName, expected.AnalysisRow);
    if (!current.Succeeded || current.Snapshot is null)
    {
      return false;
    }

    var normalized = current.Snapshot with
    {
      ActiveCell = new CellReference(expected.AnalysisRow, current.Snapshot.ActiveCell.Column),
      LayoutSignals = current.Snapshot.LayoutSignals with { ActiveRow = expected.AnalysisRow },
    };
    return string.Equals(Fingerprint(normalized), expected.SnapshotFingerprint, StringComparison.Ordinal);
  }

  private static bool AnalysisMatchesRequest(
    AutomaticPlacementAnalysisResult analysis,
    string worksheetName,
    IReadOnlyList<AutomaticPlacementImage> images) =>
    (string.Equals(worksheetName, "ActiveSheet", StringComparison.OrdinalIgnoreCase) ||
      string.Equals(worksheetName, analysis.WorksheetName, StringComparison.OrdinalIgnoreCase)) &&
    analysis.Steps.Count == images.Count &&
    analysis.Steps.Select(step => step.Image).SequenceEqual(images);

  private static string ResolveCaseLabel(SheetSnapshot snapshot, int startRow)
  {
    var anchor = snapshot.LayoutSignals.Anchors.Last(anchor => anchor.Row <= startRow);
    return FormatCaseLabel(anchor);
  }

  private static RequestedCaseResolution ResolveRequestedCase(
    SheetSnapshot snapshot,
    string? requestedCaseLabel)
  {
    if (string.IsNullOrWhiteSpace(requestedCaseLabel))
    {
      return new RequestedCaseResolution(true, snapshot.ActiveCell.Row, string.Empty);
    }

    var matches = snapshot.LayoutSignals.Anchors
      .Where(anchor => string.Equals(FormatCaseLabel(anchor), requestedCaseLabel.Trim(), StringComparison.CurrentCultureIgnoreCase))
      .ToArray();
    return matches.Length switch
    {
      1 => new RequestedCaseResolution(true, matches[0].Row, string.Empty),
      0 => new RequestedCaseResolution(false, 0, $"Case '{requestedCaseLabel.Trim()}' が見つかりません。"),
      _ => new RequestedCaseResolution(false, 0, $"Case '{requestedCaseLabel.Trim()}' が複数あるため選択できません。"),
    };
  }

  public static string FormatCaseLabel(CaseAnchorSignal anchor)
  {
    var values = new[] { anchor.ColumnAValue, anchor.ColumnBValue }
      .Where(value => !string.IsNullOrWhiteSpace(value));
    var label = string.Join("-", values);
    return string.IsNullOrWhiteSpace(label) ? $"開始行 {anchor.Row}" : label;
  }

  private static EvidenceSide SideForColumn(
    int column,
    EvidenceCaseLayout layout,
    EvidenceSide fallback) =>
    column >= layout.OldRegion.FirstColumn && column <= layout.OldRegion.LastColumn
      ? EvidenceSide.Old
      : column >= layout.NewRegion.FirstColumn && column <= layout.NewRegion.LastColumn
        ? EvidenceSide.New
        : fallback;

  private static RowInsertion ResolveAppliedInsertion(int currentCaseEnd, RowInsertion insertion)
  {
    return currentCaseEnd < ExcelWorksheetLimits.MaximumRow && insertion.AtRow == currentCaseEnd + 1
      ? insertion with
      {
        AtRow = currentCaseEnd,
        Reason = $"{insertion.Reason} Insert above the Case bottom boundary.",
      }
      : insertion;
  }

  private AutomaticPlacementResult Compensate(
    WorkbookIdentity workbook,
    AutomaticPlacementAnalysisResult analysis,
    IReadOnlyList<AutomaticPlacedImage> placed,
    IReadOnlyList<AppliedRowInsertion> insertedRows,
    string failure)
  {
    var errors = new List<string>();
    foreach (var image in placed.Reverse())
    {
      var deletion = imagePlacementService.DeletePlacedImage(workbook, image.WorksheetName, image.ShapeName);
      if (!deletion.Succeeded)
      {
        errors.Add($"Shape {image.ShapeName}: {deletion.Message}");
      }
    }

    foreach (var insertion in insertedRows.Reverse())
    {
      var deletion = rowMutationService.DeleteRowsIfSafe(
        workbook,
        insertion.WorksheetName,
        insertion.StartRow,
        insertion.Count);
      if (!deletion.Succeeded || !deletion.Changed)
      {
        errors.Add($"Rows {insertion.StartRow}-{insertion.StartRow + insertion.Count - 1}: {deletion.Message}");
      }
    }

    var compensated = errors.Count == 0;
    var message = compensated
      ? $"自動配置に失敗したため変更を取り消しました: {failure}"
      : $"自動配置に失敗し、一部を自動復旧できませんでした: {failure} {string.Join(" ", errors)}";
    return new AutomaticPlacementResult(
      false,
      compensated,
      analysis,
      placed,
      insertedRows,
      errors,
      message);
  }

  private static string? ValidateImages(IReadOnlyList<AutomaticPlacementImage>? images, bool requireFiles)
  {
    if (images is null || images.Count == 0)
    {
      return "配置する画像がありません。";
    }

    foreach (var image in images)
    {
      if (image is null || string.IsNullOrWhiteSpace(image.ImagePath))
      {
        return "画像パスが指定されていません。";
      }

      if (!double.IsFinite(image.Dimensions.WidthPoints) || image.Dimensions.WidthPoints <= 0 ||
        !double.IsFinite(image.Dimensions.HeightPoints) || image.Dimensions.HeightPoints <= 0)
      {
        return $"画像サイズが不正です: {image.ImagePath}";
      }

      if (requireFiles && !File.Exists(image.ImagePath))
      {
        return $"画像ファイルが見つかりません: {image.ImagePath}";
      }
    }

    return null;
  }

  private static double AvailableWidth(
    SheetSnapshot snapshot,
    ColumnRange columns,
    double horizontalMarginPoints)
  {
    if (!double.IsFinite(horizontalMarginPoints) || horizontalMarginPoints < 0)
    {
      throw new ArgumentOutOfRangeException(nameof(horizontalMarginPoints));
    }

    if (columns.Count < 2)
    {
      throw new InvalidOperationException("配置先Sideには余白列と画像列が必要です。");
    }

    var width = Enumerable.Range(columns.FirstColumn + 1, columns.Count - 1)
      .Sum(column => snapshot.ColumnWidths.GetValueOrDefault(column));
    var availableWidth = width - (horizontalMarginPoints * 2);
    if (!double.IsFinite(availableWidth) || availableWidth <= 0)
    {
      throw new InvalidOperationException("配置先Sideの列幅を取得できません。");
    }

    return availableWidth;
  }

  private static string Fingerprint(SheetSnapshot snapshot)
  {
    var text = new StringBuilder()
      .Append(snapshot.WorksheetName).Append('|')
      .Append(snapshot.ActiveCell.Row).Append(',').Append(snapshot.ActiveCell.Column).Append('|')
      .Append(snapshot.RawUsedFirstRow).Append(',').Append(snapshot.RawUsedLastRow).Append(',')
      .Append(snapshot.RawUsedFirstColumn).Append(',').Append(snapshot.RawUsedLastColumn).Append('|')
      .Append(snapshot.IsReadOnly).Append(',').Append(snapshot.IsProtected).Append('|');
    var signals = snapshot.LayoutSignals;
    text.Append('L').Append(signals.ActiveRow).Append(',').Append(signals.RawUsedLastRow).Append(',')
      .Append(signals.LogicalEvidenceLastRow).Append(',').Append(signals.NewFirstColumn).Append(',')
      .Append(signals.ObservedLastEvidenceColumn).Append('|');
    foreach (var anchor in signals.Anchors.OrderBy(anchor => anchor.Row))
    {
      text.Append('A').Append(anchor.Row).Append(',').Append(anchor.HasValueInColumnA).Append(',')
        .Append(anchor.HasValueInColumnB).Append(',').Append(anchor.HasTopBorder).Append('|');
    }

    foreach (var boundary in signals.VerticalBoundaries
      .OrderBy(boundary => boundary.RightColumn)
      .ThenBy(boundary => boundary.StartRow))
    {
      text.Append('V').Append(boundary.RightColumn).Append(',').Append(boundary.StartRow).Append(',')
        .Append(boundary.EndRow).Append('|');
    }

    foreach (var boundary in signals.HorizontalBoundaries
      .OrderBy(boundary => boundary.Row)
      .ThenBy(boundary => boundary.FirstColumn))
    {
      text.Append('H').Append(boundary.Row).Append(',').Append(boundary.FirstColumn).Append(',')
        .Append(boundary.LastColumn).Append('|');
    }

    foreach (var column in signals.OldHeaderColumns.Order())
    {
      text.Append('O').Append(column).Append('|');
    }

    foreach (var cell in snapshot.Cells.OrderBy(cell => cell.Row).ThenBy(cell => cell.Column))
    {
      text.Append('C').Append(cell.Row).Append(',').Append(cell.Column).Append(',')
        .Append(cell.HasValueOrFormula).Append(',').Append(cell.HasCommentOrNote).Append(',')
        .Append(cell.HasHyperlink).Append(',').Append(cell.IntersectsMerge).Append('|');
    }

    foreach (var shape in snapshot.Shapes.OrderBy(shape => shape.Name, StringComparer.Ordinal))
    {
      text.Append('S').Append(shape.Name).Append(',').Append(shape.StartRow).Append(',')
        .Append(shape.EndRow).Append(',').Append(shape.StartColumn).Append(',')
        .Append(shape.EndColumn).Append(',').Append(shape.IsManagedImage).Append('|');
    }

    foreach (var height in snapshot.RowHeights.OrderBy(pair => pair.Key))
    {
      text.Append('R').Append(height.Key).Append('=').Append(height.Value.ToString("R", System.Globalization.CultureInfo.InvariantCulture)).Append('|');
    }

    foreach (var width in snapshot.ColumnWidths.OrderBy(pair => pair.Key))
    {
      text.Append('W').Append(width.Key).Append('=').Append(width.Value.ToString("R", System.Globalization.CultureInfo.InvariantCulture)).Append('|');
    }

    return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text.ToString())));
  }

  private static void ApplyPlanToModel(
    PlacementPlan plan,
    EvidenceSide side,
    ref EvidenceCaseLayout layout,
    List<ContentSpan> contents,
    Dictionary<int, double> rowHeights)
  {
    foreach (var insertion in plan.Insertions)
    {
      for (var index = 0; index < contents.Count; index++)
      {
        var content = contents[index];
        contents[index] = content with
        {
          StartRow = content.StartRow >= insertion.AtRow ? checked(content.StartRow + insertion.Count) : content.StartRow,
          EndRow = content.EndRow >= insertion.AtRow ? checked(content.EndRow + insertion.Count) : content.EndRow,
        };
      }

      var shiftedHeights = rowHeights
        .OrderByDescending(pair => pair.Key)
        .Where(pair => pair.Key >= insertion.AtRow)
        .ToArray();
      foreach (var pair in shiftedHeights)
      {
        rowHeights.Remove(pair.Key);
        rowHeights[checked(pair.Key + insertion.Count)] = pair.Value;
      }

      layout = layout with { EndRow = checked(layout.EndRow + insertion.Count) };
    }

    contents.Add(new ContentSpan(side, plan.StartRow, plan.EndRow, ContentKind.ManagedImage));
  }

  private sealed record RequestedCaseResolution(bool Succeeded, int Row, string Message);
}

public sealed record AutomaticPlacementImage(string ImagePath, ImageDimensions Dimensions);

public sealed record AutomaticPlacementStep(
  int Index,
  AutomaticPlacementImage Image,
  PlacementPlan Plan,
  double AvailableWidthPoints);

public sealed record AutomaticPlacementAnalysisResult(
  bool Succeeded,
  string WorksheetName,
  LayoutAnalysisResult? LayoutAnalysis,
  IReadOnlyList<AutomaticPlacementStep> Steps,
  string SnapshotFingerprint,
  string CaseLabel,
  int AnalysisRow,
  EvidenceSide ResolvedSide,
  string Message)
{
  public SheetLayoutSignals? LayoutSignals { get; init; }

  public static AutomaticPlacementAnalysisResult Failed(string message) =>
    new(false, string.Empty, null, [], string.Empty, string.Empty, 0, EvidenceSide.New, message);
}

public sealed record AppliedRowInsertion(string WorksheetName, int StartRow, int Count, string Reason);

public sealed record AutomaticPlacedImage(
  int Index,
  string ShapeName,
  string WorksheetName,
  CellReference FocusCell,
  bool FocusSucceeded,
  ManagedShapeTarget Target,
  PlacementPlan Plan);

public sealed record AutomaticPlacementResult(
  bool Succeeded,
  bool CompensationSucceeded,
  AutomaticPlacementAnalysisResult? Analysis,
  IReadOnlyList<AutomaticPlacedImage> PlacedImages,
  IReadOnlyList<AppliedRowInsertion> AppliedInsertions,
  IReadOnlyList<string> CompensationErrors,
  string Message)
{
  public static AutomaticPlacementResult Failed(
    string message,
    AutomaticPlacementAnalysisResult? analysis = null) =>
    new(false, true, analysis, [], [], [], message);
}
