using System.Security.Cryptography;
using System.Text;
using CraftEvidence.Core.Models;
using CraftEvidence.Core.Services;

namespace CraftEvidence.Excel;

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
    double horizontalMarginPoints = 6)
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
    double horizontalMarginPoints = 6)
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
        $"{snapshot.WorksheetName} の{side}側へ{plans.Count}件を配置する計画を作成しました。");
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
    double horizontalMarginPoints = 6)
  {
    var validation = ValidateImages(images, requireFiles: true);
    if (validation is not null)
    {
      return AutomaticPlacementResult.Failed(validation);
    }

    var initialAnalysis = Analyze(workbook, worksheetName, side, images, preferActiveGap, horizontalMarginPoints);
    if (!initialAnalysis.Succeeded)
    {
      return AutomaticPlacementResult.Failed(initialAnalysis.Message, initialAnalysis);
    }

    var appliedRows = new List<AppliedRowInsertion>();
    var placed = new List<AutomaticPlacedImage>();
    var executedSteps = new List<AutomaticPlacementStep>();
    for (var index = 0; index < images.Count; index++)
    {
      var liveAnalysis = Analyze(
        workbook,
        initialAnalysis.WorksheetName,
        side,
        [images[index]],
        index == 0 && preferActiveGap,
        horizontalMarginPoints);
      if (!liveAnalysis.Succeeded || liveAnalysis.Steps.Count != 1)
      {
        return Compensate(workbook, initialAnalysis, placed, appliedRows, liveAnalysis.Message);
      }

      var step = liveAnalysis.Steps[0] with { Index = index };
      if (!SnapshotStillMatches(workbook, liveAnalysis))
      {
        return Compensate(
          workbook,
          initialAnalysis,
          placed,
          appliedRows,
          "Worksheet changed after analysis; refresh and retry.");
      }

      var currentCaseEnd = liveAnalysis.LayoutAnalysis!.Layout!.EndRow;
      foreach (var insertion in step.Plan.Insertions)
      {
        var appliedInsertion = ResolveAppliedInsertion(currentCaseEnd, insertion);
        var mutation = rowMutationService.InsertRows(
          workbook,
          liveAnalysis.WorksheetName,
          appliedInsertion);
        if (!mutation.Succeeded || !mutation.Changed)
        {
          return Compensate(workbook, initialAnalysis, placed, appliedRows, mutation.Message);
        }

        appliedRows.Add(new AppliedRowInsertion(mutation.WorksheetName, mutation.StartRow, mutation.Count, insertion.Reason));
        currentCaseEnd = checked(currentCaseEnd + mutation.Count);
      }

      if (step.Plan.Insertions.Count > 0)
      {
        liveAnalysis = Analyze(
          workbook,
          liveAnalysis.WorksheetName,
          side,
          [images[index]],
          preferActiveGap: index == 0 && preferActiveGap,
          horizontalMarginPoints: horizontalMarginPoints);
        if (!liveAnalysis.Succeeded || liveAnalysis.Steps.Count != 1)
        {
          return Compensate(workbook, initialAnalysis, placed, appliedRows, liveAnalysis.Message);
        }

        step = liveAnalysis.Steps[0] with { Index = index };
        if (step.Plan.Insertions.Count > 0)
        {
          return Compensate(
            workbook,
            initialAnalysis,
            placed,
            appliedRows,
            "行挿入後も安全な配置領域を確保できませんでした。");
        }
      }

      if (!SnapshotStillMatches(workbook, liveAnalysis))
      {
        return Compensate(
          workbook,
          initialAnalysis,
          placed,
          appliedRows,
          "Worksheet changed immediately before placement; changes were cancelled.");
      }

      var placement = imagePlacementService.PlaceImage(
        workbook,
        liveAnalysis.WorksheetName,
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
    var current = snapshotService.Capture(workbook, expected.WorksheetName);
    return current.Succeeded && current.Snapshot is not null &&
      string.Equals(Fingerprint(current.Snapshot), expected.SnapshotFingerprint, StringComparison.Ordinal);
  }

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

    var width = Enumerable.Range(columns.FirstColumn, columns.Count)
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
  string Message)
{
  public static AutomaticPlacementAnalysisResult Failed(string message) =>
    new(false, string.Empty, null, [], string.Empty, message);
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
