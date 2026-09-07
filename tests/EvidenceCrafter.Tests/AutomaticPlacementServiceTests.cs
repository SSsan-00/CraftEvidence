using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class AutomaticPlacementServiceTests
{
  [TestMethod]
  public void AnalyzeSnapshot_ExposesCurrentCaseLabelForPreview()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(signals with
    {
      Anchors = [new CaseAnchorSignal(3, true, true, false, "1", "1")],
    });

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("1-1", result.CaseLabel);
  }

  [TestMethod]
  public void AnalyzeSnapshot_RequestedCaseLabel_OverridesActiveCase()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(source with
    {
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(53, true, true, true, "1", "2"),
      ],
    });

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))],
      requestedCaseLabel: "1-1");

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.AreEqual("1-1", result.CaseLabel);
    Assert.AreEqual(3, result.AnalysisRow);
    Assert.AreEqual(3, result.LayoutAnalysis!.Layout!.StartRow);
  }

  [TestMethod]
  public void AnalyzeSnapshot_UnknownRequestedCaseLabel_FailsClearly()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(source with
    {
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(53, true, true, true, "1", "2"),
      ],
    });

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))],
      requestedCaseLabel: "9-9");

    Assert.IsFalse(result.Succeeded);
    StringAssert.Contains(result.Message, "Case '9-9' が見つかりません");
  }

  [TestMethod]
  public void AnalyzeSnapshot_MultipleImages_UsesGapThenStacksWithRequiredBand()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(signals);
    var images = new[]
    {
      new AutomaticPlacementImage("first.png", new ImageDimensions(120, 60)),
      new AutomaticPlacementImage("second.png", new ImageDimensions(120, 60)),
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      images);

    Assert.IsTrue(result.Succeeded, result.Message);
    Assert.HasCount(2, result.Steps);
    Assert.AreEqual(PlacementMode.Gap, result.Steps[0].Plan.Mode);
    Assert.AreEqual(60, result.Steps[0].Plan.StartRow);
    Assert.AreEqual(66, result.Steps[1].Plan.StartRow);
    Assert.AreEqual(new CellReference(66, 4), result.Steps[1].Plan.FocusCell);
    Assert.IsFalse(string.IsNullOrWhiteSpace(result.SnapshotFingerprint));
  }

  [TestMethod]
  public void AnalyzeSnapshot_ProtectedSheet_FailsWithoutPlan()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json")) with
    {
      IsProtected = true,
    };

    var result = new ExcelAutomaticPlacementService().AnalyzeSnapshot(
      snapshot,
      EvidenceSide.New,
      [new AutomaticPlacementImage("image.png", new ImageDimensions(120, 60))]);

    Assert.IsFalse(result.Succeeded);
    Assert.IsEmpty(result.Steps);
    StringAssert.Contains(result.Message, "保護");
  }

  [TestMethod]
  public void AnalyzeSnapshot_CustomMargin_ReducesAvailableWidthOnBothSides()
  {
    var snapshot = Snapshot(FixtureLoader.LoadLayout("default-final-case.json"));
    var image = new[] { new AutomaticPlacementImage("image.png", new ImageDimensions(1_000, 500)) };
    var service = new ExcelAutomaticPlacementService();

    var defaultMargin = service.AnalyzeSnapshot(snapshot, EvidenceSide.New, image, horizontalMarginPoints: 6);
    var wideMargin = service.AnalyzeSnapshot(snapshot, EvidenceSide.New, image, horizontalMarginPoints: 20);

    Assert.IsTrue(defaultMargin.Succeeded, defaultMargin.Message);
    Assert.IsTrue(wideMargin.Succeeded, wideMargin.Message);
    Assert.AreEqual(28, defaultMargin.Steps[0].AvailableWidthPoints - wideMargin.Steps[0].AvailableWidthPoints, 0.001);
  }

  [TestMethod]
  public void ContentOccupancy_IgnoresShapeOutsideEvidenceColumns_ButKeepsCrossSideShape()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json");
    var snapshot = Snapshot(signals) with
    {
      Shapes =
      [
        new SnapshotShape("outside", 60, 61, 1, 2, false),
        new SnapshotShape("cross-side", 62, 63, 17, 18, false),
      ],
    };
    var layout = new CaseLayoutAnalyzer().Analyze(signals).Layout!;

    var spans = new ContentOccupancyAnalyzer().Analyze(snapshot, layout);

    Assert.HasCount(1, spans);
    Assert.AreEqual("cross-side", snapshot.Shapes[1].Name);
    Assert.IsNull(spans[0].Side);
    Assert.AreEqual(62, spans[0].StartRow);
  }

  [TestMethod]
  public void CaseNavigation_ResolvesConfirmedPreviousAndNextAnchors()
  {
    var source = FixtureLoader.LoadLayout("default-final-case.json");
    var signals = source with
    {
      ActiveRow = 60,
      Anchors = source.Anchors.Append(new CaseAnchorSignal(103, false, true, true)).ToArray(),
    };

    var previous = ExcelCaseNavigationService.ResolveTargetRow(signals, CaseNavigationDirection.Previous);
    var next = ExcelCaseNavigationService.ResolveTargetRow(signals, CaseNavigationDirection.Next);

    Assert.AreEqual(3, previous);
    Assert.AreEqual(103, next);
  }

  private static SheetSnapshot Snapshot(SheetLayoutSignals signals) =>
    new(
      "Evidence",
      new CellReference(signals.ActiveRow, 3),
      1,
      signals.RawUsedLastRow,
      1,
      signals.ObservedLastEvidenceColumn ?? 32,
      false,
      false,
      signals,
      [],
      [],
      Enumerable.Range(1, signals.RawUsedLastRow).ToDictionary(row => row, _ => 15.0),
      Enumerable.Range(3, 30).ToDictionary(column => column, _ => 20.0));
}
