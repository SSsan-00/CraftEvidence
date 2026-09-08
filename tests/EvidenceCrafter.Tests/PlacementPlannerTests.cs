using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class PlacementPlannerTests
{
  private readonly PlacementPlanner planner = new(new ImageSizingService());
  private static readonly EvidenceCaseLayout Layout = new(
    3,
    52,
    new ColumnRange(3, 17),
    new ColumnRange(18, 32));

  [TestMethod]
  public void Plan_EmptyCase_StartsAtCaseAnchor()
  {
    var request = CreateRequest(contents: []);

    var result = planner.Plan(request);

    Assert.AreEqual(PlacementMode.CaseStart, result.Mode);
    Assert.AreEqual(5, result.StartRow);
    Assert.AreEqual(6, result.EndRow);
    Assert.AreEqual(new CellReference(5, 4), result.FocusCell);
    Assert.HasCount(0, result.Insertions);
  }

  [TestMethod]
  public void Plan_CaseNumberAtB2_StartsImageAtD4()
  {
    var request = CreateRequest(contents: [], activeRow: 10) with
    {
      Layout = Layout with { StartRow = 2 },
    };

    var result = planner.Plan(request);

    Assert.AreEqual(PlacementMode.CaseStart, result.Mode);
    Assert.AreEqual(new CellReference(4, 4), result.FocusCell);
  }

  [TestMethod]
  public void Plan_OldSide_FocusesPlacedImageTopLeftCell()
  {
    var request = CreateRequest(contents: [], side: EvidenceSide.Old);

    var result = planner.Plan(request);

    Assert.AreEqual(new CellReference(5, 19), result.FocusCell);
  }

  [TestMethod]
  public void Plan_CaptionInsideReservedBand_DoesNotAddAnotherGap()
  {
    ContentSpan[] contents =
    [
      new(EvidenceSide.New, 3, 12, ContentKind.ManagedImage),
      new(EvidenceSide.New, 13, 14, ContentKind.Cell),
    ];

    var result = planner.Plan(CreateRequest(contents));

    Assert.AreEqual(PlacementMode.Tail, result.Mode);
    Assert.AreEqual(15, result.StartRow);
  }

  [TestMethod]
  public void Plan_ContentBelowReservedBand_AddsNewTwoRowBand()
  {
    ContentSpan[] contents =
    [
      new(EvidenceSide.New, 3, 12, ContentKind.ManagedImage),
      new(EvidenceSide.New, 16, 16, ContentKind.Cell),
    ];

    var result = planner.Plan(CreateRequest(contents));

    Assert.AreEqual(19, result.StartRow);
  }

  [TestMethod]
  public void Plan_InsufficientActiveGap_InsertsBeforeLaterContent()
  {
    ContentSpan[] contents =
    [
      new(EvidenceSide.New, 3, 10, ContentKind.ManagedImage),
      new(EvidenceSide.New, 22, 30, ContentKind.ManagedImage),
    ];
    var request = CreateRequest(
      contents,
      image: new ImageDimensions(100, 45),
      activeRow: 20,
      preferGap: true);

    var result = planner.Plan(request);

    Assert.AreEqual(PlacementMode.Gap, result.Mode);
    Assert.AreEqual(20, result.StartRow);
    Assert.AreEqual(22, result.EndRow);
    Assert.IsTrue(result.Insertions.Any(insertion => insertion.AtRow == 22 && insertion.Count == 3));
  }

  [TestMethod]
  public void Plan_ContentCrossingReservedBand_DoesNotOverlapIt()
  {
    ContentSpan[] contents =
    [
      new(EvidenceSide.New, 3, 12, ContentKind.ManagedImage),
      new(EvidenceSide.New, 13, 16, ContentKind.Merge),
    ];

    var result = planner.Plan(CreateRequest(contents));

    Assert.AreEqual(19, result.StartRow);
  }

  [TestMethod]
  public void Plan_OtherSideNearCaseEnd_ExtendsTail()
  {
    ContentSpan[] contents = [new(EvidenceSide.Old, 45, 51, ContentKind.ManagedImage)];

    var result = planner.Plan(CreateRequest(contents));

    Assert.IsTrue(result.Insertions.Any(insertion => insertion.Count == 3));
  }

  [TestMethod]
  public void Plan_ActiveRowInsidePreviousImageGap_FallsBackToTail()
  {
    ContentSpan[] contents = [new(EvidenceSide.New, 3, 18, ContentKind.ManagedImage)];

    var result = planner.Plan(CreateRequest(contents, activeRow: 19, preferGap: true));

    Assert.AreEqual(PlacementMode.Tail, result.Mode);
    Assert.AreEqual(21, result.StartRow);
  }

  [TestMethod]
  public void Plan_RejectsSpacingBelowSpecificationMinimum()
  {
    var request = CreateRequest(contents: []) with { ImageGapRows = 1 };

    Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => planner.Plan(request));
  }

  [TestMethod]
  public void Plan_ImageNearCaseEnd_ExtendsForFourRowTail()
  {
    ContentSpan[] contents = [new(EvidenceSide.New, 3, 48, ContentKind.Cell)];

    var result = planner.Plan(CreateRequest(contents));

    Assert.AreEqual(51, result.StartRow);
    Assert.AreEqual(52, result.EndRow);
    Assert.IsTrue(result.Insertions.Any(insertion => insertion.Count == 4));
  }

  [TestMethod]
  public void Plan_GapInsertionBeyondWorksheetLimit_IsRejected()
  {
    var layout = Layout with { EndRow = 1_048_576 };
    ContentSpan[] contents =
    [
      new(EvidenceSide.New, 3, 10, ContentKind.ManagedImage),
      new(EvidenceSide.New, 22, 30, ContentKind.ManagedImage),
    ];
    var request = CreateRequest(
      contents,
      image: new ImageDimensions(100, 45),
      activeRow: 20,
      preferGap: true) with
    {
      Layout = layout,
    };

    Assert.ThrowsExactly<InvalidOperationException>(() => planner.Plan(request));
  }

  [TestMethod]
  public void Plan_ColumnsBeyondExcelLimit_AreRejected()
  {
    var request = CreateRequest(contents: []) with
    {
      Layout = new EvidenceCaseLayout(
        3,
        52,
        new ColumnRange(3, 17_000),
        new ColumnRange(17_001, 17_002)),
    };

    Assert.ThrowsExactly<ArgumentException>(() => planner.Plan(request));
  }

  [TestMethod]
  public void Plan_OverlappingSideColumns_AreRejected()
  {
    var request = CreateRequest(contents: []) with
    {
      Layout = Layout with
      {
        OldRegion = new ColumnRange(17, 32),
      },
    };

    Assert.ThrowsExactly<ArgumentException>(() => planner.Plan(request));
  }

  [TestMethod]
  public void Plan_NewOnly_PreservesInsetAndRejectsOld()
  {
    var request = CreateRequest([]) with { Layout = Layout with { OldRegion = null } };
    Assert.AreEqual(new CellReference(5, 4), planner.Plan(request).FocusCell);
    Assert.ThrowsExactly<InvalidOperationException>(() => planner.Plan(request with { Side = EvidenceSide.Old }));
  }

  private static PlacementRequest CreateRequest(
    IReadOnlyList<ContentSpan> contents,
    ImageDimensions? image = null,
    int activeRow = 3,
    bool preferGap = false,
    EvidenceSide side = EvidenceSide.New) =>
    new(
      Layout,
      side,
      image ?? new ImageDimensions(100, 30),
      200,
      activeRow,
      preferGap,
      contents,
      new Dictionary<int, double>(),
      ImageGapRows: 2,
      TailRows: 4,
      DefaultRowHeightPoints: 15);
}
