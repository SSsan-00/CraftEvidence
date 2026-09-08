using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class CaseLayoutAnalyzerTests
{
  [TestMethod]
  public void Analyze_IndividualEvidenceWorkbook_ResolvesThirtyRowCaseDespiteFormattedTail()
  {
    var result = analyzer.Analyze(FixtureLoader.LoadLayout("individual-evidence-layout.json"));

    Assert.IsTrue(result.IsSafe, string.Join(" ", result.Reasons));
    Assert.AreEqual(33, result.Layout!.StartRow);
    Assert.AreEqual(62, result.Layout.EndRow);
    Assert.AreEqual(new ColumnRange(3, 17), result.Layout.NewRegion);
    Assert.AreEqual(new ColumnRange(18, 32), result.Layout.OldRegion);
  }

  private readonly CaseLayoutAnalyzer analyzer = new();

  [TestMethod]
  public void Analyze_DefaultFinalCase_UsesMatchingBorderSignals()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json");

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(LayoutConfidence.High, result.Confidence);
    Assert.IsNotNull(result.Layout);
    Assert.AreEqual(53, result.Layout.StartRow);
    Assert.AreEqual(102, result.Layout.EndRow);
    Assert.AreEqual(new ColumnRange(3, 17), result.Layout.NewRegion);
    Assert.AreEqual(new ColumnRange(18, 32), result.Layout.OldRegion);
  }

  [TestMethod]
  public void Analyze_DynamicBothLayout_DoesNotUseDefaultQRBoundary()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json");

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.IsNotNull(result.Layout);
    Assert.AreEqual(5, result.Layout.StartRow);
    Assert.AreEqual(44, result.Layout.EndRow);
    Assert.AreEqual(new ColumnRange(3, 12), result.Layout.NewRegion);
    Assert.AreEqual(new ColumnRange(13, 22), result.Layout.OldRegion);
  }

  [TestMethod]
  public void Analyze_CommonNarrowLayout_UsesObservedGHBoundary()
  {
    var signals = FixtureLoader.LoadLayout("common-narrow-layout.json");

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.IsNotNull(result.Layout);
    Assert.AreEqual(new ColumnRange(3, 7), result.Layout.NewRegion);
    Assert.AreEqual(new ColumnRange(8, 12), result.Layout.OldRegion);
  }

  [TestMethod]
  public void Analyze_CommonNarrowLayout_AllowsFormattingBelowLogicalEvidenceEnd()
  {
    var signals = FixtureLoader.LoadLayout("common-narrow-layout.json");

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(24, signals.RawUsedLastRow);
    Assert.AreEqual(13, result.Layout?.EndRow);
  }

  [TestMethod]
  public void Analyze_FinalCaseWithConflictingEndSignals_WarnsAndUsesLogicalEnd()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with
    {
      VerticalBoundaries = [new VerticalBoundarySignal(17, 3, 101)],
    };

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(102, result.Layout!.EndRow);
    Assert.IsFalse(result.Layout.CanDeleteTrailingRows);
    StringAssert.Contains(string.Join(' ', result.Reasons), "disagree");
  }

  [TestMethod]
  public void Analyze_ConflictingHeaderAndVerticalBoundary_UsesHeaderAndWarns()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      OldHeaderColumns = [18],
    };

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(new ColumnRange(3, 17), result.Layout!.NewRegion);
    StringAssert.Contains(string.Join(' ', result.Reasons), "disagree");
  }

  [TestMethod]
  public void Analyze_NoAnchorAtOrAboveActiveRow_IsUnsafe()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with
    {
      ActiveRow = 2,
    };

    var result = analyzer.Analyze(signals);

    Assert.IsFalse(result.IsSafe);
    StringAssert.Contains(string.Join(' ', result.Reasons), "No Case anchor");
  }

  [TestMethod]
  public void Analyze_HeaderOnlyBoundary_IsSafe()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      VerticalBoundaries = [],
    };

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(new ColumnRange(3, 12), result.Layout!.NewRegion);
  }

  [TestMethod]
  public void Analyze_UnborderedNonFirstAnchor_IsSafe()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with
    {
      Anchors =
      [
        new CaseAnchorSignal(3, true, true, false, "1", "1"),
        new CaseAnchorSignal(20, false, true, false, null, "2"),
        new CaseAnchorSignal(53, false, true, true, null, "3"),
      ],
      ActiveRow = 25,
    };

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(20, result.Layout!.StartRow);
    Assert.AreEqual(52, result.Layout.EndRow);
  }

  [TestMethod]
  public void Analyze_FinalCase_IgnoresItsOwnTopBorder()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with
    {
      HorizontalBoundaries =
      [
        new HorizontalBoundarySignal(53, 1, 32),
        new HorizontalBoundarySignal(102, 1, 32),
      ],
    };

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(102, result.Layout?.EndRow);
  }

  [TestMethod]
  public void Analyze_FinalCaseEndDisagreesWithLogicalEvidenceEnd_UsesLogicalEnd()
  {
    var signals = FixtureLoader.LoadLayout("default-final-case.json") with
    {
      LogicalEvidenceLastRow = 101,
    };

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(101, result.Layout!.EndRow);
    Assert.IsFalse(result.Layout.CanDeleteTrailingRows);
  }

  [TestMethod]
  public void Analyze_ObservedAndHorizontalRightEdgesDisagree_UsesObservedWidthAndWarns()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      ObservedLastEvidenceColumn = 24,
      HorizontalBoundaries = [new HorizontalBoundarySignal(44, 1, 22)],
    };

    var result = analyzer.Analyze(signals);

    Assert.IsTrue(result.IsSafe);
    Assert.AreEqual(24, result.Layout!.LastEvidenceColumn);
    StringAssert.Contains(string.Join(' ', result.Reasons), "disagree");
  }

  [TestMethod]
  public void Analyze_NonFinalCaseBeyondLogicalEvidenceEnd_IsUnsafe()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      LogicalEvidenceLastRow = 40,
    };

    var result = analyzer.Analyze(signals);

    Assert.IsFalse(result.IsSafe);
    StringAssert.Contains(string.Join(' ', result.Reasons), "out-of-range worksheet coordinates");
  }

  [TestMethod]
  public void Analyze_ColumnsBeyondExcelLimit_AreUnsafe()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      ObservedLastEvidenceColumn = 17_002,
      VerticalBoundaries = [new VerticalBoundarySignal(17_000, 5, 84)],
      HorizontalBoundaries = [new HorizontalBoundarySignal(84, 1, 17_002)],
      OldHeaderColumns = [17_001],
    };

    var result = analyzer.Analyze(signals);

    Assert.IsFalse(result.IsSafe);
    StringAssert.Contains(string.Join(' ', result.Reasons), "worksheet limits");
  }

  [TestMethod]
  public void Analyze_MultipleEvidenceRightEdges_IsUnsafeInsteadOfThrowing()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      ObservedLastEvidenceColumn = null,
      HorizontalBoundaries =
      [
        new HorizontalBoundarySignal(44, 1, 22),
        new HorizontalBoundarySignal(44, 1, 24),
      ],
    };

    var result = analyzer.Analyze(signals);

    Assert.IsFalse(result.IsSafe);
    StringAssert.Contains(string.Join(' ', result.Reasons), "right edge cannot be determined");
  }

  [TestMethod]
  public void Analyze_BordersDoNotChangeLayoutOrPlacement()
  {
    var bordered = FixtureLoader.LoadLayout("dynamic-both-layout.json");
    var borderless = bordered with { VerticalBoundaries = [], HorizontalBoundaries = [] };
    Assert.AreEqual(analyzer.Analyze(bordered).Layout, analyzer.Analyze(borderless).Layout);
    Assert.IsFalse(analyzer.Analyze(borderless).Reasons.Any(reason => reason.StartsWith("Warning:")));
  }

  [TestMethod]
  public void Analyze_NewOnly_UsesFullWidthAndRejectsOld()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      OldHeaderColumns = [], VerticalBoundaries = [], HorizontalBoundaries = [], ActiveRow = 45,
    };
    var result = analyzer.Analyze(signals);
    Assert.IsTrue(result.IsSafe, string.Join(' ', result.Reasons));
    Assert.AreEqual(SideLayoutKind.NewOnly, result.Layout!.Kind);
    Assert.AreEqual(new ColumnRange(3, 22), result.Layout.NewRegion);
    Assert.IsNull(result.Layout.OldRegion);
    Assert.IsFalse(result.Layout.SupportsSide(EvidenceSide.Old));
    Assert.ThrowsExactly<InvalidOperationException>(() => result.Layout.RegionFor(EvidenceSide.Old));
    Assert.AreEqual(84, result.Layout.EndRow);
    Assert.IsFalse(result.Layout.CanDeleteTrailingRows);
  }

  [TestMethod]
  public void Analyze_NoNewHeader_IsUnsafeEvenWithBorders()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with { NewHeaderColumns = [] };
    Assert.IsFalse(analyzer.Analyze(signals).IsSafe);
  }

  [TestMethod]
  public void Analyze_SingleColumnSide_IsUnsafe()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with { OldHeaderColumns = [4] };
    Assert.IsFalse(analyzer.Analyze(signals).IsSafe);
  }

  [TestMethod]
  public void Analyze_DuplicateCaseNumber_ReportsAllRows()
  {
    var signals = FixtureLoader.LoadLayout("dynamic-both-layout.json") with
    {
      Anchors = [new(5, true, true, false, "1", "1"), new(45, false, true, false, null, "1")],
    };
    var result = analyzer.Analyze(signals);
    Assert.IsFalse(result.IsSafe);
    StringAssert.Contains(string.Join(' ', result.Reasons), "Case '1-1' is duplicated at rows 5, 45");
  }
}
