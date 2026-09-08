using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ContentOccupancyAnalyzerTests
{
  [TestMethod]
  public void Analyze_NewOnly_ProtectsCrossingShapesAndIgnoresOutsideShapes()
  {
    var layout = new EvidenceCaseLayout(3, 52, new ColumnRange(3, 17), null);
    var snapshot = new SheetSnapshot("B1", new CellReference(3, 3), 1, 52, 1, 20, false, false,
      new SheetLayoutSignals(3, 52, 52, 3, 17, [], [], [], []),
      [new SnapshotCell(10, 5, true, false, false, false), new SnapshotCell(11, 18, true, false, false, false)],
      [new("inside", 20, 25, 4, 8, true), new("crossing", 30, 35, 16, 20, false), new("outside", 40, 45, 18, 20, false)],
      new Dictionary<int, double>(), new Dictionary<int, double>());
    var spans = new ContentOccupancyAnalyzer().Analyze(snapshot, layout);
    Assert.HasCount(3, spans);
    Assert.IsTrue(spans.Any(span => span.StartRow == 10 && span.Side == EvidenceSide.New));
    Assert.IsTrue(spans.Any(span => span.StartRow == 20 && span.Side == EvidenceSide.New));
    Assert.IsTrue(spans.Any(span => span.StartRow == 30 && span.Side is null));
  }
}
