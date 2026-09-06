using CraftEvidence.Core.Models;
using CraftEvidence.Core.Services;

namespace CraftEvidence.Tests;

[TestClass]
public sealed class TrailingRowDeletionPlannerTests
{
  private readonly TrailingRowDeletionPlanner planner = new();

  [TestMethod]
  public void Plan_DeletesOnlyRowsBeyondTail()
  {
    var rows = Enumerable.Range(15, 6)
      .Select(row => new RowSafetyState(row, false, false, false, false, false))
      .ToArray();

    var result = planner.Plan(caseStartRow: 3, caseEndRow: 20, lastContentRow: 10, tailRows: 4, rows);

    CollectionAssert.AreEqual(new[] { 15, 16, 17, 18, 19 }, result.ToArray());
  }

  [TestMethod]
  public void Plan_StopsAtFirstUnsafeTrailingRow()
  {
    RowSafetyState[] rows =
    [
      new(15, false, false, false, false, false),
      new(16, false, false, false, false, false),
      new(17, true, false, false, false, false),
      new(18, false, false, false, false, false),
      new(19, false, false, false, false, false),
      new(20, false, false, false, false, false),
    ];

    var result = planner.Plan(caseStartRow: 3, caseEndRow: 20, lastContentRow: 10, tailRows: 4, rows);

    CollectionAssert.AreEqual(new[] { 18, 19 }, result.ToArray());
  }

  [TestMethod]
  public void Plan_MissingSafetySnapshot_PreventsFurtherDeletion()
  {
    RowSafetyState[] rows =
    [
      new(18, false, false, false, false, false),
      new(19, false, false, false, false, false),
    ];

    var result = planner.Plan(caseStartRow: 3, caseEndRow: 20, lastContentRow: 10, tailRows: 4, rows);

    CollectionAssert.AreEqual(new[] { 18, 19 }, result.ToArray());
  }

  [TestMethod]
  public void Plan_EmptyLaterCase_NeverDeletesEarlierCaseRows()
  {
    var rows = Enumerable.Range(5, 98)
      .Select(row => new RowSafetyState(row, false, false, false, false, false))
      .ToArray();

    var result = planner.Plan(
      caseStartRow: 53,
      caseEndRow: 102,
      lastContentRow: 0,
      tailRows: 4,
      rows);

    Assert.IsTrue(result.All(row => row >= 57));
    CollectionAssert.DoesNotContain(result.ToArray(), 52);
  }

  [TestMethod]
  public void Plan_TailRowsBeyondExcelLimit_IsRejectedWithoutOverflow()
  {
    var rows = Enumerable.Range(11, 10)
      .Select(row => new RowSafetyState(row, false, false, false, false, false))
      .ToArray();

    Assert.ThrowsExactly<ArgumentOutOfRangeException>(() => planner.Plan(
      caseStartRow: 3,
      caseEndRow: 20,
      lastContentRow: 10,
      tailRows: int.MaxValue,
      rows));
  }
}
