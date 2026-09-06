using CraftEvidence.Core.Models;
using CraftEvidence.Excel;

namespace CraftEvidence.Tests;

[TestClass]
public sealed class VerifiedPlacementFinalizerTests
{
  private static readonly WorkbookIdentity Workbook = new(
    "connection",
    "Book.xlsx",
    "C:\\temp\\Book.xlsx",
    new IntPtr(42),
    24,
    false);

  private static readonly PlacementPlan Plan = new(
    PlacementMode.Tail,
    20,
    25,
    new FittedImage(100, 80, 1),
    [],
    new CellReference(20, 3),
    "test");

  [TestMethod]
  public void FinalizePlacement_VerifiedPlacement_UsesPlanFocusCell()
  {
    var focusService = new RecordingFocusService();
    var finalizer = new VerifiedPlacementFinalizer(focusService);

    var result = finalizer.FinalizePlacement(true, Workbook, "Evidence", Plan);

    Assert.IsTrue(result.PlacementVerified);
    Assert.IsTrue(result.FocusSucceeded);
    Assert.AreEqual(new CellReference(20, 3), focusService.LastCell);
  }

  [TestMethod]
  public void FinalizePlacement_FailedVerification_DoesNotMoveFocus()
  {
    var focusService = new RecordingFocusService();
    var finalizer = new VerifiedPlacementFinalizer(focusService);

    var result = finalizer.FinalizePlacement(false, Workbook, "Evidence", Plan);

    Assert.IsFalse(result.PlacementVerified);
    Assert.IsFalse(result.FocusSucceeded);
    Assert.IsNull(focusService.LastCell);
  }

  [TestMethod]
  public void FinalizePlacement_FocusFailure_PreservesVerifiedPlacementState()
  {
    var focusService = new RecordingFocusService(succeeded: false);
    var finalizer = new VerifiedPlacementFinalizer(focusService);

    var result = finalizer.FinalizePlacement(true, Workbook, "Evidence", Plan);

    Assert.IsTrue(result.PlacementVerified);
    Assert.IsFalse(result.FocusSucceeded);
    Assert.AreEqual(new CellReference(20, 3), focusService.LastCell);
  }

  private sealed class RecordingFocusService(bool succeeded = true) : IPlacementFocusService
  {
    public CellReference? LastCell { get; private set; }

    public FocusResult FocusPlacedImage(
      WorkbookIdentity workbook,
      string worksheetName,
      CellReference focusCell)
    {
      LastCell = focusCell;
      return new FocusResult(succeeded, succeeded ? "focused" : "focus failed");
    }
  }
}
