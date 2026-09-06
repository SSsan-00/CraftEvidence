using CraftEvidence.Core.Models;

namespace CraftEvidence.Excel;

/// <summary>
/// The single postcondition boundary for a placement operation. Mutation orchestration must call
/// this only after Apply and Verify have completed; failed verification never changes UI focus.
/// </summary>
public sealed class VerifiedPlacementFinalizer(IPlacementFocusService focusService)
{
  public PlacementFinalizationResult FinalizePlacement(
    bool placementVerified,
    WorkbookIdentity workbook,
    string worksheetName,
    PlacementPlan placementPlan)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);
    ArgumentNullException.ThrowIfNull(placementPlan);

    if (!placementVerified)
    {
      return new PlacementFinalizationResult(
        PlacementVerified: false,
        FocusSucceeded: false,
        "Placement verification failed; Excel focus was not changed.");
    }

    var focus = focusService.FocusPlacedImage(workbook, worksheetName, placementPlan.FocusCell);
    return new PlacementFinalizationResult(
      PlacementVerified: true,
      FocusSucceeded: focus.Succeeded,
      focus.Message);
  }
}

public sealed record PlacementFinalizationResult(
  bool PlacementVerified,
  bool FocusSucceeded,
  string Message);
