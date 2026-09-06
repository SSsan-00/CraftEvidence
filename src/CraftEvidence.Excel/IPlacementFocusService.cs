using CraftEvidence.Core.Models;

namespace CraftEvidence.Excel;

public interface IPlacementFocusService
{
  FocusResult FocusPlacedImage(
    WorkbookIdentity workbook,
    string worksheetName,
    CellReference focusCell);
}

public sealed record FocusResult(bool Succeeded, string Message);
