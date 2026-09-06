using CraftEvidence.Core.Models;
using CraftEvidence.Core.Services;

namespace CraftEvidence.Excel;

/// <summary>Runs conservative maintenance for the Case containing a known managed-image row.</summary>
public sealed class ExcelCaseMaintenanceService
{
  private readonly ExcelSheetSnapshotService snapshotService = new();
  private readonly CaseLayoutAnalyzer layoutAnalyzer = new();
  private readonly ExcelRowMutationService rowMutationService;

  public ExcelCaseMaintenanceService(ExcelRowMutationService? rowMutationService = null)
  {
    this.rowMutationService = rowMutationService ?? new ExcelRowMutationService();
  }

  public RowMutationResult TrimCaseTail(
    WorkbookIdentity workbook,
    string worksheetName,
    int caseRow,
    int tailRows = 4)
  {
    var captured = snapshotService.Capture(workbook, worksheetName, caseRow);
    if (!captured.Succeeded || captured.Snapshot is null)
    {
      return RowMutationResult.Failed(RowMutationOperation.Delete, worksheetName, captured.Message);
    }

    var signals = captured.Snapshot.LayoutSignals with { ActiveRow = caseRow };
    var analyzed = layoutAnalyzer.Analyze(signals);
    if (!analyzed.IsSafe || analyzed.Layout is null)
    {
      return RowMutationResult.Failed(
        RowMutationOperation.Delete,
        captured.Snapshot.WorksheetName,
        $"Case末尾を安全に解析できないため行整理を行いません: {string.Join(" ", analyzed.Reasons)}");
    }

    return rowMutationService.DeleteTrailingRowsWithSnapshot(
      workbook,
      captured.Snapshot.WorksheetName,
      analyzed.Layout.StartRow,
      analyzed.Layout.EndRow,
      tailRows);
  }
}
