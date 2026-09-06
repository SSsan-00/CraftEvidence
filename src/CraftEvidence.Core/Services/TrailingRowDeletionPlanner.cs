using CraftEvidence.Core.Models;

namespace CraftEvidence.Core.Services;

/// <summary>Plans deletion only for a fully safe, contiguous trailing range.</summary>
public sealed class TrailingRowDeletionPlanner
{
  public IReadOnlyList<int> Plan(
    int caseStartRow,
    int caseEndRow,
    int lastContentRow,
    int tailRows,
    IReadOnlyList<RowSafetyState> rows)
  {
    ArgumentNullException.ThrowIfNull(rows);
    if (caseStartRow < 1 ||
      caseEndRow < caseStartRow ||
      caseEndRow > ExcelWorksheetLimits.MaximumRow ||
      lastContentRow < 0 ||
      (lastContentRow > 0 && lastContentRow < caseStartRow) ||
      lastContentRow > caseEndRow ||
      tailRows is < 4 or > ExcelWorksheetLimits.MaximumRow)
    {
      throw new ArgumentOutOfRangeException(nameof(caseEndRow));
    }

    var contentBoundary = lastContentRow == 0 ? caseStartRow - 1 : lastContentRow;
    var firstCandidateValue = Math.Max((long)contentBoundary + tailRows + 1, caseStartRow);
    // The analyzed end row is the Case's structural border. Keep it so the
    // next capture can still identify the Case after trailing blank rows shrink.
    var lastDeletableRow = caseEndRow - 1;
    if (firstCandidateValue > lastDeletableRow)
    {
      return [];
    }

    var firstCandidate = (int)firstCandidateValue;
    var byRow = rows.ToDictionary(row => row.Row);
    var planned = new List<int>();
    for (var row = lastDeletableRow; row >= firstCandidate; row--)
    {
      if (!byRow.TryGetValue(row, out var state) || !state.CanDelete)
      {
        break;
      }

      planned.Add(row);
    }

    planned.Sort();
    return planned;
  }
}
