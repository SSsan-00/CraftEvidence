using CraftEvidence.Core.Models;

namespace CraftEvidence.Core.Services;

/// <summary>Preserves only an explicit Workbook selection; it never substitutes another Workbook.</summary>
public static class WorkbookSelectionPolicy
{
  public static WorkbookIdentity? Resolve(
    string? previousConnectionId,
    IReadOnlyList<WorkbookIdentity> availableWorkbooks)
  {
    ArgumentNullException.ThrowIfNull(availableWorkbooks);
    return previousConnectionId is null
      ? null
      : availableWorkbooks.FirstOrDefault(item => item.ConnectionId == previousConnectionId);
  }
}
