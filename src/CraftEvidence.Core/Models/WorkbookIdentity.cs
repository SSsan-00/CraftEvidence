namespace CraftEvidence.Core.Models;

/// <summary>Stable, COM-free identity for a Workbook discovered in the Running Object Table.</summary>
public sealed record WorkbookIdentity(
  string ConnectionId,
  string Name,
  string FullPath,
  nint ExcelWindowHandle,
  uint ProcessId,
  bool IsReadOnly,
  nint WindowSessionToken = default,
  string RotMonikerDisplayName = "",
  bool HasWorkbookRegistration = false,
  nint ExcelDocumentWindowHandle = default)
{
  public string DisplayLabel =>
    string.IsNullOrWhiteSpace(FullPath) || string.Equals(Name, FullPath, StringComparison.OrdinalIgnoreCase)
      ? $"{Name}  (Excel PID {ProcessId})"
      : $"{Name} — {FullPath}  (Excel PID {ProcessId})";
}
