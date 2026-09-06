using CraftEvidence.Core.Models;

namespace CraftEvidence.Excel;

public interface IExcelSessionCatalog
{
  ExcelDiscoveryResult Discover();
}

public sealed record ExcelDiscoveryResult(
  IReadOnlyList<WorkbookIdentity> Workbooks,
  IReadOnlyList<string> Warnings);
