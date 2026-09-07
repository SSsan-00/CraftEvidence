using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Excel;

public interface IExcelSessionCatalog
{
  ExcelDiscoveryResult Discover();
}

public sealed record ExcelDiscoveryResult(
  IReadOnlyList<WorkbookIdentity> Workbooks,
  IReadOnlyList<string> Warnings);
