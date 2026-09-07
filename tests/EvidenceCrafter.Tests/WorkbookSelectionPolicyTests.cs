using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class WorkbookSelectionPolicyTests
{
  private static readonly WorkbookIdentity First = Create("first");
  private static readonly WorkbookIdentity Second = Create("second");

  [TestMethod]
  public void Resolve_InitialDiscovery_DoesNotAutoSelectFirstWorkbook()
  {
    var result = WorkbookSelectionPolicy.Resolve(null, [First, Second]);

    Assert.IsNull(result);
  }

  [TestMethod]
  public void Resolve_MissingPreviousWorkbook_DoesNotSubstituteAnotherWorkbook()
  {
    var result = WorkbookSelectionPolicy.Resolve("closed", [First, Second]);

    Assert.IsNull(result);
  }

  [TestMethod]
  public void Resolve_ExistingPreviousWorkbook_PreservesSelection()
  {
    var result = WorkbookSelectionPolicy.Resolve("second", [First, Second]);

    Assert.AreSame(Second, result);
  }

  private static WorkbookIdentity Create(string id) =>
    new(id, $"{id}.xlsx", $"C:\\temp\\{id}.xlsx", IntPtr.Zero, 1, false);
}
