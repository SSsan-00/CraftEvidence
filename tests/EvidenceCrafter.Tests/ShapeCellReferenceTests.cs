using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ShapeCellReferenceTests
{
  [TestMethod]
  [DataRow("R5C4", 5, 4)]
  [DataRow("R1048576C16384", 1048576, 16384)]
  [DataRow("R0C4", 7, 9)]
  [DataRow("R1048577C4", 7, 9)]
  [DataRow("R1C16385", 7, 9)]
  [DataRow("R[-2]C[3]", 7, 9)]
  [DataRow("R1C1:R2C2", 7, 9)]
  [DataRow("$A$1", 7, 9)]
  [DataRow(null, 7, 9)]
  public void ReadReference_UsesAbsoluteCoordinatesOrOriginalProperties(string? address, int row, int column)
  {
    Assert.AreEqual(new CellReference(row, column),
      ExcelSheetSnapshotService.ReadShapeCellReference(new CellStub(address)));
  }

  [TestMethod]
  public void ReadReference_AddressUnavailable_FallsBackToOriginalProperties()
  {
    Assert.AreEqual(new CellReference(7, 9),
      ExcelSheetSnapshotService.ReadShapeCellReference(new CellStub(null, unavailable: true)));
  }

  public sealed class CellStub(string? address, bool unavailable = false)
  {
    public int Row => 7;
    public int Column => 9;

    [IndexerName("Address")]
    public string? this[bool rowAbsolute, bool columnAbsolute, int style, bool external] =>
      unavailable ? throw new COMException("Address is unavailable.", unchecked((int)0x80020003)) : address;
  }
}
