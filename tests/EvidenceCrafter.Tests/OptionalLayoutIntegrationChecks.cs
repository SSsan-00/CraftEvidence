using System.Globalization;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

public sealed partial class ExcelSessionCatalogIntegrationTests
{
  // Run only in the existing supervised, disposable Excel process, before its legacy fixture.
  private static void VerifyOptionalLayoutScenarios(
    object worksheet,
    WorkbookIdentity identity,
    string placementImagePath)
  {
    object? area = null;
    object? shapes = null;
    object? selection = null;
    object? previousValues = null;
    var worksheetName = Convert.ToString(GetRequiredProperty(worksheet, "Name"), CultureInfo.InvariantCulture)!;
    var automatic = new ExcelAutomaticPlacementService();
    var snapshotService = new ExcelSheetSnapshotService();
    var analyzer = new CaseLayoutAnalyzer();
    var request = new[] { new AutomaticPlacementImage(placementImagePath, new ImageDimensions(120, 30)) };
    var widthsWithoutBorders = new Dictionary<bool, double>();
    try
    {
      area = GetRequiredProperty(worksheet, "Range", "A1:AF112");
      previousValues = GetRequiredProperty(area, "Value2");
      shapes = GetRequiredProperty(worksheet, "Shapes");
      Assert.AreEqual(0, Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture),
        "The disposable optional-layout fixture must start before image creation.");
      selection = GetRequiredProperty(worksheet, "Cells", 10, 7);

      foreach (var hasBorders in new[] { false, true })
      {
        foreach (var bothSides in new[] { false, true })
        {
          _ = InvokeMethod(area, "Clear");
          // Formatting extends UsedRange to the printable width, even with no borders or values there.
          SetProperty(area, "NumberFormat", "0.00");
          SetCellValue(worksheet, 2, 3, "新");
          if (bothSides) SetCellValue(worksheet, 2, 18, "旧");
          SetCellValue(worksheet, 3, 1, 1);
          SetCellValue(worksheet, 3, 2, 1);
          SetCellValue(worksheet, 33, 2, 2);
          SetCellValue(worksheet, 83, 2, 3);
          if (hasBorders)
          {
            foreach (var row in new[] { 3, 33, 83 }) SetRangeBorder(worksheet, row, 1, row, 32, 8);
            if (bothSides) SetRangeBorder(worksheet, 3, 17, 112, 17, 10);
            SetRangeBorder(worksheet, 112, 1, 112, 32, 9);
          }
          _ = InvokeMethod(worksheet, "Activate");
          _ = InvokeMethod(selection, "Select");
          var context = $"borders={hasBorders}, bothSides={bothSides}";
          var captured = snapshotService.Capture(identity, worksheetName);
          Assert.IsTrue(captured.Succeeded, $"{context}: {captured.Message}");
          Assert.IsNotNull(captured.Snapshot);
          var layout = analyzer.Analyze(captured.Snapshot.LayoutSignals).Layout;
          Assert.IsNotNull(layout, context);
          Assert.AreEqual(3, layout.StartRow, context);
          Assert.AreEqual(32, layout.EndRow, context);
          Assert.AreEqual(bothSides, layout.SupportsSide(EvidenceSide.Old), context);

          var planned = automatic.AnalyzeSnapshot(captured.Snapshot, EvidenceSide.New, request,
            requestedCaseLabel: "1-1");
          Assert.IsTrue(planned.Succeeded, $"{context}: {planned.Message}");
          Assert.AreEqual(new CellReference(5, 4), planned.Steps[0].Plan.FocusCell,
            "Placement must use the Case anchor, not the selected G10 cell.");
          Assert.IsEmpty(planned.Steps[0].Plan.Insertions, context);
          if (hasBorders)
          {
            Assert.AreEqual(widthsWithoutBorders[bothSides], planned.Steps[0].AvailableWidthPoints, 0.01, context);
          }
          else
          {
            widthsWithoutBorders[bothSides] = planned.Steps[0].AvailableWidthPoints;
          }
          var oldPlan = automatic.AnalyzeSnapshot(captured.Snapshot, EvidenceSide.Old, request,
            requestedCaseLabel: "1-1");
          Assert.AreEqual(bothSides, oldPlan.Succeeded, $"{context}: {oldPlan.Message}");

          var second = analyzer.Analyze(captured.Snapshot.LayoutSignals with { ActiveRow = 33 }).Layout;
          Assert.IsNotNull(second, context);
          Assert.AreEqual(82, second.EndRow, "Nonuniform Case spacing must use the next actual anchor.");
          var last = analyzer.Analyze(captured.Snapshot.LayoutSignals with { ActiveRow = 83 }).Layout;
          Assert.IsNotNull(last, context);
          Assert.IsFalse(last.CanDeleteTrailingRows, "The final Case must never allow automatic trailing-row deletion.");

          SetCellValue(worksheet, 33, 2, 1);
          var duplicate = automatic.Analyze(identity, worksheetName, EvidenceSide.New, request,
            requestedCaseLabel: "1-1");
          Assert.IsFalse(duplicate.Succeeded, "Duplicate canonical Case numbers must never choose an arbitrary block.");
        }
      }
      Assert.IsGreaterThan(widthsWithoutBorders[true], widthsWithoutBorders[false],
        "New-only images must use more width than the New region of a two-sided sheet.");
    }
    finally
    {
      // Any test-created shape is disposable; cleanup also runs after assertion failure.
      if (shapes is not null)
      {
        while (Convert.ToInt32(GetRequiredProperty(shapes, "Count"), CultureInfo.InvariantCulture) > 0)
        {
          object? shape = null;
          try
          {
            shape = InvokeMethod(shapes, "Item", 1);
            if (shape is not null) _ = InvokeMethod(shape, "Delete");
          }
          finally { Release(shape); }
        }
      }
      if (area is not null)
      {
        _ = InvokeMethod(area, "Clear");
        if (previousValues is not null) SetProperty(area, "Value2", previousValues);
      }
      Release(selection);
      Release(shapes);
      Release(area);
    }
  }
}
