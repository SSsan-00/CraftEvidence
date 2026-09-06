using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using CraftEvidence.Core.Models;

namespace CraftEvidence.Excel;

/// <summary>Captures the current worksheet state into immutable, COM-free records.</summary>
public sealed class ExcelSheetSnapshotService
{
  private const int MaximumSnapshotCells = 250_000;
  private const int NewFirstColumn = 3;
  private const int XlEdgeTop = 8;
  private const int XlEdgeBottom = 9;
  private const int XlEdgeRight = 10;
  private const int XlLineStyleNone = -4142;

  public SheetSnapshotResult Capture(WorkbookIdentity workbook, string worksheetName)
  {
    ArgumentNullException.ThrowIfNull(workbook);
    ArgumentException.ThrowIfNullOrWhiteSpace(worksheetName);

    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return SheetSnapshotResult.Failed(worksheetName, "Excel snapshot capture must run on an STA thread.");
    }

    if (!workbook.HasWorkbookRegistration ||
      string.IsNullOrWhiteSpace(workbook.RotMonikerDisplayName) ||
      !WorkbookSessionTokenRegistry.IsProcessMonitored(workbook.ProcessId) ||
      !WorkbookSessionTokenRegistry.IsValid(workbook.WindowSessionToken))
    {
      return SheetSnapshotResult.Failed(
        worksheetName,
        "Workbook close monitoring is unavailable; refresh before analyzing the worksheet.");
    }

    IRunningObjectTable? runningObjectTable = null;
    IEnumMoniker? monikerEnumerator = null;
    IBindCtx? bindContext = null;
    try
    {
      Marshal.ThrowExceptionForHR(NativeMethods.GetRunningObjectTable(0, out runningObjectTable));
      runningObjectTable.EnumRunning(out monikerEnumerator);
      Marshal.ThrowExceptionForHR(NativeMethods.CreateBindCtx(0, out bindContext));
      var monikers = new IMoniker[1];
      while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0)
      {
        object? runningObject = null;
        try
        {
          monikers[0].GetDisplayName(bindContext, null, out var displayName);
          if (!string.Equals(displayName, workbook.RotMonikerDisplayName, StringComparison.Ordinal))
          {
            continue;
          }

          runningObjectTable.GetObject(monikers[0], out runningObject);
          return runningObject is null
            ? SheetSnapshotResult.Failed(worksheetName, "The selected Workbook is no longer available.")
            : TryCaptureRunningObject(runningObject, workbook, worksheetName) ??
              SheetSnapshotResult.Failed(worksheetName, "The selected Workbook could not be matched.");
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      return SheetSnapshotResult.Failed(
        worksheetName,
        "The selected Workbook is no longer available in the Running Object Table.");
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      return SheetSnapshotResult.Failed(
        worksheetName,
        $"Excel snapshot capture failed (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  private static SheetSnapshotResult? TryCaptureRunningObject(
    object runningObject,
    WorkbookIdentity identity,
    string worksheetName)
  {
    if (TryGetProperty(runningObject, "Workbooks", out var workbooks))
    {
      try
      {
        if (!ApplicationMatches(runningObject, identity))
        {
          return null;
        }

        var count = ReadInt(workbooks!, "Count");
        for (var index = 1; index <= count; index++)
        {
          object? candidate = null;
          try
          {
            candidate = InvokeProperty(workbooks!, "Item", index);
            if (candidate is not null && WorkbookMatches(candidate, identity))
            {
              return CaptureWorkbook(runningObject, candidate, identity, worksheetName);
            }
          }
          finally
          {
            ComRelease.Release(candidate);
          }
        }

        return null;
      }
      finally
      {
        ComRelease.Release(workbooks);
      }
    }

    if (!TryGetProperty(runningObject, "Application", out var application) || application is null)
    {
      return null;
    }

    try
    {
      return ApplicationMatches(application, identity) && WorkbookMatches(runningObject, identity)
        ? CaptureWorkbook(application, runningObject, identity, worksheetName)
        : null;
    }
    finally
    {
      ComRelease.Release(application);
    }
  }

  private static SheetSnapshotResult CaptureWorkbook(
    object application,
    object workbook,
    WorkbookIdentity identity,
    string worksheetName)
  {
    if (!WorkbookWindowMatchesIdentity(workbook, identity))
    {
      return SheetSnapshotResult.Failed(
        worksheetName,
        "The selected Workbook was closed or reopened; refresh and select it again.");
    }

    object? worksheet = null;
    object? activeCell = null;
    object? workbookActiveSheet = null;
    object? windows = null;
    object? window = null;
    object? usedRange = null;
    object? usedRows = null;
    object? usedColumns = null;
    var eventsWereEnabled = false;
    var applicationStateRead = false;
    try
    {
      eventsWereEnabled = Convert.ToBoolean(
        GetRequiredProperty(application, "EnableEvents"),
        CultureInfo.InvariantCulture);
      applicationStateRead = true;
      SetProperty(application, "EnableEvents", false);
      worksheet = ResolveWorksheet(application, workbook, worksheetName, out var resolvedName);
      workbookActiveSheet = GetRequiredProperty(workbook, "ActiveSheet");
      var activeSheetName = Convert.ToString(
        GetRequiredProperty(workbookActiveSheet, "Name"),
        CultureInfo.CurrentCulture);
      if (!string.Equals(activeSheetName, resolvedName, StringComparison.Ordinal))
      {
        return SheetSnapshotResult.Failed(
          resolvedName,
          "対象SheetをExcelでアクティブにしてから再解析してください。");
      }

      windows = GetRequiredProperty(workbook, "Windows");
      window = GetRequiredProperty(windows, "Item", 1);
      activeCell = GetRequiredProperty(window, "RangeSelection");
      usedRange = GetRequiredProperty(worksheet, "UsedRange");
      usedRows = GetRequiredProperty(usedRange, "Rows");
      usedColumns = GetRequiredProperty(usedRange, "Columns");

      var firstRow = ReadInt(usedRange, "Row");
      var firstColumn = ReadInt(usedRange, "Column");
      var rowCount = ReadInt(usedRows, "Count");
      var columnCount = ReadInt(usedColumns, "Count");
      var lastRow = checked(firstRow + rowCount - 1);
      var lastColumn = checked(firstColumn + columnCount - 1);
      if (firstRow < 1 || firstColumn < 1 ||
        lastRow > ExcelWorksheetLimits.MaximumRow ||
        lastColumn > ExcelWorksheetLimits.MaximumColumn ||
        (long)rowCount * columnCount > MaximumSnapshotCells)
      {
        return SheetSnapshotResult.Failed(
          resolvedName,
          $"Worksheet UsedRange is too large or invalid for a safe snapshot ({rowCount:N0} x {columnCount:N0}).");
      }

      var activeReference = new CellReference(ReadInt(activeCell, "Row"), ReadInt(activeCell, "Column"));
      var values = InvokeProperty(usedRange, "Value2");
      var formulas = InvokeProperty(usedRange, "Formula");
      var comments = ReadLinkedCells(worksheet, "Comments");
      comments.UnionWith(ReadLinkedCells(worksheet, "CommentsThreaded", optional: true));
      var hyperlinks = ReadLinkedCells(worksheet, "Hyperlinks");
      var anchors = ReadAnchors(
        worksheet,
        values,
        formulas,
        firstRow,
        firstColumn,
        rowCount,
        columnCount,
        lastColumn);
      var firstAnchorRow = anchors.Count == 0 ? Math.Max(firstRow, 1) : anchors[0].Row;
      var oldHeaderColumns = ReadOldHeaderColumns(
        values,
        formulas,
        firstAnchorRow - 1,
        firstRow,
        firstColumn,
        rowCount,
        columnCount,
        lastColumn);
      var verticalBoundaries = ReadVerticalBoundaries(
        worksheet,
        firstAnchorRow,
        lastRow,
        lastColumn,
        oldHeaderColumns,
        anchors.Count == 0 ? activeReference.Row : anchors[^1].Row);
      var boundaryEndRows = verticalBoundaries.Select(boundary => boundary.EndRow).Distinct().ToArray();
      var candidateEndRow = boundaryEndRows.Length == 1 ? boundaryEndRows[0] : lastRow;
      var horizontalBoundaries = ReadBottomBoundary(worksheet, candidateEndRow, lastColumn);
      var observedLastColumn = horizontalBoundaries.Count == 1
        ? horizontalBoundaries[0].LastColumn
        : lastColumn;
      var horizontalEndRows = horizontalBoundaries.Select(boundary => boundary.Row).Distinct().ToArray();
      var logicalLastRow = boundaryEndRows.Length == 1
        ? boundaryEndRows[0]
        : horizontalEndRows.Length == 1
          ? horizontalEndRows[0]
          : lastRow;

      var cells = ReadOccupiedCells(
        worksheet,
        values,
        formulas,
        comments,
        hyperlinks,
        firstRow,
        firstColumn,
        rowCount,
        columnCount,
        logicalLastRow,
        observedLastColumn);
      var shapes = ReadShapes(worksheet);
      var rowHeights = ReadRowHeights(worksheet, firstAnchorRow, logicalLastRow);
      var columnWidths = ReadColumnWidths(worksheet, NewFirstColumn, observedLastColumn);
      var signals = new SheetLayoutSignals(
        activeReference.Row,
        lastRow,
        logicalLastRow,
        NewFirstColumn,
        observedLastColumn,
        anchors,
        verticalBoundaries,
        horizontalBoundaries,
        oldHeaderColumns);

      if (!WorkbookWindowMatchesIdentity(workbook, identity))
      {
        return SheetSnapshotResult.Failed(
          resolvedName,
          "Workbook identity changed while the worksheet snapshot was being captured.");
      }

      var snapshot = new SheetSnapshot(
        resolvedName,
        activeReference,
        firstRow,
        lastRow,
        firstColumn,
        lastColumn,
        Convert.ToBoolean(GetRequiredProperty(workbook, "ReadOnly"), CultureInfo.InvariantCulture),
        IsWorksheetProtected(worksheet),
        signals,
        cells,
        shapes,
        rowHeights,
        columnWidths);
      return new SheetSnapshotResult(true, snapshot, $"{resolvedName} のライブSnapshotを取得しました。");
    }
    finally
    {
      if (applicationStateRead)
      {
        try
        {
          SetProperty(application, "EnableEvents", eventsWereEnabled);
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          // Excel may have disconnected after capture; no further COM recovery is possible.
        }
      }

      ComRelease.Release(usedColumns);
      ComRelease.Release(usedRows);
      ComRelease.Release(usedRange);
      ComRelease.Release(activeCell);
      ComRelease.Release(window);
      ComRelease.Release(windows);
      ComRelease.Release(workbookActiveSheet);
      ComRelease.Release(worksheet);
    }
  }

  private static IReadOnlyList<CaseAnchorSignal> ReadAnchors(
    object worksheet,
    object? values,
    object? formulas,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount,
    int lastColumn)
  {
    var result = new List<CaseAnchorSignal>();
    for (var row = firstRow; row < firstRow + rowCount; row++)
    {
      var hasA = IsNonEmpty(MatrixValue(values, row, 1, firstRow, firstColumn, rowCount, columnCount)) ||
        IsNonEmpty(MatrixValue(formulas, row, 1, firstRow, firstColumn, rowCount, columnCount));
      var hasB = IsNonEmpty(MatrixValue(values, row, 2, firstRow, firstColumn, rowCount, columnCount)) ||
        IsNonEmpty(MatrixValue(formulas, row, 2, firstRow, firstColumn, rowCount, columnCount));
      if (!hasA && !hasB)
      {
        continue;
      }

      result.Add(new CaseAnchorSignal(
        row,
        hasA,
        hasB,
        result.Count > 0 && HasRangeBorder(worksheet, row, 1, row, lastColumn, XlEdgeTop),
        DisplayValue(MatrixValue(values, row, 1, firstRow, firstColumn, rowCount, columnCount)),
        DisplayValue(MatrixValue(values, row, 2, firstRow, firstColumn, rowCount, columnCount))));
    }

    return result;
  }

  private static IReadOnlyList<int> ReadOldHeaderColumns(
    object? values,
    object? formulas,
    int headerRow,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount,
    int lastColumn)
  {
    if (headerRow < firstRow || headerRow >= firstRow + rowCount)
    {
      return [];
    }

    var result = new List<int>();
    for (var column = NewFirstColumn + 1; column <= lastColumn; column++)
    {
      if (IsNonEmpty(MatrixValue(values, headerRow, column, firstRow, firstColumn, rowCount, columnCount)) ||
        IsNonEmpty(MatrixValue(formulas, headerRow, column, firstRow, firstColumn, rowCount, columnCount)))
      {
        result.Add(column);
      }
    }

    return result;
  }

  private static IReadOnlyList<VerticalBoundarySignal> ReadVerticalBoundaries(
    object worksheet,
    int firstAnchorRow,
    int lastRow,
    int lastColumn,
    IReadOnlyList<int> oldHeaderColumns,
    int minimumEndRow)
  {
    if (lastColumn <= NewFirstColumn || firstAnchorRow > lastRow)
    {
      return [];
    }

    var columns = oldHeaderColumns.Count > 0
      ? oldHeaderColumns.Select(column => column - 1)
      : Enumerable.Range(NewFirstColumn, lastColumn - NewFirstColumn);
    var result = new List<VerticalBoundarySignal>();
    foreach (var column in columns.Distinct().Where(column => column >= NewFirstColumn && column < lastColumn))
    {
      if (!HasCellBorder(worksheet, firstAnchorRow, column, XlEdgeRight))
      {
        continue;
      }

      var endRow = firstAnchorRow;
      while (endRow < lastRow && HasCellBorder(worksheet, endRow + 1, column, XlEdgeRight))
      {
        endRow++;
      }

      if (endRow >= minimumEndRow)
      {
        result.Add(new VerticalBoundarySignal(column, firstAnchorRow, endRow));
      }
    }

    return result;
  }

  private static IReadOnlyList<HorizontalBoundarySignal> ReadBottomBoundary(
    object worksheet,
    int row,
    int lastColumn)
  {
    if (row < 1 || row > ExcelWorksheetLimits.MaximumRow ||
      !HasCellBorder(worksheet, row, 1, XlEdgeBottom))
    {
      return [];
    }

    var endColumn = 1;
    while (endColumn < lastColumn && HasCellBorder(worksheet, row, endColumn + 1, XlEdgeBottom))
    {
      endColumn++;
    }

    return endColumn >= NewFirstColumn + 1
      ? [new HorizontalBoundarySignal(row, 1, endColumn)]
      : [];
  }

  private static IReadOnlyList<SnapshotCell> ReadOccupiedCells(
    object worksheet,
    object? values,
    object? formulas,
    HashSet<(int Row, int Column)> comments,
    HashSet<(int Row, int Column)> hyperlinks,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount,
    int logicalLastRow,
    int evidenceLastColumn)
  {
    var result = new List<SnapshotCell>();
    var startRow = Math.Max(firstRow, 1);
    var endRow = Math.Min(firstRow + rowCount - 1, logicalLastRow);
    var startColumn = Math.Max(firstColumn, NewFirstColumn);
    var endColumn = Math.Min(firstColumn + columnCount - 1, evidenceLastColumn);
    for (var row = startRow; row <= endRow; row++)
    {
      for (var column = startColumn; column <= endColumn; column++)
      {
        var hasContent = IsNonEmpty(
            MatrixValue(values, row, column, firstRow, firstColumn, rowCount, columnCount)) ||
          IsNonEmpty(MatrixValue(formulas, row, column, firstRow, firstColumn, rowCount, columnCount));
        var hasComment = comments.Contains((row, column));
        var hasHyperlink = hyperlinks.Contains((row, column));
        var isMerged = IsMergedCell(worksheet, row, column);
        if (hasContent || hasComment || hasHyperlink || isMerged)
        {
          result.Add(new SnapshotCell(row, column, hasContent, hasComment, hasHyperlink, isMerged));
        }
      }
    }

    return result;
  }

  private static IReadOnlyList<SnapshotShape> ReadShapes(object worksheet)
  {
    object? shapes = null;
    var result = new List<SnapshotShape>();
    try
    {
      shapes = GetRequiredProperty(worksheet, "Shapes");
      var count = ReadInt(shapes, "Count");
      for (var index = 1; index <= count; index++)
      {
        object? shape = null;
        object? topLeft = null;
        object? bottomRight = null;
        try
        {
          shape = InvokeProperty(shapes, "Item", index);
          if (shape is null)
          {
            continue;
          }

          topLeft = GetRequiredProperty(shape, "TopLeftCell");
          bottomRight = GetRequiredProperty(shape, "BottomRightCell");
          var name = Convert.ToString(GetRequiredProperty(shape, "Name"), CultureInfo.CurrentCulture) ?? string.Empty;
          var alternativeText = TryGetProperty(shape, "AlternativeText", out var text)
            ? Convert.ToString(text, CultureInfo.InvariantCulture)
            : null;
          var isManaged = ManagedShapeMetadata.IsManagedName(name) &&
            ManagedShapeMetadata.TryParse(alternativeText, out _);
          result.Add(new SnapshotShape(
            name,
            ReadInt(topLeft, "Row"),
            ReadInt(bottomRight, "Row"),
            ReadInt(topLeft, "Column"),
            ReadInt(bottomRight, "Column"),
            isManaged));
        }
        finally
        {
          ComRelease.Release(bottomRight);
          ComRelease.Release(topLeft);
          ComRelease.Release(shape);
        }
      }

      return result;
    }
    finally
    {
      ComRelease.Release(shapes);
    }
  }

  private static IReadOnlyDictionary<int, double> ReadRowHeights(
    object worksheet,
    int firstRow,
    int lastRow)
  {
    var result = new Dictionary<int, double>();
    for (var row = firstRow; row <= lastRow; row++)
    {
      object? cell = null;
      try
      {
        cell = GetRequiredProperty(worksheet, "Cells", row, 1);
        var height = Convert.ToDouble(GetRequiredProperty(cell, "Height"), CultureInfo.InvariantCulture);
        if (double.IsFinite(height) && height > 0)
        {
          result[row] = height;
        }
      }
      finally
      {
        ComRelease.Release(cell);
      }
    }

    return result;
  }

  private static IReadOnlyDictionary<int, double> ReadColumnWidths(
    object worksheet,
    int firstColumn,
    int lastColumn)
  {
    var result = new Dictionary<int, double>();
    for (var column = firstColumn; column <= lastColumn; column++)
    {
      object? cell = null;
      try
      {
        cell = GetRequiredProperty(worksheet, "Cells", 1, column);
        var width = Convert.ToDouble(GetRequiredProperty(cell, "Width"), CultureInfo.InvariantCulture);
        if (double.IsFinite(width) && width > 0)
        {
          result[column] = width;
        }
      }
      finally
      {
        ComRelease.Release(cell);
      }
    }

    return result;
  }

  private static HashSet<(int Row, int Column)> ReadLinkedCells(
    object worksheet,
    string collectionName,
    bool optional = false)
  {
    object? collection = null;
    var result = new HashSet<(int Row, int Column)>();
    try
    {
      if (optional && (!TryGetProperty(worksheet, collectionName, out collection) || collection is null))
      {
        return result;
      }

      collection ??= GetRequiredProperty(worksheet, collectionName);
      var count = ReadInt(collection, "Count");
      for (var index = 1; index <= count; index++)
      {
        object? item = null;
        object? parent = null;
        try
        {
          item = InvokeProperty(collection, "Item", index);
          if (item is null)
          {
            continue;
          }

          parent = GetRequiredProperty(item, "Parent");
          result.Add((ReadInt(parent, "Row"), ReadInt(parent, "Column")));
        }
        finally
        {
          ComRelease.Release(parent);
          ComRelease.Release(item);
        }
      }

      return result;
    }
    finally
    {
      ComRelease.Release(collection);
    }
  }

  private static bool IsMergedCell(object worksheet, int row, int column)
  {
    object? cell = null;
    try
    {
      cell = GetRequiredProperty(worksheet, "Cells", row, column);
      var value = GetRequiredProperty(cell, "MergeCells");
      return value is not null && Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }
    finally
    {
      ComRelease.Release(cell);
    }
  }

  private static bool HasRangeBorder(
    object worksheet,
    int firstRow,
    int firstColumn,
    int lastRow,
    int lastColumn,
    int borderIndex)
  {
    object? firstCell = null;
    object? lastCell = null;
    object? range = null;
    try
    {
      firstCell = GetRequiredProperty(worksheet, "Cells", firstRow, firstColumn);
      lastCell = GetRequiredProperty(worksheet, "Cells", lastRow, lastColumn);
      range = GetRequiredProperty(worksheet, "Range", firstCell, lastCell);
      return HasBorder(range, borderIndex);
    }
    finally
    {
      ComRelease.Release(range);
      ComRelease.Release(lastCell);
      ComRelease.Release(firstCell);
    }
  }

  private static bool HasCellBorder(object worksheet, int row, int column, int borderIndex)
  {
    object? cell = null;
    try
    {
      cell = GetRequiredProperty(worksheet, "Cells", row, column);
      return HasBorder(cell, borderIndex);
    }
    finally
    {
      ComRelease.Release(cell);
    }
  }

  private static bool HasBorder(object range, int borderIndex)
  {
    object? borders = null;
    object? border = null;
    try
    {
      borders = GetRequiredProperty(range, "Borders");
      border = InvokeProperty(borders, "Item", borderIndex);
      if (border is null || !TryGetProperty(border, "LineStyle", out var lineStyle) || lineStyle is null)
      {
        return false;
      }

      var value = Convert.ToInt32(lineStyle, CultureInfo.InvariantCulture);
      return value != 0 && value != XlLineStyleNone;
    }
    finally
    {
      ComRelease.Release(border);
      ComRelease.Release(borders);
    }
  }

  private static object? MatrixValue(
    object? matrix,
    int row,
    int column,
    int firstRow,
    int firstColumn,
    int rowCount,
    int columnCount)
  {
    if (row < firstRow || row >= firstRow + rowCount ||
      column < firstColumn || column >= firstColumn + columnCount ||
      matrix is null)
    {
      return null;
    }

    if (matrix is not Array array)
    {
      return rowCount == 1 && columnCount == 1 ? matrix : null;
    }

    return array.GetValue(
      array.GetLowerBound(0) + row - firstRow,
      array.GetLowerBound(1) + column - firstColumn);
  }

  private static bool IsNonEmpty(object? value) =>
    value is not null &&
    value is not DBNull &&
    (value is not string text || !string.IsNullOrWhiteSpace(text));

  private static string? DisplayValue(object? value)
  {
    var text = Convert.ToString(value, CultureInfo.CurrentCulture)?.Trim();
    if (string.IsNullOrEmpty(text))
    {
      return null;
    }

    text = string.Join(" ", text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    return text.Length <= 60 ? text : $"{text[..57]}...";
  }

  private static object ResolveWorksheet(
    object application,
    object workbook,
    string worksheetName,
    out string resolvedName)
  {
    if (string.Equals(worksheetName, "ActiveSheet", StringComparison.OrdinalIgnoreCase))
    {
      var activeSheet = GetRequiredProperty(workbook, "ActiveSheet");
      try
      {
        resolvedName = Convert.ToString(
            GetRequiredProperty(activeSheet, "Name"),
            CultureInfo.CurrentCulture) ??
          throw new InvalidOperationException("The active Excel sheet has no name.");
        return activeSheet;
      }
      catch
      {
        ComRelease.Release(activeSheet);
        throw;
      }
    }

    var worksheets = GetRequiredProperty(workbook, "Worksheets");
    try
    {
      var worksheet = InvokeProperty(worksheets, "Item", worksheetName) ??
        throw new InvalidOperationException($"Worksheet was not found: {worksheetName}");
      try
      {
        resolvedName = Convert.ToString(
            GetRequiredProperty(worksheet, "Name"),
            CultureInfo.CurrentCulture) ??
          throw new InvalidOperationException("The selected Excel sheet has no name.");
        return worksheet;
      }
      catch
      {
        ComRelease.Release(worksheet);
        throw;
      }
    }
    finally
    {
      ComRelease.Release(worksheets);
    }
  }

  private static bool IsWorksheetProtected(object worksheet) =>
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectContents"), CultureInfo.InvariantCulture) ||
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectDrawingObjects"), CultureInfo.InvariantCulture) ||
    Convert.ToBoolean(GetRequiredProperty(worksheet, "ProtectScenarios"), CultureInfo.InvariantCulture);

  private static bool ApplicationMatches(object application, WorkbookIdentity identity)
  {
    var windowHandle = new IntPtr(Convert.ToInt64(
      GetRequiredProperty(application, "Hwnd"),
      CultureInfo.InvariantCulture));
    var threadId = NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
    return threadId != 0 && processId == identity.ProcessId;
  }

  private static bool WorkbookMatches(object workbook, WorkbookIdentity identity)
  {
    var name = Convert.ToString(GetRequiredProperty(workbook, "Name"), CultureInfo.CurrentCulture);
    var fullName = Convert.ToString(GetRequiredProperty(workbook, "FullName"), CultureInfo.CurrentCulture);
    return string.Equals(name, identity.Name, StringComparison.OrdinalIgnoreCase) &&
      string.Equals(fullName, identity.FullPath, StringComparison.OrdinalIgnoreCase);
  }

  private static bool WorkbookWindowMatchesIdentity(object workbook, WorkbookIdentity identity)
  {
    if (identity.WindowSessionToken == IntPtr.Zero ||
      !WorkbookSessionTokenRegistry.IsProcessMonitored(identity.ProcessId) ||
      !WorkbookSessionTokenRegistry.IsValid(identity.WindowSessionToken))
    {
      return false;
    }

    object? windows = null;
    object? window = null;
    try
    {
      windows = GetRequiredProperty(workbook, "Windows");
      if (ReadInt(windows, "Count") < 1)
      {
        return false;
      }

      window = InvokeProperty(windows, "Item", 1);
      if (window is null)
      {
        return false;
      }

      var windowHandle = new IntPtr(Convert.ToInt64(
        GetRequiredProperty(window, "Hwnd"),
        CultureInfo.InvariantCulture));
      var desktopWindow = NativeMethods.FindWindowEx(windowHandle, IntPtr.Zero, "XLDESK", null);
      var documentWindowHandle = desktopWindow == IntPtr.Zero
        ? IntPtr.Zero
        : NativeMethods.FindWindowEx(desktopWindow, IntPtr.Zero, "EXCEL7", null);
      NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
      return windowHandle == identity.ExcelWindowHandle &&
        documentWindowHandle == identity.ExcelDocumentWindowHandle &&
        processId == identity.ProcessId &&
        NativeMethods.GetProp(documentWindowHandle, WorkbookSessionTokenRegistry.WindowPropertyName) ==
          identity.WindowSessionToken;
    }
    finally
    {
      ComRelease.Release(window);
      ComRelease.Release(windows);
    }
  }

  private static int ReadInt(object target, string propertyName) =>
    Convert.ToInt32(GetRequiredProperty(target, propertyName), CultureInfo.InvariantCulture);

  private static object GetRequiredProperty(object target, string propertyName, params object[] arguments) =>
    InvokeProperty(target, propertyName, arguments) ??
    throw new InvalidOperationException($"COM property returned null: {propertyName}");

  private static bool TryGetProperty(object target, string propertyName, out object? value)
  {
    try
    {
      value = InvokeProperty(target, propertyName);
      return true;
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      value = null;
      return false;
    }
  }

  private static object? InvokeProperty(object target, string propertyName, params object[] arguments) =>
    target.GetType().InvokeMember(
      propertyName,
      BindingFlags.GetProperty,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static object? InvokeMethod(object target, string methodName, params object[] arguments) =>
    target.GetType().InvokeMember(
      methodName,
      BindingFlags.InvokeMethod,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static void SetProperty(object target, string propertyName, object value) =>
    _ = target.GetType().InvokeMember(
      propertyName,
      BindingFlags.SetProperty,
      binder: null,
      target,
      [value],
      CultureInfo.CurrentCulture);

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException;

  private static int GetAutomationHResult(Exception exception) =>
    exception is TargetInvocationException { InnerException: not null } invocationException
      ? invocationException.InnerException!.HResult
      : exception.HResult;

  private static class NativeMethods
  {
    [DllImport("ole32.dll")]
    internal static extern int GetRunningObjectTable(
      int reserved,
      [MarshalAs(UnmanagedType.Interface)] out IRunningObjectTable runningObjectTable);

    [DllImport("ole32.dll")]
    internal static extern int CreateBindCtx(
      int reserved,
      [MarshalAs(UnmanagedType.Interface)] out IBindCtx bindContext);

    [DllImport("user32.dll")]
    internal static extern uint GetWindowThreadProcessId(nint windowHandle, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint FindWindowEx(
      nint parentWindow,
      nint childAfter,
      string className,
      string? windowName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint GetProp(nint windowHandle, string propertyName);
  }
}

public sealed record SheetSnapshotResult(
  bool Succeeded,
  SheetSnapshot? Snapshot,
  string Message)
{
  public static SheetSnapshotResult Failed(string worksheetName, string message) =>
    new(false, null, $"{worksheetName}: {message}");
}
