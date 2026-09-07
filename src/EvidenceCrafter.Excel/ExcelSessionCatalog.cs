using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using System.Text;
using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Excel;

/// <summary>
/// Discovers running Excel Workbooks without launching Excel, retaining COM objects, or requiring Office PIAs.
/// </summary>
public sealed class ExcelSessionCatalog : IExcelSessionCatalog
{
  public ExcelDiscoveryResult Discover()
  {
    var warnings = new List<string>();
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      warnings.Add("Excel discovery must run on an STA thread.");
      return new ExcelDiscoveryResult([], warnings);
    }

    IRunningObjectTable? runningObjectTable = null;
    IEnumMoniker? monikerEnumerator = null;
    IBindCtx? bindContext = null;
    try
    {
      Marshal.ThrowExceptionForHR(NativeMethods.GetRunningObjectTable(0, out runningObjectTable));
      runningObjectTable.EnumRunning(out monikerEnumerator);
      Marshal.ThrowExceptionForHR(NativeMethods.CreateBindCtx(0, out bindContext));

      var discovered = new Dictionary<string, DiscoveredWorkbook>(StringComparer.OrdinalIgnoreCase);
      var monikers = new IMoniker[1];
      while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0)
      {
        object? runningObject = null;
        try
        {
          runningObjectTable.GetObject(monikers[0], out runningObject);
          if (runningObject is not null)
          {
            var registrationDisplayName = GetRegistrationDisplayName(monikers[0], bindContext);
            InspectRunningObject(runningObject, registrationDisplayName, discovered, warnings);
          }
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          warnings.Add($"A running object disconnected during Excel discovery (0x{GetAutomationHResult(exception):X8}).");
        }
        finally
        {
          ComRelease.Release(runningObject);
          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }

      var workbooks = discovered.Values
        .Select(item => item.Identity)
        .OrderBy(item => item.ProcessId)
        .ThenBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase)
        .ThenBy(item => item.FullPath, StringComparer.OrdinalIgnoreCase)
        .ToArray();
      return new ExcelDiscoveryResult(workbooks, warnings);
    }
    catch (COMException exception)
    {
      warnings.Add($"Excel Running Object Table access failed (0x{exception.HResult:X8}).");
      return new ExcelDiscoveryResult([], warnings);
    }
    finally
    {
      ComRelease.Release(bindContext);
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }
  }

  private static void InspectRunningObject(
    object runningObject,
    string registrationDisplayName,
    IDictionary<string, DiscoveredWorkbook> discovered,
    ICollection<string> warnings)
  {
    if (TryGetProperty(runningObject, "Workbooks", out var workbooks))
    {
      try
      {
        AddApplication(runningObject, workbooks, registrationDisplayName, discovered, warnings);
      }
      finally
      {
        ComRelease.Release(workbooks);
      }

      return;
    }

    if (TryGetProperty(runningObject, "Worksheets", out var worksheets))
    {
      try
      {
        if (TryGetProperty(runningObject, "Application", out var application))
        {
          try
          {
            AddWorkbookRoot(runningObject, application, registrationDisplayName, discovered, warnings);
          }
          finally
          {
            ComRelease.Release(application);
          }
        }
      }
      finally
      {
        ComRelease.Release(worksheets);
      }
    }
  }

  private static void AddWorkbookRoot(
    object workbook,
    object? application,
    string registrationDisplayName,
    IDictionary<string, DiscoveredWorkbook> discovered,
    ICollection<string> warnings)
  {
    if (application is null)
    {
      return;
    }

    try
    {
      AddWorkbook(workbook, registrationDisplayName, hasWorkbookRegistration: true, discovered, warnings);
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      warnings.Add($"A Workbook could not be inspected (0x{GetAutomationHResult(exception):X8}).");
    }
  }

  private static void AddApplication(
    object application,
    object? workbooks,
    string registrationDisplayName,
    IDictionary<string, DiscoveredWorkbook> discovered,
    ICollection<string> warnings)
  {
    if (workbooks is null)
    {
      return;
    }

    var count = Convert.ToInt32(GetRequiredProperty(workbooks, "Count"), CultureInfo.InvariantCulture);
    for (var index = 1; index <= count; index++)
    {
      object? workbook = null;
      try
      {
        workbook = InvokeProperty(workbooks, "Item", index);
        if (workbook is not null)
        {
          AddWorkbook(workbook, registrationDisplayName, hasWorkbookRegistration: false, discovered, warnings);
        }
      }
      catch (Exception exception) when (IsAutomationFailure(exception))
      {
        warnings.Add($"Workbook #{index} disconnected during discovery (0x{GetAutomationHResult(exception):X8}).");
      }
      finally
      {
        ComRelease.Release(workbook);
      }
    }
  }

  private static void AddWorkbook(
    object workbook,
    string registrationDisplayName,
    bool hasWorkbookRegistration,
    IDictionary<string, DiscoveredWorkbook> discovered,
    ICollection<string> warnings)
  {
    var name = Convert.ToString(GetRequiredProperty(workbook, "Name"), CultureInfo.CurrentCulture) ??
      "Unnamed workbook";
    var fullPath = TryGetProperty(workbook, "FullName", out var fullNameValue)
      ? Convert.ToString(fullNameValue, CultureInfo.CurrentCulture) ?? name
      : name;
    var isReadOnly = TryGetProperty(workbook, "ReadOnly", out var readOnlyValue) &&
      Convert.ToBoolean(readOnlyValue, CultureInfo.InvariantCulture);

    if (!TryGetWorkbookWindowIdentity(
      workbook,
      out var windowHandle,
      out var documentWindowHandle,
      out var processId))
    {
      warnings.Add($"Workbook '{name}' has no valid Excel window and was skipped.");
      return;
    }

    var stableKey = $"{processId}|{fullPath}|{name}";
    var windowSessionToken = GetOrCreateWindowSessionToken(documentWindowHandle, stableKey, processId);
    if (windowSessionToken == IntPtr.Zero)
    {
      warnings.Add($"Workbook '{name}' could not be assigned a window session and was skipped.");
      return;
    }

    var connectionId = CreateConnectionId(
      windowHandle,
      processId,
      fullPath,
      name,
      windowSessionToken,
      registrationDisplayName,
      documentWindowHandle);
    var identity = new WorkbookIdentity(
      connectionId,
      name,
      fullPath,
      windowHandle,
      processId,
      isReadOnly,
      windowSessionToken,
      registrationDisplayName,
      hasWorkbookRegistration,
      documentWindowHandle);
    if (!discovered.TryGetValue(stableKey, out var existing) ||
      hasWorkbookRegistration && !existing.HasWorkbookRegistration)
    {
      discovered[stableKey] = new DiscoveredWorkbook(identity, hasWorkbookRegistration);
    }
  }

  private static string GetRegistrationDisplayName(IMoniker moniker, IBindCtx bindContext)
  {
    moniker.GetDisplayName(bindContext, null, out var displayName);
    return displayName;
  }

  private static nint GetOrCreateWindowSessionToken(nint windowHandle, string workbookKey, uint processId)
  {
    return WorkbookSessionTokenRegistry.GetOrCreateWindowSessionToken(windowHandle, workbookKey, processId);
  }

  private static bool TryGetWorkbookWindowIdentity(
    object workbook,
    out nint windowHandle,
    out nint documentWindowHandle,
    out uint processId)
  {
    object? windows = null;
    object? window = null;
    windowHandle = IntPtr.Zero;
    documentWindowHandle = IntPtr.Zero;
    processId = 0;
    try
    {
      windows = GetRequiredProperty(workbook, "Windows");
      var count = Convert.ToInt32(GetRequiredProperty(windows, "Count"), CultureInfo.InvariantCulture);
      if (count < 1)
      {
        return false;
      }

      window = InvokeProperty(windows, "Item", 1);
      if (window is null)
      {
        return false;
      }

      windowHandle = GetWindowHandle(window);
      var desktopWindow = NativeMethods.FindWindowEx(windowHandle, IntPtr.Zero, "XLDESK", null);
      documentWindowHandle = desktopWindow == IntPtr.Zero
        ? IntPtr.Zero
        : NativeMethods.FindWindowEx(desktopWindow, IntPtr.Zero, "EXCEL7", null);
      var threadId = windowHandle == IntPtr.Zero
        ? 0
        : NativeMethods.GetWindowThreadProcessId(windowHandle, out processId);
      return threadId != 0 && processId != 0 && documentWindowHandle != IntPtr.Zero;
    }
    finally
    {
      ComRelease.Release(window);
      ComRelease.Release(windows);
    }
  }

  private static nint GetWindowHandle(object application)
  {
    var value = GetRequiredProperty(application, "Hwnd");
    return new IntPtr(Convert.ToInt64(value, CultureInfo.InvariantCulture));
  }

  private static object GetRequiredProperty(object target, string propertyName) =>
    InvokeProperty(target, propertyName) ??
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

  private static object? InvokeProperty(object target, string propertyName, params object[]? arguments) =>
    target.GetType().InvokeMember(
      propertyName,
      BindingFlags.GetProperty,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or InvalidOperationException;

  private static int GetAutomationHResult(Exception exception) =>
    exception is TargetInvocationException { InnerException: not null } invocationException
      ? invocationException.InnerException!.HResult
      : exception.HResult;

  private static string CreateConnectionId(
    nint windowHandle,
    uint processId,
    string fullPath,
    string name,
    nint windowSessionToken,
    string registrationDisplayName,
    nint documentWindowHandle)
  {
    var value =
      $"{windowHandle}|{documentWindowHandle}|{processId}|{fullPath}|{name}|{windowSessionToken}|" +
      registrationDisplayName;
    var hash = SHA256.HashData(Encoding.UTF8.GetBytes(value));
    return Convert.ToHexString(hash);
  }

  private sealed record DiscoveredWorkbook(WorkbookIdentity Identity, bool HasWorkbookRegistration);

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

  }
}
