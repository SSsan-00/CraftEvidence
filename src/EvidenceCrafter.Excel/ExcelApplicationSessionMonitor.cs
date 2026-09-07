using System.Collections.Concurrent;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EvidenceCrafter.Core.Models;

namespace EvidenceCrafter.Excel;

/// <summary>Confirms Workbook window closure before invalidating its session token.</summary>
public sealed class ExcelApplicationSessionMonitor : IDisposable
{
  private const int WorkbookBeforeCloseDispId = 1570;
  private const int SheetSelectionChangeDispId = 1558;
  private static readonly TimeSpan WorkbookClosePollInterval = TimeSpan.FromMilliseconds(250);
  private static readonly Guid AppEventsInterfaceId = new("00024413-0000-0000-C000-000000000046");
  private readonly CancellationTokenSource closeWatchCancellation = new();
  private readonly ConcurrentDictionary<nint, PendingWorkbookClose> pendingWorkbookCloses = new();
  private readonly Dictionary<uint, Subscription> subscriptions = [];
  private SynchronizationContext? ownerContext;
  private int closeWatcherRunning;
  private volatile bool disposed;

  public int CloseEventCount { get; private set; }

  public event EventHandler<ExcelSelectionChangedEventArgs>? SelectionChanged;

  /// <summary>Immediately invalidates identities when monitoring cannot be trusted.</summary>
  public static void InvalidateSessions(IReadOnlyList<WorkbookIdentity> workbooks)
  {
    ArgumentNullException.ThrowIfNull(workbooks);
    foreach (var processId in workbooks.Select(item => item.ProcessId).Distinct())
    {
      WorkbookSessionTokenRegistry.MarkProcessUnmonitored(processId);
    }
  }

  /// <summary>Stops new session registrations and synchronously releases every native session property.</summary>
  public static void BeginShutdownSessions() => WorkbookSessionTokenRegistry.BeginShutdown();

  public IReadOnlyList<string> Refresh(IReadOnlyList<WorkbookIdentity> workbooks)
  {
    ObjectDisposedException.ThrowIf(disposed, this);
    ArgumentNullException.ThrowIfNull(workbooks);
    if (Thread.CurrentThread.GetApartmentState() is not ApartmentState.STA)
    {
      return ["Excel session monitoring must run on its owning STA thread."];
    }

    ownerContext ??= SynchronizationContext.Current;
    var targetProcessIds = workbooks.Select(item => item.ProcessId).Where(id => id != 0).ToHashSet();
    var warnings = new List<string>();
    var failedProcessIds = new HashSet<uint>();
    foreach (var processId in subscriptions.Keys.Where(id => !targetProcessIds.Contains(id)).ToArray())
    {
      Detach(processId);
    }

    foreach (var subscription in subscriptions.ToArray())
    {
      try
      {
        if (!GetApplicationEventsEnabled(subscription.Value.Application))
        {
          warnings.Add($"Excel events are disabled for process {subscription.Key}; close monitoring is paused and direct validation will be used.");
          failedProcessIds.Add(subscription.Key);
          Detach(subscription.Key, invalidateTokens: false);
        }
        else
        {
          WorkbookSessionTokenRegistry.RegisterMonitoredProcess(subscription.Key);
        }
      }
      catch (Exception exception) when (IsAutomationFailure(exception))
      {
        warnings.Add(
          $"Excel session monitoring validation failed for process {subscription.Key} " +
          $"(0x{GetAutomationHResult(exception):X8}).");
        failedProcessIds.Add(subscription.Key);
        Detach(subscription.Key, invalidateTokens: false);
      }
    }

    IRunningObjectTable? runningObjectTable = null;
    IEnumMoniker? monikerEnumerator = null;
    try
    {
      Marshal.ThrowExceptionForHR(NativeMethods.GetRunningObjectTable(0, out runningObjectTable));
      runningObjectTable.EnumRunning(out monikerEnumerator);
      var monikers = new IMoniker[1];
      while (monikerEnumerator.Next(1, monikers, IntPtr.Zero) == 0)
      {
        object? runningObject = null;
        object? application = null;
        IConnectionPoint? connectionPoint = null;
        var retainedApplication = false;
        uint candidateProcessId = 0;
        try
        {
          runningObjectTable.GetObject(monikers[0], out runningObject);
          if (runningObject is null)
          {
            continue;
          }

          if (TryGetProperty(runningObject, "Workbooks", out var collection))
          {
            ComRelease.Release(collection);
            application = runningObject;
          }
          else if (TryGetProperty(runningObject, "Worksheets", out var worksheets))
          {
            ComRelease.Release(worksheets);
            _ = TryGetProperty(runningObject, "Application", out application);
          }

          if (application is null)
          {
            continue;
          }

          candidateProcessId = GetApplicationProcessId(application);
          if (!targetProcessIds.Contains(candidateProcessId) ||
            failedProcessIds.Contains(candidateProcessId) ||
            subscriptions.ContainsKey(candidateProcessId))
          {
            continue;
          }

          if (!GetApplicationEventsEnabled(application))
          {
            warnings.Add($"Excel events are disabled for process {candidateProcessId}; close monitoring is paused and direct validation will be used.");
            failedProcessIds.Add(candidateProcessId);
            WorkbookSessionTokenRegistry.MarkProcessUnmonitored(candidateProcessId);
            continue;
          }

          var connectionPointContainer = (IConnectionPointContainer)application;
          var appEventsInterface = AppEventsInterfaceId;
          connectionPointContainer.FindConnectionPoint(ref appEventsInterface, out connectionPoint);
          if (connectionPoint is null)
          {
            throw new InvalidOperationException("Excel AppEvents connection point was not found.");
          }

          var sink = new AppEventsSink(
            (object workbook, ref bool cancel) => OnWorkbookBeforeClose(candidateProcessId, workbook),
            (sheet, target) => OnSheetSelectionChange(candidateProcessId, sheet, target));
          connectionPoint.Advise(sink, out var cookie);
          subscriptions[candidateProcessId] = new Subscription(application, connectionPoint, cookie, sink);
          WorkbookSessionTokenRegistry.RegisterMonitoredProcess(candidateProcessId);
          connectionPoint = null;
          retainedApplication = true;
        }
        catch (Exception exception) when (IsAutomationFailure(exception))
        {
          warnings.Add($"Excel session monitoring failed (0x{GetAutomationHResult(exception):X8}).");
          if (candidateProcessId != 0 && targetProcessIds.Contains(candidateProcessId))
          {
            failedProcessIds.Add(candidateProcessId);
            WorkbookSessionTokenRegistry.MarkProcessUnmonitored(candidateProcessId);
          }
        }
        finally
        {
          ComRelease.Release(connectionPoint);
          if (!retainedApplication && !ReferenceEquals(application, runningObject))
          {
            ComRelease.Release(application);
          }

          if (!retainedApplication || !ReferenceEquals(application, runningObject))
          {
            ComRelease.Release(runningObject);
          }

          ComRelease.Release(monikers[0]);
          monikers[0] = null!;
        }
      }
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      warnings.Add($"Excel session monitor refresh failed (0x{GetAutomationHResult(exception):X8}).");
    }
    finally
    {
      ComRelease.Release(monikerEnumerator);
      ComRelease.Release(runningObjectTable);
    }

    foreach (var processId in targetProcessIds.Where(id =>
      !subscriptions.ContainsKey(id) && !failedProcessIds.Contains(id)))
    {
      warnings.Add($"Excel session monitoring is unavailable for process {processId}; direct validation remains available.");
      WorkbookSessionTokenRegistry.MarkProcessUnmonitored(processId);
    }

    return warnings;
  }

  public void Dispose()
  {
    if (disposed)
    {
      return;
    }

    disposed = true;
    closeWatchCancellation.Cancel();
    pendingWorkbookCloses.Clear();
    foreach (var processId in subscriptions.Keys.ToArray())
    {
      Detach(processId);
    }

    closeWatchCancellation.Dispose();
  }

  private void OnWorkbookBeforeClose(uint processId, object workbook)
  {
    CloseEventCount++;
    try
    {
      var windowSession = TryGetWorkbookWindowSession(workbook);
      if (windowSession.Token == IntPtr.Zero || windowSession.DocumentWindowHandle == IntPtr.Zero)
      {
        FailClosed(processId);
        return;
      }

      var pendingClose = new PendingWorkbookClose(processId, windowSession);
      if (pendingWorkbookCloses.TryAdd(windowSession.Token, pendingClose))
      {
        EnsureCloseWatcher();
      }
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      FailClosed(processId);
    }
  }

  private void EnsureCloseWatcher()
  {
    if (Interlocked.CompareExchange(ref closeWatcherRunning, 1, 0) == 0)
    {
      _ = WatchPendingWorkbookClosesAsync(closeWatchCancellation.Token);
    }
  }

  private async Task WatchPendingWorkbookClosesAsync(CancellationToken cancellationToken)
  {
    try
    {
      while (!cancellationToken.IsCancellationRequested)
      {
        foreach (var item in pendingWorkbookCloses.ToArray())
        {
          var pendingClose = item.Value;
          if (NativeMethods.IsWindow(pendingClose.WindowSession.DocumentWindowHandle) &&
            NativeMethods.GetProp(
              pendingClose.WindowSession.DocumentWindowHandle,
              WorkbookSessionTokenRegistry.WindowPropertyName) == pendingClose.WindowSession.Token)
          {
            continue;
          }

          if (pendingWorkbookCloses.TryRemove(item.Key, out _))
          {
            WorkbookSessionTokenRegistry.Invalidate(pendingClose.WindowSession.Token);
            PostDetachIfNoWorkbooks(pendingClose.ProcessId);
          }
        }

        if (pendingWorkbookCloses.IsEmpty)
        {
          return;
        }

        await Task.Delay(WorkbookClosePollInterval, cancellationToken).ConfigureAwait(false);
      }
    }
    catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
    {
      // Monitor disposal cancels pending close confirmation without changing a still-open Workbook identity.
    }
    finally
    {
      _ = Interlocked.Exchange(ref closeWatcherRunning, 0);
      if (!cancellationToken.IsCancellationRequested && !pendingWorkbookCloses.IsEmpty)
      {
        EnsureCloseWatcher();
      }
    }
  }

  private void PostDetachIfNoWorkbooks(uint processId)
  {
    var context = ownerContext;
    if (disposed || context is null)
    {
      return;
    }

    try
    {
      context.Post(_ => DetachIfNoWorkbooks(processId), null);
    }
    catch (Exception exception) when (exception is InvalidAsynchronousStateException or InvalidOperationException)
    {
      WorkbookSessionTokenRegistry.UnregisterMonitoredProcess(processId);
    }
  }

  private void DetachIfNoWorkbooks(uint processId)
  {
    if (disposed || !subscriptions.TryGetValue(processId, out var subscription))
    {
      return;
    }

    object? workbooks = null;
    try
    {
      workbooks = GetRequiredProperty(subscription.Application, "Workbooks");
      if (Convert.ToInt32(
        GetRequiredProperty(workbooks, "Count"),
        CultureInfo.InvariantCulture) == 0)
      {
        Detach(processId);
      }
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      FailClosed(processId);
    }
    finally
    {
      ComRelease.Release(workbooks);
    }
  }

  private void Detach(uint processId, bool invalidateTokens = true)
  {
    if (!subscriptions.Remove(processId, out var subscription))
    {
      return;
    }

    if (invalidateTokens)
    {
      WorkbookSessionTokenRegistry.UnregisterMonitoredProcess(processId);
    }
    else
    {
      WorkbookSessionTokenRegistry.MarkProcessUnmonitored(processId);
    }

    try
    {
      subscription.ConnectionPoint.Unadvise(subscription.Cookie);
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // The Excel instance may already be disconnected.
    }
    finally
    {
      ComRelease.Release(subscription.ConnectionPoint);
      ComRelease.Release(subscription.Application);
    }
  }

  private void FailClosed(uint processId)
  {
    WorkbookSessionTokenRegistry.UnregisterMonitoredProcess(processId);
    if (ownerContext is null)
    {
      Detach(processId);
    }
    else
    {
      ownerContext.Post(_ => Detach(processId), null);
    }
  }

  private static WorkbookWindowSession TryGetWorkbookWindowSession(object workbook)
  {
    object? windows = null;
    object? window = null;
    try
    {
      windows = GetRequiredProperty(workbook, "Windows");
      window = InvokeProperty(windows, "Item", 1);
      if (window is null)
      {
        return new WorkbookWindowSession(IntPtr.Zero, IntPtr.Zero);
      }

      var windowHandle = new IntPtr(Convert.ToInt64(
        GetRequiredProperty(window, "Hwnd"),
        CultureInfo.InvariantCulture));
      var desktopWindow = NativeMethods.FindWindowEx(windowHandle, IntPtr.Zero, "XLDESK", null);
      var documentWindow = desktopWindow == IntPtr.Zero
        ? IntPtr.Zero
        : NativeMethods.FindWindowEx(desktopWindow, IntPtr.Zero, "EXCEL7", null);
      var token = documentWindow == IntPtr.Zero
        ? IntPtr.Zero
        : NativeMethods.GetProp(documentWindow, WorkbookSessionTokenRegistry.WindowPropertyName);
      return new WorkbookWindowSession(documentWindow, token);
    }
    finally
    {
      ComRelease.Release(window);
      ComRelease.Release(windows);
    }
  }

  private static uint GetApplicationProcessId(object application)
  {
    var windowHandle = new IntPtr(Convert.ToInt64(
      GetRequiredProperty(application, "Hwnd"),
      CultureInfo.InvariantCulture));
    var threadId = NativeMethods.GetWindowThreadProcessId(windowHandle, out var processId);
    return threadId == 0 ? 0 : processId;
  }

  private static bool GetApplicationEventsEnabled(object application) =>
    Convert.ToBoolean(GetRequiredProperty(application, "EnableEvents"), CultureInfo.InvariantCulture);

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

  private static object? InvokeProperty(object target, string propertyName, params object[] arguments) =>
    target.GetType().InvokeMember(
      propertyName,
      BindingFlags.GetProperty,
      binder: null,
      target,
      arguments,
      CultureInfo.CurrentCulture);

  private static bool IsAutomationFailure(Exception exception) =>
    exception is COMException or TargetInvocationException or MissingMemberException or
      InvalidOperationException or InvalidCastException or ArgumentException;

  private static int GetAutomationHResult(Exception exception) =>
    exception is TargetInvocationException { InnerException: not null } invocationException
      ? invocationException.InnerException!.HResult
      : exception.HResult;

  [ComVisible(true)]
  [Guid("00024413-0000-0000-C000-000000000046")]
  [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
  public interface IExcelAppEvents
  {
    [DispId(SheetSelectionChangeDispId)]
    void SheetSelectionChange(
      [MarshalAs(UnmanagedType.IDispatch)] object sheet,
      [MarshalAs(UnmanagedType.IDispatch)] object target);

    [DispId(WorkbookBeforeCloseDispId)]
    void WorkbookBeforeClose(
      [MarshalAs(UnmanagedType.IDispatch)] object workbook,
      [In, Out] ref bool cancel);
  }

  public delegate void WorkbookBeforeCloseHandler(object workbook, ref bool cancel);

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.None)]
  public sealed class AppEventsSink : IExcelAppEvents
  {
    private readonly WorkbookBeforeCloseHandler workbookBeforeClose;
    private readonly Action<object, object>? sheetSelectionChange;

    public AppEventsSink(WorkbookBeforeCloseHandler workbookBeforeClose)
      : this(workbookBeforeClose, null)
    {
    }

    public AppEventsSink(
      WorkbookBeforeCloseHandler workbookBeforeClose,
      Action<object, object>? sheetSelectionChange)
    {
      this.workbookBeforeClose = workbookBeforeClose;
      this.sheetSelectionChange = sheetSelectionChange;
    }

    public void SheetSelectionChange(object sheet, object target) =>
      sheetSelectionChange?.Invoke(sheet, target);

    public void WorkbookBeforeClose(object workbook, ref bool cancel) => workbookBeforeClose(workbook, ref cancel);
  }

  private void OnSheetSelectionChange(uint processId, object sheet, object target)
  {
    object? workbook = null;
    try
    {
      workbook = GetRequiredProperty(sheet, "Parent");
      SelectionChanged?.Invoke(this, new ExcelSelectionChangedEventArgs(
        processId,
        Convert.ToString(GetRequiredProperty(workbook, "FullName"), CultureInfo.CurrentCulture) ?? string.Empty,
        Convert.ToString(GetRequiredProperty(sheet, "Name"), CultureInfo.CurrentCulture) ?? string.Empty,
        Convert.ToInt32(GetRequiredProperty(target, "Row"), CultureInfo.InvariantCulture),
        Convert.ToInt32(GetRequiredProperty(target, "Column"), CultureInfo.InvariantCulture)));
    }
    catch (Exception exception) when (IsAutomationFailure(exception))
    {
      // A selection event can race with Workbook close. The next refresh will reconcile state.
    }
    finally
    {
      ComRelease.Release(workbook);
    }
  }

  private sealed record WorkbookWindowSession(nint DocumentWindowHandle, nint Token);

  private sealed record PendingWorkbookClose(
    uint ProcessId,
    WorkbookWindowSession WindowSession);

  private sealed record Subscription(
    object Application,
    IConnectionPoint ConnectionPoint,
    int Cookie,
    AppEventsSink Sink);

  private static class NativeMethods
  {
    [DllImport("ole32.dll")]
    internal static extern int GetRunningObjectTable(
      int reserved,
      [MarshalAs(UnmanagedType.Interface)] out IRunningObjectTable runningObjectTable);

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

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool IsWindow(nint windowHandle);
  }
}

public sealed class ExcelSelectionChangedEventArgs(
  uint processId,
  string workbookFullPath = "",
  string worksheetName = "",
  int row = 0,
  int column = 0) : EventArgs
{
  public uint ProcessId { get; } = processId;
  public string WorkbookFullPath { get; } = workbookFullPath;
  public string WorksheetName { get; } = worksheetName;
  public int Row { get; } = row;
  public int Column { get; } = column;
}
