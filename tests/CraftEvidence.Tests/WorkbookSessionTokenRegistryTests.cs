using System.Runtime.InteropServices;
using CraftEvidence.Core.Models;
using CraftEvidence.Excel;

namespace CraftEvidence.Tests;

[TestClass]
[DoNotParallelize]
public sealed class WorkbookSessionTokenRegistryTests
{
  private static int nextProcessId = int.MaxValue;

  [TestMethod]
  public void UnregisterMonitoredProcess_InvalidatesEveryTokenForTheProcess()
  {
    var processId = NextProcessId();
    var firstToken = new IntPtr(Interlocked.Decrement(ref nextProcessId));
    var secondToken = new IntPtr(Interlocked.Decrement(ref nextProcessId));
    WorkbookSessionTokenRegistry.Register($"{processId}|first", processId, firstToken);
    WorkbookSessionTokenRegistry.Register($"{processId}|second", processId, secondToken);
    WorkbookSessionTokenRegistry.RegisterMonitoredProcess(processId);

    WorkbookSessionTokenRegistry.UnregisterMonitoredProcess(processId);

    Assert.IsFalse(WorkbookSessionTokenRegistry.IsProcessMonitored(processId));
    Assert.IsFalse(WorkbookSessionTokenRegistry.IsValid(firstToken));
    Assert.IsFalse(WorkbookSessionTokenRegistry.IsValid(secondToken));
  }

  [TestMethod]
  public void InvalidateSessions_DisablesOptionalMonitoringButPreservesTokens()
  {
    var processId = NextProcessId();
    var token = new IntPtr(Interlocked.Decrement(ref nextProcessId));
    WorkbookSessionTokenRegistry.Register($"{processId}|target", processId, token);
    WorkbookSessionTokenRegistry.RegisterMonitoredProcess(processId);
    WorkbookIdentity[] workbooks =
    [
      new("connection", "target.xlsx", "C:\\temp\\target.xlsx", IntPtr.Zero, processId, false),
    ];

    ExcelApplicationSessionMonitor.InvalidateSessions(workbooks);

    Assert.IsFalse(WorkbookSessionTokenRegistry.IsProcessMonitored(processId));
    Assert.IsTrue(WorkbookSessionTokenRegistry.IsValid(token));
    WorkbookSessionTokenRegistry.InvalidateProcess(processId);
  }

  [TestMethod]
  public void InvalidateWorkbook_ReleasesAllRegistryStateDuringRepeatedOpenClose()
  {
    var processId = NextProcessId();
    WorkbookSessionTokenRegistry.RegisterMonitoredProcess(processId);
    for (var index = 0; index < 10_000; index++)
    {
      var workbookKey = $"{processId}|workbook-{index}";
      var token = new IntPtr(Interlocked.Decrement(ref nextProcessId));
      WorkbookSessionTokenRegistry.Register(workbookKey, processId, token);
      WorkbookSessionTokenRegistry.Invalidate(workbookKey);
      Assert.IsFalse(WorkbookSessionTokenRegistry.IsValid(token));
    }

    var state = WorkbookSessionTokenRegistry.GetState(processId);
    Assert.AreEqual(0, state.ValidTokenCount);
    Assert.AreEqual(0, state.ProcessTokenCount);
    Assert.AreEqual(0, state.WorkbookCount);
    WorkbookSessionTokenRegistry.UnregisterMonitoredProcess(processId);
  }

  [TestMethod]
  public void Register_ReplacesPreviousWorkbookTokenWithoutRetainingIt()
  {
    var processId = NextProcessId();
    var workbookKey = $"{processId}|target";
    var firstToken = new IntPtr(Interlocked.Decrement(ref nextProcessId));
    var replacementToken = new IntPtr(Interlocked.Decrement(ref nextProcessId));
    WorkbookSessionTokenRegistry.Register(workbookKey, processId, firstToken);

    WorkbookSessionTokenRegistry.Register(workbookKey, processId, replacementToken);

    Assert.IsFalse(WorkbookSessionTokenRegistry.IsValid(firstToken));
    Assert.IsTrue(WorkbookSessionTokenRegistry.IsValid(replacementToken));
    var state = WorkbookSessionTokenRegistry.GetState(processId);
    Assert.AreEqual(1, state.ValidTokenCount);
    Assert.AreEqual(1, state.ProcessTokenCount);
    Assert.AreEqual(1, state.WorkbookCount);
    WorkbookSessionTokenRegistry.InvalidateProcess(processId);
  }

  [TestMethod]
  public void AppEventsSink_PreservesCancelValuePassedByAnotherSubscriber()
  {
    var callbackCalled = false;
    var sink = new ExcelApplicationSessionMonitor.AppEventsSink((object _, ref bool cancel) =>
    {
      callbackCalled = true;
    });
    var cancel = true;

    sink.WorkbookBeforeClose(new object(), ref cancel);

    Assert.IsTrue(callbackCalled);
    Assert.IsTrue(cancel);
  }

  [TestMethod]
  public void RegisterAndInvalidate_KeepReassignedPropertyThenRemoveIt()
  {
    var processId = NextProcessId();
    var token = new IntPtr(Interlocked.Decrement(ref nextProcessId));
    var windowHandle = NativeMethods.CreateWindowEx(
      0,
      "STATIC",
      string.Empty,
      0,
      0,
      0,
      0,
      0,
      new IntPtr(-3),
      IntPtr.Zero,
      IntPtr.Zero,
      IntPtr.Zero);
    Assert.AreNotEqual(IntPtr.Zero, windowHandle);
    try
    {
      Assert.IsTrue(NativeMethods.SetProp(
        windowHandle,
        WorkbookSessionTokenRegistry.WindowPropertyName,
        token));
      WorkbookSessionTokenRegistry.Register($"{processId}|old", processId, token, windowHandle);

      WorkbookSessionTokenRegistry.Register($"{processId}|new", processId, token, windowHandle);

      Assert.AreEqual(
        token,
        NativeMethods.GetProp(windowHandle, WorkbookSessionTokenRegistry.WindowPropertyName),
        "Reassigning the same live token to a new Workbook key must preserve its window property.");
      WorkbookSessionTokenRegistry.Invalidate(token);
      Assert.AreEqual(
        IntPtr.Zero,
        NativeMethods.GetProp(windowHandle, WorkbookSessionTokenRegistry.WindowPropertyName));
    }
    finally
    {
      WorkbookSessionTokenRegistry.Invalidate(token);
      _ = NativeMethods.RemoveProp(windowHandle, WorkbookSessionTokenRegistry.WindowPropertyName);
      _ = NativeMethods.DestroyWindow(windowHandle);
    }
  }

  [TestMethod]
  public void GetOrCreateWindowSessionToken_RegistersAndReusesTheNativeWindowToken()
  {
    var processId = NextProcessId();
    var windowHandle = NativeMethods.CreateWindowEx(
      0,
      "STATIC",
      string.Empty,
      0,
      0,
      0,
      0,
      0,
      new IntPtr(-3),
      IntPtr.Zero,
      IntPtr.Zero,
      IntPtr.Zero);
    Assert.AreNotEqual(IntPtr.Zero, windowHandle);
    try
    {
      var workbookKey = $"{processId}|target";
      var firstToken = WorkbookSessionTokenRegistry.GetOrCreateWindowSessionToken(
        windowHandle,
        workbookKey,
        processId);
      var secondToken = WorkbookSessionTokenRegistry.GetOrCreateWindowSessionToken(
        windowHandle,
        workbookKey,
        processId);

      Assert.AreNotEqual(IntPtr.Zero, firstToken);
      Assert.AreEqual(firstToken, secondToken);
      Assert.IsTrue(WorkbookSessionTokenRegistry.IsValid(firstToken));
      Assert.AreEqual(
        firstToken,
        NativeMethods.GetProp(windowHandle, WorkbookSessionTokenRegistry.WindowPropertyName));
      WorkbookSessionTokenRegistry.Invalidate(firstToken);
    }
    finally
    {
      WorkbookSessionTokenRegistry.InvalidateProcess(processId);
      _ = NativeMethods.RemoveProp(windowHandle, WorkbookSessionTokenRegistry.WindowPropertyName);
      _ = NativeMethods.DestroyWindow(windowHandle);
    }
  }

  [TestMethod]
  public void InvalidateAll_RemovesWindowPropertiesAndRegistryState()
  {
    var processId = NextProcessId();
    var token = new IntPtr(Interlocked.Decrement(ref nextProcessId));
    var windowHandle = NativeMethods.CreateWindowEx(
      0,
      "STATIC",
      string.Empty,
      0,
      0,
      0,
      0,
      0,
      new IntPtr(-3),
      IntPtr.Zero,
      IntPtr.Zero,
      IntPtr.Zero);
    Assert.AreNotEqual(IntPtr.Zero, windowHandle);
    try
    {
      Assert.IsTrue(NativeMethods.SetProp(
        windowHandle,
        WorkbookSessionTokenRegistry.WindowPropertyName,
        token));
      WorkbookSessionTokenRegistry.Register($"{processId}|target", processId, token, windowHandle);
      WorkbookSessionTokenRegistry.RegisterMonitoredProcess(processId);

      WorkbookSessionTokenRegistry.InvalidateAll();

      Assert.AreEqual(
        IntPtr.Zero,
        NativeMethods.GetProp(windowHandle, WorkbookSessionTokenRegistry.WindowPropertyName));
      Assert.IsFalse(WorkbookSessionTokenRegistry.IsValid(token));
      Assert.IsFalse(WorkbookSessionTokenRegistry.IsProcessMonitored(processId));
      Assert.AreEqual(new WorkbookSessionTokenRegistry.RegistryState(0, 0, 0),
        WorkbookSessionTokenRegistry.GetState(processId));
    }
    finally
    {
      WorkbookSessionTokenRegistry.Invalidate(token);
      _ = NativeMethods.RemoveProp(windowHandle, WorkbookSessionTokenRegistry.WindowPropertyName);
      _ = NativeMethods.DestroyWindow(windowHandle);
    }
  }

  private static uint NextProcessId() => unchecked((uint)Interlocked.Decrement(ref nextProcessId));

  private static class NativeMethods
  {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetProp(nint windowHandle, string propertyName, nint value);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint GetProp(nint windowHandle, string propertyName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint RemoveProp(nint windowHandle, string propertyName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint CreateWindowEx(
      uint extendedStyle,
      string className,
      string windowName,
      uint style,
      int x,
      int y,
      int width,
      int height,
      nint parentWindow,
      nint menu,
      nint instance,
      nint parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool DestroyWindow(nint windowHandle);
  }
}
