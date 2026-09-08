using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.Tests;

[TestClass]
public sealed class ExcelSessionCatalogTests
{
  [TestMethod]
  [DoNotParallelize]
  public void Discover_UnsupportedAndThrowingRegistrations_DoNotStopOtherCandidates()
  {
    Exception? failure = null;
    var thread = new Thread(() =>
    {
      IRunningObjectTable? table = null;
      IStream? stream = null;
      var registrations = new List<int>();
      try
      {
        Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out table));
        Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out stream));
        foreach (var candidate in new object[] { stream, new DiscoveryProbe { ThrowOnRead = true } })
        {
          Marshal.ThrowExceptionForHR(CreateItemMoniker("!", Guid.NewGuid().ToString("N"), out var moniker));
          try
          {
            registrations.Add(table.Register(1, candidate, moniker));
          }
          finally
          {
            _ = Marshal.ReleaseComObject(moniker);
          }
        }

        var result = new ExcelSessionCatalog().Discover();
        Assert.IsTrue(result.Warnings.Any(item => item.Contains("IDispatch 非対応")));
      }
      catch (Exception exception)
      {
        failure = exception;
      }
      finally
      {
        foreach (var cookie in registrations)
        {
          table!.Revoke(cookie);
        }

        if (stream is not null) { _ = Marshal.ReleaseComObject(stream); }
        if (table is not null) { _ = Marshal.ReleaseComObject(table); }
      }
    }) { IsBackground = true };
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    Assert.IsTrue(thread.Join(TimeSpan.FromSeconds(20)), "Discovery timed out.");
    Assert.IsNull(failure, failure?.ToString());
  }

  [TestMethod]
  public void SupportsDispatch_NativeStreamWithoutAutomation_IsSkippedWithoutLeakingReference()
  {
    Marshal.ThrowExceptionForHR(CreateStreamOnHGlobal(IntPtr.Zero, true, out var stream));
    try
    {
      Assert.IsTrue(Marshal.IsComObject(stream));
      Assert.IsFalse(ExcelSessionCatalog.SupportsDispatch(stream));
      Assert.IsFalse(ExcelSessionCatalog.SupportsDispatch(stream));
      // The guard must not release the caller's reference while checking the interface.
      stream.Stat(out var status, 1);
      Assert.AreEqual(0L, status.cbSize);
    }
    finally
    {
      _ = Marshal.ReleaseComObject(stream);
    }
  }

  [TestMethod]
  public void DiscoveryFailure_OnlyAccessDeniedSuggestsCheckingPermissions()
  {
    var denied = new TargetInvocationException(new COMException("denied", unchecked((int)0x80070005)));
    StringAssert.Contains(ExcelSessionCatalog.DescribeDiscoveryFailure(denied), "権限レベル");
    StringAssert.Contains(ExcelSessionCatalog.DescribeDiscoveryFailure(denied), "0x80070005");
    Assert.DoesNotContain("権限レベル", ExcelSessionCatalog.DescribeDiscoveryFailure(
      new NotSupportedException("COM target does not implement IDispatch.")));
  }

  [DllImport("ole32.dll")]
  private static extern int CreateStreamOnHGlobal(
    nint memory,
    [MarshalAs(UnmanagedType.Bool)] bool deleteOnRelease,
    out IStream stream);

  [DllImport("ole32.dll")]
  private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable table);

  [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
  private static extern int CreateItemMoniker(string delimiter, string item, out IMoniker moniker);
}

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class DiscoveryProbe
{
  public bool ThrowOnRead { get; set; }
  public int ReadCount { get; private set; }
  public object? Workbooks
  {
    get
    {
      ReadCount++;
      return ThrowOnRead ? throw new NotSupportedException("COM target does not implement IDispatch.") : null;
    }
  }
}
