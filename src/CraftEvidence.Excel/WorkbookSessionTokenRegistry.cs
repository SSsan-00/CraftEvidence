using System.Runtime.InteropServices;

namespace CraftEvidence.Excel;

internal static class WorkbookSessionTokenRegistry
{
  internal static readonly string WindowPropertyName =
    $"CraftEvidence.WorkbookSession.{Environment.ProcessId}";
  private static readonly Lock RegistryLock = new();
  private static readonly Dictionary<nint, TokenRegistration> ValidTokens = [];
  private static readonly HashSet<uint> MonitoredProcesses = [];
  private static readonly Dictionary<uint, HashSet<nint>> TokensByProcess = [];
  private static readonly Dictionary<string, TokenRegistration> TokensByWorkbook =
    new(StringComparer.OrdinalIgnoreCase);
  private static long nextWindowSessionToken;
  private static bool shutdownStarted;

  static WorkbookSessionTokenRegistry()
  {
    AppDomain.CurrentDomain.ProcessExit += (_, _) => BeginShutdown();
  }

  internal static bool IsValid(nint token)
  {
    lock (RegistryLock)
    {
      return token != IntPtr.Zero && ValidTokens.ContainsKey(token);
    }
  }

  internal static bool IsProcessMonitored(uint processId)
  {
    lock (RegistryLock)
    {
      return processId != 0 && MonitoredProcesses.Contains(processId);
    }
  }

  internal static void RegisterMonitoredProcess(uint processId)
  {
    if (processId != 0)
    {
      lock (RegistryLock)
      {
        _ = MonitoredProcesses.Add(processId);
      }
    }
  }

  internal static void UnregisterMonitoredProcess(uint processId)
  {
    lock (RegistryLock)
    {
      _ = MonitoredProcesses.Remove(processId);
      InvalidateProcessLocked(processId);
    }
  }

  internal static void MarkProcessUnmonitored(uint processId)
  {
    lock (RegistryLock)
    {
      _ = MonitoredProcesses.Remove(processId);
    }
  }

  internal static bool Register(
    string workbookKey,
    uint processId,
    nint token,
    nint documentWindowHandle = default)
  {
    if (processId == 0 || token == IntPtr.Zero)
    {
      return false;
    }

    lock (RegistryLock)
    {
      if (shutdownStarted)
      {
        RemoveWindowPropertyIfMatchesLocked(documentWindowHandle, token);
        return false;
      }

      if (documentWindowHandle != IntPtr.Zero &&
        NativeMethods.GetProp(documentWindowHandle, WindowPropertyName) != token)
      {
        return false;
      }

      RegisterLocked(workbookKey, processId, token, documentWindowHandle);
      return true;
    }
  }

  internal static nint GetOrCreateWindowSessionToken(
    nint documentWindowHandle,
    string workbookKey,
    uint processId)
  {
    if (documentWindowHandle == IntPtr.Zero || processId == 0)
    {
      return IntPtr.Zero;
    }

    lock (RegistryLock)
    {
      if (shutdownStarted)
      {
        RemoveWindowPropertyIfMatchesLocked(
          documentWindowHandle,
          NativeMethods.GetProp(documentWindowHandle, WindowPropertyName));
        return IntPtr.Zero;
      }

      var existing = NativeMethods.GetProp(documentWindowHandle, WindowPropertyName);
      if (existing != IntPtr.Zero && ValidTokens.ContainsKey(existing))
      {
        RegisterLocked(workbookKey, processId, existing, documentWindowHandle);
        return existing;
      }

      var tokenValue = Interlocked.Increment(ref nextWindowSessionToken);
      if (tokenValue <= 0)
      {
        return IntPtr.Zero;
      }

      var token = new IntPtr(tokenValue);
      if (!NativeMethods.SetProp(documentWindowHandle, WindowPropertyName, token))
      {
        return IntPtr.Zero;
      }

      RegisterLocked(workbookKey, processId, token, documentWindowHandle);
      return token;
    }
  }

  internal static void BeginShutdown()
  {
    lock (RegistryLock)
    {
      shutdownStarted = true;
      InvalidateAllLocked();
    }
  }

  internal static void Invalidate(nint token)
  {
    if (token != IntPtr.Zero)
    {
      lock (RegistryLock)
      {
        RemoveTokenLocked(token);
      }
    }
  }

  internal static void Invalidate(string workbookKey)
  {
    lock (RegistryLock)
    {
      if (TokensByWorkbook.TryGetValue(workbookKey, out var registration))
      {
        RemoveTokenLocked(registration.Token);
      }
    }
  }

  internal static void InvalidateProcess(uint processId)
  {
    if (processId == 0)
    {
      return;
    }

    lock (RegistryLock)
    {
      InvalidateProcessLocked(processId);
    }
  }

  internal static RegistryState GetState(uint processId)
  {
    lock (RegistryLock)
    {
      return new RegistryState(
        ValidTokens.Values.Count(item => item.ProcessId == processId),
        TokensByProcess.TryGetValue(processId, out var processTokens) ? processTokens.Count : 0,
        TokensByWorkbook.Values.Count(item => item.ProcessId == processId));
    }
  }

  internal static void InvalidateAll()
  {
    lock (RegistryLock)
    {
      InvalidateAllLocked();
    }
  }

  private static void RegisterLocked(
    string workbookKey,
    uint processId,
    nint token,
    nint documentWindowHandle)
  {
    if (TokensByWorkbook.TryGetValue(workbookKey, out var previousWorkbookRegistration) &&
      previousWorkbookRegistration.Token != token)
    {
      RemoveTokenLocked(previousWorkbookRegistration.Token);
    }

    if (ValidTokens.TryGetValue(token, out var previousTokenRegistration) &&
      (!string.Equals(previousTokenRegistration.WorkbookKey, workbookKey, StringComparison.OrdinalIgnoreCase) ||
        previousTokenRegistration.ProcessId != processId ||
        previousTokenRegistration.DocumentWindowHandle != documentWindowHandle))
    {
      var preserveWindowProperty =
        previousTokenRegistration.DocumentWindowHandle == documentWindowHandle;
      RemoveTokenLocked(token, removeWindowProperty: !preserveWindowProperty);
    }

    var registration = new TokenRegistration(workbookKey, processId, token, documentWindowHandle);
    ValidTokens[token] = registration;
    TokensByWorkbook[workbookKey] = registration;
    if (!TokensByProcess.TryGetValue(processId, out var processTokens))
    {
      processTokens = [];
      TokensByProcess[processId] = processTokens;
    }

    _ = processTokens.Add(token);
  }

  private static void InvalidateAllLocked()
  {
    foreach (var token in ValidTokens.Keys.ToArray())
    {
      RemoveTokenLocked(token);
    }

    ValidTokens.Clear();
    TokensByProcess.Clear();
    TokensByWorkbook.Clear();
    MonitoredProcesses.Clear();
  }

  private static void InvalidateProcessLocked(uint processId)
  {
    if (TokensByProcess.TryGetValue(processId, out var processTokens))
    {
      foreach (var token in processTokens.ToArray())
      {
        RemoveTokenLocked(token);
      }
    }

    _ = TokensByProcess.Remove(processId);
    foreach (var workbookKey in TokensByWorkbook
      .Where(item => item.Value.ProcessId == processId)
      .Select(item => item.Key)
      .ToArray())
    {
      _ = TokensByWorkbook.Remove(workbookKey);
    }
  }

  private static void RemoveTokenLocked(nint token, bool removeWindowProperty = true)
  {
    if (!ValidTokens.Remove(token, out var registration))
    {
      return;
    }

    if (removeWindowProperty)
    {
      RemoveWindowPropertyIfMatchesLocked(registration.DocumentWindowHandle, token);
    }

    if (TokensByWorkbook.TryGetValue(registration.WorkbookKey, out var workbookRegistration) &&
      workbookRegistration.Token == token)
    {
      _ = TokensByWorkbook.Remove(registration.WorkbookKey);
    }

    if (TokensByProcess.TryGetValue(registration.ProcessId, out var processTokens))
    {
      _ = processTokens.Remove(token);
      if (processTokens.Count == 0)
      {
        _ = TokensByProcess.Remove(registration.ProcessId);
      }
    }
  }

  private static void RemoveWindowPropertyIfMatchesLocked(nint windowHandle, nint token)
  {
    if (windowHandle != IntPtr.Zero && token != IntPtr.Zero &&
      NativeMethods.GetProp(windowHandle, WindowPropertyName) == token)
    {
      _ = NativeMethods.RemoveProp(windowHandle, WindowPropertyName);
    }
  }

  internal sealed record RegistryState(int ValidTokenCount, int ProcessTokenCount, int WorkbookCount);

  private sealed record TokenRegistration(
    string WorkbookKey,
    uint ProcessId,
    nint Token,
    nint DocumentWindowHandle);

  private static class NativeMethods
  {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetProp(nint windowHandle, string propertyName, nint data);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint GetProp(nint windowHandle, string propertyName);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern nint RemoveProp(nint windowHandle, string propertyName);
  }
}
