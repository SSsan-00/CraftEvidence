using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.App;

/// <summary>Owns Excel event subscriptions on a background STA with a Windows message loop.</summary>
internal sealed class ExcelApplicationSessionMonitorHost : IDisposable
{
  private readonly TaskCompletionSource<WorkerState> ready =
    new(TaskCreationOptions.RunContinuationsAsynchronously);
  private readonly TaskCompletionSource stopped =
    new(TaskCreationOptions.RunContinuationsAsynchronously);
  private readonly Thread workerThread;
  private readonly Lock stateLock = new();
  private bool disposed;

  public event EventHandler<ExcelSelectionChangedEventArgs>? SelectionChanged;

  public ExcelApplicationSessionMonitorHost()
  {
    workerThread = new Thread(RunWorker)
    {
      IsBackground = true,
      Name = "EvidenceCrafter Excel session monitor",
    };
    workerThread.SetApartmentState(ApartmentState.STA);
    workerThread.Start();
  }

  public Task<IReadOnlyList<string>> RefreshAsync(IReadOnlyList<WorkbookIdentity> workbooks)
  {
    ArgumentNullException.ThrowIfNull(workbooks);
    var snapshot = workbooks.ToArray();
    return InvokeAsync(monitor => monitor.Refresh(snapshot));
  }

  public void InvalidateSessions(IReadOnlyList<WorkbookIdentity> workbooks) =>
    ExcelApplicationSessionMonitor.InvalidateSessions(workbooks);

  public void Dispose()
  {
    lock (stateLock)
    {
      if (disposed)
      {
        return;
      }

      disposed = true;
    }

    ExcelApplicationSessionMonitor.BeginShutdownSessions();
    _ = ShutdownAsync();
  }

  private async Task<T> InvokeAsync<T>(Func<ExcelApplicationSessionMonitor, T> action)
  {
    ArgumentNullException.ThrowIfNull(action);
    lock (stateLock)
    {
      ObjectDisposedException.ThrowIf(disposed, this);
    }

    var worker = await ready.Task.ConfigureAwait(false);
    var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
    worker.Context.Post(_ =>
    {
      try
      {
        lock (stateLock)
        {
          ObjectDisposedException.ThrowIf(disposed, this);
        }

        completion.TrySetResult(action(worker.Monitor));
      }
      catch (Exception exception)
      {
        completion.TrySetException(exception);
      }
    }, null);
    return await completion.Task.ConfigureAwait(false);
  }

  private async Task ShutdownAsync()
  {
    try
    {
      var worker = await ready.Task.ConfigureAwait(false);
      worker.Context.Post(_ =>
      {
        try
        {
          worker.Monitor.Dispose();
        }
        finally
        {
          System.Windows.Forms.Application.ExitThread();
        }
      }, null);
    }
    catch (Exception)
    {
      // A failed worker has no live COM subscription to release.
    }

    _ = stopped.Task;
  }

  private void RunWorker()
  {
    ExcelApplicationSessionMonitor? monitor = null;
    WindowsFormsSynchronizationContext? context = null;
    try
    {
      context = new WindowsFormsSynchronizationContext();
      SynchronizationContext.SetSynchronizationContext(context);
      monitor = new ExcelApplicationSessionMonitor();
      monitor.SelectionChanged += (_, eventArgs) => SelectionChanged?.Invoke(this, eventArgs);
      ready.TrySetResult(new WorkerState(context, monitor));
      System.Windows.Forms.Application.Run();
    }
    catch (Exception exception)
    {
      ready.TrySetException(exception);
    }
    finally
    {
      try
      {
        monitor?.Dispose();
      }
      finally
      {
        context?.Dispose();
        stopped.TrySetResult();
      }
    }
  }

  private sealed record WorkerState(
    WindowsFormsSynchronizationContext Context,
    ExcelApplicationSessionMonitor Monitor);
}
