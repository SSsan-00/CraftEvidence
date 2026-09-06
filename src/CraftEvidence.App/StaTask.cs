namespace CraftEvidence.App;

internal static class StaTask
{
  public static Task<T> Run<T>(Func<T> action)
  {
    ArgumentNullException.ThrowIfNull(action);
    var completion = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
    var thread = new Thread(() =>
    {
      try
      {
        completion.SetResult(action());
      }
      catch (Exception exception)
      {
        completion.SetException(exception);
      }
    })
    {
      IsBackground = true,
      Name = "CraftEvidence Excel STA worker",
    };
    thread.SetApartmentState(ApartmentState.STA);
    thread.Start();
    return completion.Task;
  }
}
