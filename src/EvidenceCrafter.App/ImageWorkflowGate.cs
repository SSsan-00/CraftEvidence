namespace EvidenceCrafter.App;

// Owned by the UI thread; the lease covers analysis, modal dialogs and placement.
internal sealed class ImageWorkflowGate
{
  internal bool IsActive { get; private set; }
  internal bool TryBegin()
  {
    if (IsActive) return false;
    IsActive = true;
    return true;
  }
  internal void End() => IsActive = false;
}
