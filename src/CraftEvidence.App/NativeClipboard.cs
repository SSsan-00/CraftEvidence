using System.Runtime.InteropServices;

namespace CraftEvidence.App;

internal static class NativeClipboard
{
  internal const int ClipboardUpdateMessage = 0x031D;

  [DllImport("user32.dll")]
  internal static extern uint GetClipboardSequenceNumber();

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  internal static extern bool AddClipboardFormatListener(nint windowHandle);

  [DllImport("user32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  internal static extern bool RemoveClipboardFormatListener(nint windowHandle);
}
