using System.ComponentModel;
using System.Runtime.InteropServices;

namespace CraftEvidence.App;

[Flags]
internal enum ShortcutModifiers : uint
{
  None = 0,
  Alt = 0x0001,
  Control = 0x0002,
  Shift = 0x0004,
  Windows = 0x0008,
  NoRepeat = 0x4000,
}

internal sealed class GlobalShortcutRegistration : IDisposable
{
  internal const int HotKeyMessage = 0x0312;
  private readonly nint windowHandle;
  private readonly int id;
  private bool registered;

  private GlobalShortcutRegistration(nint windowHandle, int id)
  {
    this.windowHandle = windowHandle;
    this.id = id;
    registered = true;
  }

  internal static bool TryRegister(
    nint windowHandle,
    int id,
    ShortcutModifiers modifiers,
    Keys key,
    out GlobalShortcutRegistration? registration,
    out string? error)
  {
    registration = null;
    error = null;
    if (windowHandle == 0 || id is < 0 or > 0xBFFF || key == Keys.None)
    {
      error = "ショートカットの登録条件が不正です。";
      return false;
    }

    if (!NativeMethods.RegisterHotKey(windowHandle, id, (uint)modifiers, (uint)key))
    {
      error = new Win32Exception(Marshal.GetLastWin32Error()).Message;
      return false;
    }

    registration = new GlobalShortcutRegistration(windowHandle, id);
    return true;
  }

  internal bool Matches(ref Message message) =>
    registered && message.Msg == HotKeyMessage && message.WParam == id;

  public void Dispose()
  {
    if (!registered)
    {
      return;
    }

    _ = NativeMethods.UnregisterHotKey(windowHandle, id);
    registered = false;
  }

  private static class NativeMethods
  {
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool RegisterHotKey(nint windowHandle, int id, uint modifiers, uint virtualKey);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool UnregisterHotKey(nint windowHandle, int id);
  }
}
