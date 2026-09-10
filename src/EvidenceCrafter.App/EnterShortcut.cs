using System.Runtime.InteropServices;

namespace EvidenceCrafter.App;

// Form-local shortcuts must not consume IME composition or repeat into a new dialog.
internal static class EnterShortcut
{
  internal static bool IsRepeat(Message message) => (message.LParam.ToInt64() & (1L << 30)) != 0;

  internal static bool IsEditingInput(Form form)
  {
    Control? control = form.ActiveControl;
    while (control is ContainerControl container && container.ActiveControl is not null)
      control = container.ActiveControl;
    if (control is ComboBox { DroppedDown: true }) return true;
    if (control is null || !control.IsHandleCreated) return false;
    var context = ImmGetContext(control.Handle);
    if (context == 0) return false;
    try { return ImmGetCompositionString(context, 8, 0, 0) > 0; }
    finally { ImmReleaseContext(control.Handle, context); }
  }

  [DllImport("imm32.dll")]
  private static extern nint ImmGetContext(nint window);
  [DllImport("imm32.dll", CharSet = CharSet.Unicode)]
  private static extern int ImmGetCompositionString(nint context, int index, nint buffer, int length);
  [DllImport("imm32.dll")]
  private static extern bool ImmReleaseContext(nint window, nint context);
}
