using System.Runtime.CompilerServices;

namespace EvidenceCrafter.App;

internal static class UiTheme
{
  private enum Role { Form, Canvas, Surface, MutedSurface, Text, MutedText, Border, Button, PrimaryButton, TextBox, ComboBox, SideButton }
  private sealed class RoleHolder(Role value) { internal Role Value { get; } = value; }
  private static readonly ConditionalWeakTable<Control, RoleHolder> roles = new();

  internal static bool DarkMode { get; private set; }
  internal static Color Canvas => DarkMode ? Color.FromArgb(13, 17, 23) : Color.FromArgb(244, 246, 248);
  internal static Color Surface => DarkMode ? Color.FromArgb(22, 27, 34) : Color.White;
  internal static Color SurfaceMuted => DarkMode ? Color.FromArgb(33, 38, 45) : Color.FromArgb(238, 242, 246);
  internal static Color Text => DarkMode ? Color.White : Color.FromArgb(23, 32, 42);
  internal static Color TextMuted => DarkMode ? Color.FromArgb(205, 217, 229) : Color.FromArgb(102, 112, 133);
  internal static Color Border => DarkMode ? Color.FromArgb(87, 96, 106) : Color.FromArgb(215, 222, 231);
  internal static Color Primary => DarkMode ? Color.FromArgb(88, 166, 255) : Color.FromArgb(23, 105, 170);
  internal static Color PrimaryHover => DarkMode ? Color.FromArgb(121, 192, 255) : Color.FromArgb(15, 86, 141);

  internal static void SetDarkMode(bool enabled) => DarkMode = enabled;

  internal static void StyleForm(Form form)
  {
    SetRole(form, Role.Form);
    Apply(form, Role.Form);
    form.Font = new Font("Meiryo UI", 9F);
  }

  internal static void StyleSurface(Control control, bool muted = false)
  {
    var role = muted ? Role.MutedSurface : Role.Surface;
    SetRole(control, role);
    Apply(control, role);
  }

  internal static void StyleCanvas(Control control)
  {
    SetRole(control, Role.Canvas);
    Apply(control, Role.Canvas);
  }

  internal static void StyleText(Control control, bool muted = false)
  {
    var role = muted ? Role.MutedText : Role.Text;
    SetRole(control, role);
    Apply(control, role);
  }

  internal static void StyleBorder(Control control)
  {
    SetRole(control, Role.Border);
    Apply(control, Role.Border);
  }

  internal static void StyleButton(Button button, Font font, bool primary = false)
  {
    var role = primary ? Role.PrimaryButton : Role.Button;
    SetRole(button, role);
    button.AutoSize = true;
    button.AutoSizeMode = AutoSizeMode.GrowAndShrink;
    button.Font = font;
    button.MinimumSize = new Size(64, 32);
    button.FlatStyle = FlatStyle.Standard;
    button.UseVisualStyleBackColor = false;
    button.Margin = new Padding(3, 0, 3, 0);
    button.Padding = new Padding(12, 2, 12, 2);
    button.TextAlign = ContentAlignment.MiddleCenter;
    button.UseCompatibleTextRendering = false;
    Apply(button, role);
  }

  internal static void StyleTextBox(TextBox textBox)
  {
    SetRole(textBox, Role.TextBox);
    textBox.AutoSize = true;
    textBox.BorderStyle = BorderStyle.FixedSingle;
    Apply(textBox, Role.TextBox);
  }

  internal static void StyleComboBox(ComboBox comboBox)
  {
    SetRole(comboBox, Role.ComboBox);
    Apply(comboBox, Role.ComboBox);
  }

  internal static void StyleSideButton(RadioButton button)
  {
    SetRole(button, Role.SideButton);
    button.Appearance = Appearance.Button;
    button.AutoSize = true;
    button.MinimumSize = new Size(64, 32);
    button.Padding = new Padding(12, 2, 12, 2);
    button.TextAlign = ContentAlignment.MiddleCenter;
    button.UseCompatibleTextRendering = false;
    button.FlatStyle = FlatStyle.Flat;
    button.Margin = new Padding(0, 0, 4, 0);
    Apply(button, Role.SideButton);
  }

  internal static void Refresh(Control root)
  {
    if (roles.TryGetValue(root, out var holder)) Apply(root, holder.Value);
    foreach (Control child in root.Controls) Refresh(child);
  }

  private static void SetRole(Control control, Role role)
  {
    roles.Remove(control);
    roles.Add(control, new RoleHolder(role));
  }

  private static void Apply(Control control, Role role)
  {
    switch (role)
    {
      case Role.Form: control.BackColor = Canvas; control.ForeColor = Text; break;
      case Role.Canvas: control.BackColor = Canvas; control.ForeColor = Text; break;
      case Role.Surface: control.BackColor = Surface; control.ForeColor = Text; break;
      case Role.MutedSurface: control.BackColor = SurfaceMuted; control.ForeColor = Text; break;
      case Role.Text: control.ForeColor = Text; break;
      case Role.MutedText: control.ForeColor = TextMuted; break;
      case Role.Border: control.BackColor = Border; break;
      case Role.Button: ApplyButton((Button)control, false); break;
      case Role.PrimaryButton: ApplyButton((Button)control, true); break;
      case Role.TextBox: control.BackColor = Surface; control.ForeColor = Text; break;
      case Role.ComboBox: control.BackColor = Surface; control.ForeColor = Text; break;
      case Role.SideButton:
        var button = (RadioButton)control;
        button.FlatAppearance.CheckedBackColor = Primary;
        button.FlatAppearance.MouseOverBackColor = SurfaceMuted;
        break;
    }
  }

  private static void ApplyButton(Button button, bool primary)
  {
    button.BackColor = primary ? Primary : SurfaceMuted;
    button.ForeColor = primary ? Color.White : Text;
  }
}

internal sealed class ThemedCheckBox : CheckBox
{
  protected override bool ShowFocusCues => false;
}
