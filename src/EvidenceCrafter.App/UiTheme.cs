namespace EvidenceCrafter.App;

internal static class UiTheme
{
  internal static readonly Color Canvas = Color.FromArgb(244, 246, 248);
  internal static readonly Color Surface = Color.White;
  internal static readonly Color SurfaceMuted = Color.FromArgb(238, 242, 246);
  internal static readonly Color Text = Color.FromArgb(23, 32, 42);
  internal static readonly Color TextMuted = Color.FromArgb(102, 112, 133);
  internal static readonly Color Border = Color.FromArgb(215, 222, 231);
  internal static readonly Color Primary = Color.FromArgb(23, 105, 170);
  internal static readonly Color PrimaryHover = Color.FromArgb(15, 86, 141);

  internal static void StyleForm(Form form)
  {
    form.BackColor = Canvas;
    form.Font = new Font("Meiryo UI", 9F);
  }

  internal static void StyleButton(Button button, Font font, bool primary = false)
  {
    button.AutoSize = false;
    button.Font = font;
    button.Size = new Size(Math.Max(64, TextRenderer.MeasureText(button.Text, font).Width + 24), 30);
    button.FlatStyle = FlatStyle.Flat;
    button.FlatAppearance.BorderSize = 1;
    button.FlatAppearance.BorderColor = primary ? Primary : Border;
    button.FlatAppearance.MouseOverBackColor = primary ? PrimaryHover : SurfaceMuted;
    button.FlatAppearance.MouseDownBackColor = primary ? PrimaryHover : Color.FromArgb(226, 232, 239);
    button.BackColor = primary ? Primary : Surface;
    button.ForeColor = primary ? Color.White : Text;
    button.Margin = new Padding(3, 0, 3, 0);
    button.Padding = new Padding(6, 0, 6, 0);
    button.UseCompatibleTextRendering = false;
  }

  internal static void StyleTextBox(TextBox textBox)
  {
    textBox.AutoSize = false;
    textBox.Height = 30;
    textBox.BackColor = Surface;
    textBox.ForeColor = Text;
    textBox.BorderStyle = BorderStyle.FixedSingle;
  }

  internal static void StyleComboBox(ComboBox comboBox)
  {
    comboBox.BackColor = Surface;
    comboBox.ForeColor = Text;
  }
}
