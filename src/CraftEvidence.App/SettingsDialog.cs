namespace CraftEvidence.App;

internal sealed class SettingsDialog : Form
{
  private readonly NumericUpDown margin = new();
  private readonly CheckBox logging = new();
  private readonly CheckBox shortcut = new();

  internal SettingsDialog(CraftEvidenceSettings current)
  {
    ArgumentNullException.ThrowIfNull(current);
    Text = "CraftEvidence 設定";
    StartPosition = FormStartPosition.CenterParent;
    FormBorderStyle = FormBorderStyle.FixedDialog;
    MaximizeBox = false;
    MinimizeBox = false;
    AutoSize = true;
    AutoSizeMode = AutoSizeMode.GrowAndShrink;
    Padding = new Padding(12);
    Font = new Font("Meiryo UI", 9F);

    var layout = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, RowCount = 4 };
    layout.Controls.Add(new Label { AutoSize = true, Text = "左右余白 (pt)" }, 0, 0);
    margin.DecimalPlaces = 1;
    margin.Minimum = 0;
    margin.Maximum = 72;
    margin.Value = (decimal)current.HorizontalMarginPoints;
    layout.Controls.Add(margin, 1, 0);
    logging.AutoSize = true;
    logging.Text = "診断ログを有効にする（画像・セル内容・パスは記録しません）";
    logging.Checked = current.DiagnosticLoggingEnabled;
    layout.Controls.Add(logging, 0, 1);
    layout.SetColumnSpan(logging, 2);
    shortcut.AutoSize = true;
    shortcut.Text = "Ctrl+Shift+E ショートカットを有効にする（次回起動時）";
    shortcut.Checked = current.GlobalShortcutEnabled;
    layout.Controls.Add(shortcut, 0, 2);
    layout.SetColumnSpan(shortcut, 2);

    var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
    var cancel = new Button { AutoSize = true, Text = "キャンセル", DialogResult = DialogResult.Cancel };
    var save = new Button { AutoSize = true, Text = "保存", DialogResult = DialogResult.OK };
    buttons.Controls.Add(cancel);
    buttons.Controls.Add(save);
    layout.Controls.Add(buttons, 0, 3);
    layout.SetColumnSpan(buttons, 2);
    Controls.Add(layout);
    AcceptButton = save;
    CancelButton = cancel;
  }

  internal CraftEvidenceSettings Result => new()
  {
    HorizontalMarginPoints = (double)margin.Value,
    DiagnosticLoggingEnabled = logging.Checked,
    GlobalShortcutEnabled = shortcut.Checked,
  };
}
