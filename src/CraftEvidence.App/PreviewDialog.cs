using CraftEvidence.Core.Models;

namespace CraftEvidence.App;

internal sealed class PreviewDialog : Form
{
  private readonly Image image;

  public PreviewDialog(
    Image image,
    string workbookLabel,
    string worksheetName,
    EvidenceSide side,
    bool canPlace)
  {
    this.image = image ?? throw new ArgumentNullException(nameof(image));
    Text = "スクリーンショットプレビュー — 完成仕様レビュー版";
    StartPosition = FormStartPosition.CenterParent;
    MinimumSize = new Size(700, 520);
    Size = new Size(900, 680);
    AutoScaleMode = AutoScaleMode.Dpi;
    Font = new Font("Meiryo UI", 9F);

    var layout = new TableLayoutPanel
    {
      Dock = DockStyle.Fill,
      Padding = new Padding(12),
      ColumnCount = 5,
      RowCount = 3,
    };
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var context = new Label
    {
      AutoSize = true,
      Text = $"Workbook: {workbookLabel}{Environment.NewLine}Sheet: {worksheetName}  /  Side: {side}",
    };
    layout.Controls.Add(context, 0, 0);
    layout.SetColumnSpan(context, 5);

    var picture = new PictureBox
    {
      Dock = DockStyle.Fill,
      Image = image,
      SizeMode = PictureBoxSizeMode.Zoom,
      BackColor = Color.FromArgb(32, 32, 32),
    };
    layout.Controls.Add(picture, 0, 1);
    layout.SetColumnSpan(picture, 5);

    var placeButton = new Button
    {
      Text = canPlace ? "そのまま配置" : "Workbook未選択",
      AutoSize = true,
      DialogResult = DialogResult.Yes,
      Enabled = canPlace,
    };
    layout.Controls.Add(placeButton, 0, 2);

    var editButton = new Button
    {
      Text = "編集して配置",
      AutoSize = true,
      DialogResult = DialogResult.Retry,
      Enabled = canPlace,
    };
    layout.Controls.Add(editButton, 1, 2);

    var automaticButton = new Button
    {
      Text = "自動配置",
      AutoSize = true,
      DialogResult = DialogResult.OK,
      Enabled = canPlace,
    };
    layout.Controls.Add(automaticButton, 2, 2);

    var tailButton = new Button
    {
      Text = "末尾へ配置",
      AutoSize = true,
      DialogResult = DialogResult.Ignore,
      Enabled = canPlace,
    };
    layout.Controls.Add(tailButton, 3, 2);

    var closeButton = new Button
    {
      Text = "閉じる",
      AutoSize = true,
      DialogResult = DialogResult.Cancel,
    };
    layout.Controls.Add(closeButton, 4, 2);

    AcceptButton = canPlace ? placeButton : closeButton;
    CancelButton = closeButton;
    Controls.Add(layout);
  }

  protected override void Dispose(bool disposing)
  {
    if (disposing)
    {
      image.Dispose();
    }

    base.Dispose(disposing);
  }
}
