using CraftEvidence.Core.Models;
using CraftEvidence.Excel;

namespace CraftEvidence.App;

internal sealed class PreviewDialog : Form
{
  private readonly Image image;

  public PreviewDialog(
    Image image,
    string workbookLabel,
    string worksheetName,
    EvidenceSide side,
    AutomaticPlacementAnalysisResult? analysis)
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
      ColumnCount = 3,
      RowCount = 3,
    };
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var context = new Label
    {
      AutoSize = true,
      Text = BuildContext(workbookLabel, worksheetName, side, analysis),
    };
    layout.Controls.Add(context, 0, 0);
    layout.SetColumnSpan(context, 3);

    var picture = new PictureBox
    {
      Dock = DockStyle.Fill,
      Image = image,
      SizeMode = PictureBoxSizeMode.Zoom,
      BackColor = Color.FromArgb(32, 32, 32),
    };
    layout.Controls.Add(picture, 0, 1);
    layout.SetColumnSpan(picture, 3);

    var canPlace = analysis?.Succeeded == true;

    var placeButton = new Button
    {
      Text = canPlace ? "編集せず配置" : "配置不可",
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

    var closeButton = new Button
    {
      Text = "閉じる",
      AutoSize = true,
      DialogResult = DialogResult.Cancel,
    };
    layout.Controls.Add(closeButton, 2, 2);

    AcceptButton = canPlace ? placeButton : closeButton;
    CancelButton = closeButton;
    Controls.Add(layout);
  }

  private static string BuildContext(
    string workbookLabel,
    string worksheetName,
    EvidenceSide side,
    AutomaticPlacementAnalysisResult? analysis)
  {
    if (analysis?.Succeeded != true || analysis.Steps.Count == 0)
    {
      return $"Workbook: {workbookLabel}{Environment.NewLine}Sheet: {worksheetName}{Environment.NewLine}" +
        $"配置先を解析できません: {analysis?.Message ?? "Workbookが選択されていません。"}";
    }

    var step = analysis.Steps[0];
    var insertionCount = step.Plan.Insertions.Sum(insertion => insertion.Count);
    return $"配置予定  |  Sheet: {analysis.WorksheetName}  |  Case: {analysis.CaseLabel}  |  Side: {side}{Environment.NewLine}" +
      $"開始セル: {ColumnName(step.Plan.FocusCell.Column)}{step.Plan.FocusCell.Row}  |  " +
      $"配置方法: {ModeLabel(step.Plan.Mode)}  |  追加予定行: {insertionCount}行  |  " +
      $"画像幅: {step.Plan.Image.WidthPoints:0.#}pt / 配置可能幅: {step.AvailableWidthPoints:0.#}pt";
  }

  private static string ModeLabel(PlacementMode mode) => mode switch
  {
    PlacementMode.CaseStart => "Case先頭",
    PlacementMode.Gap => "選択中の空き領域",
    _ => "既存画像の末尾",
  };

  private static string ColumnName(int column)
  {
    var name = string.Empty;
    while (column > 0)
    {
      column--;
      name = (char)('A' + column % 26) + name;
      column /= 26;
    }
    return name;
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
