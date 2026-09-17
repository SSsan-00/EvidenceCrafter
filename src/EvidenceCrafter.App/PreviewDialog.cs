using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.App;

internal sealed class PreviewDialog : Form
{
  private readonly Image image;
  private readonly Button placeButton;
  private readonly Button editButton;
  private readonly Rectangle? captureWindowBounds;

  public PreviewDialog(
    Image image,
    string workbookLabel,
    string worksheetName,
    EvidenceSide side,
    AutomaticPlacementAnalysisResult? analysis,
    Rectangle? captureWindowBounds = null)
  {
    this.image = image ?? throw new ArgumentNullException(nameof(image));
    this.captureWindowBounds = captureWindowBounds;
    Text = "スクリーンショットプレビュー";
    StartPosition = captureWindowBounds is null ? FormStartPosition.CenterParent : FormStartPosition.Manual;
    MinimumSize = new Size(700, 520);
    Size = new Size(900, 680);
    AutoScaleMode = AutoScaleMode.Dpi;
    Font = new Font("Meiryo UI", 9F);
    UiTheme.StyleForm(this);

    var layout = new TableLayoutPanel
    {
      Dock = DockStyle.Fill,
      Padding = new Padding(12),
      BackColor = UiTheme.Canvas,
      ColumnCount = 1,
      RowCount = 3,
    };
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
    layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

    var context = new Label
    {
      AutoSize = true,
      Text = BuildContext(workbookLabel, worksheetName, side, analysis),
      ForeColor = UiTheme.TextOn(UiTheme.SurfaceMuted),
      BackColor = UiTheme.SurfaceMuted,
      Padding = new Padding(8),
    };
    layout.Controls.Add(context, 0, 0);

    var picture = new PictureBox
    {
      Dock = DockStyle.Fill,
      Image = image,
      SizeMode = PictureBoxSizeMode.Zoom,
      BackColor = Color.FromArgb(32, 32, 32),
    };
    layout.Controls.Add(picture, 0, 1);

    var canPlace = analysis?.Succeeded == true;

    placeButton = new Button
    {
      Text = canPlace ? "編集せず配置" : "配置不可",
      AutoSize = true,
      DialogResult = DialogResult.Yes,
      Enabled = canPlace,
    };

    editButton = new Button
    {
      Text = "編集して配置",
      AutoSize = true,
      DialogResult = DialogResult.Retry,
      Enabled = canPlace,
    };

    var closeButton = new Button
    {
      Text = "閉じる",
      AutoSize = true,
      DialogResult = DialogResult.Cancel,
    };
    var buttons = new FlowLayoutPanel
    {
      AutoSize = true,
      Dock = DockStyle.Fill,
      FlowDirection = FlowDirection.RightToLeft,
      WrapContents = false,
    };
    foreach (var button in new[] { placeButton, editButton, closeButton })
    {
      UiTheme.StyleButton(button, Font, ReferenceEquals(button, placeButton) && canPlace);
      button.MinimumSize = new Size(130, 32);
      buttons.Controls.Add(button);
    }
    layout.Controls.Add(buttons, 0, 2);

    AcceptButton = canPlace ? placeButton : closeButton;
    CancelButton = closeButton;
    Controls.Add(layout);
  }

  protected override void OnShown(EventArgs eventArgs)
  {
    base.OnShown(eventArgs);
    if (captureWindowBounds is { } bounds) CaptureWindowPlacement.CenterInWindow(this, bounds);
  }

  protected override bool ProcessCmdKey(ref Message message, Keys keyData)
  {
    if (keyData == Keys.Enter || keyData == (Keys.Shift | Keys.Enter))
    {
      if (!EnterShortcut.IsRepeat(message))
      {
        var button = keyData == Keys.Enter ? placeButton : editButton;
        if (button.Enabled) button.PerformClick();
      }
      return true;
    }
    return base.ProcessCmdKey(ref message, keyData);
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
    return $"配置予定  |  Sheet: {analysis.WorksheetName}  |  CASE: {analysis.CaseLabel}  |  Side: {analysis.ResolvedSide}{Environment.NewLine}" +
      $"構成: {(analysis.LayoutAnalysis!.Layout!.Kind == SideLayoutKind.NewOnly ? "Newのみ" : "New/Old")}  |  " +
      $"開始セル: {ColumnName(step.Plan.FocusCell.Column)}{step.Plan.FocusCell.Row}  |  " +
      $"配置方法: {ModeLabel(step.Plan.Mode)}  |  追加予定行: {insertionCount}行  |  " +
      $"画像幅: {step.Plan.Image.WidthPoints:0.#}pt / 配置可能幅: {step.AvailableWidthPoints:0.#}pt";
  }

  private static string ModeLabel(PlacementMode mode) => mode switch
  {
    PlacementMode.CaseStart => "CASE先頭",
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
