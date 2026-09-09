using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using EvidenceCrafter.Core.Models;
using EvidenceCrafter.Core.Services;
using EvidenceCrafter.Excel;

namespace EvidenceCrafter.App;

public sealed class MainForm : Form
{
  private static readonly TimeSpan ExcelDiscoveryTimeout = TimeSpan.FromSeconds(15);
  private static readonly TimeSpan ExcelMonitorTimeout = TimeSpan.FromSeconds(15);
  private const int ClipboardRetryLimit = 4;
  private readonly IExcelSessionCatalog sessionCatalog;
  private readonly ExcelApplicationSessionMonitorHost sessionMonitor = new();
  private readonly ExcelImagePlacementService imagePlacementService = new();
  private readonly ExcelRowMutationService rowMutationService = new();
  private readonly ExcelAutomaticPlacementService automaticPlacementService = new();
  private readonly ExcelManagedShapeService managedShapeService = new();
  private readonly ExcelCaseNavigationService caseNavigationService = new();
  private readonly CaseLayoutAnalyzer placementContextLayoutAnalyzer = new();
  private readonly ExcelCaseMaintenanceService caseMaintenanceService;
  private readonly ExcelManagedReplacementLayoutService replacementLayoutService;
  private readonly AppSettingsStore settingsStore = new();
  private readonly DiagnosticLog diagnosticLog = new();
  private readonly ComboBox workbookSelector = new();
  private readonly TextBox worksheetNameBox = new();
  private readonly TextBox caseLabelBox = new();
  private readonly NumericUpDown insertRowCountBox = new();
  private readonly TextBox deleteCaseStartBox = new();
  private readonly TextBox deleteCaseEndBox = new();
  private readonly Button insertRowsButton = new();
  private readonly Button deleteRowsButton = new();
  private readonly Button undoButton = new();
  private readonly Button redoButton = new();
  private readonly Button replaceImageButton = new();
  private readonly Button deleteImageButton = new();
  private readonly RadioButton newSideButton = new();
  private readonly RadioButton oldSideButton = new();
  private readonly ToolTip sideToolTip = new();
  private readonly Button previousCaseButton = new();
  private readonly Button nextCaseButton = new();
  private readonly Label statusLabel = new();
  private readonly Button refreshButton = new();
  private readonly Button captureScreenButton = new();
  private readonly ComboBox advanceModeBox = new();
  private readonly System.Windows.Forms.Timer clipboardRetryTimer = new();
  private readonly System.Windows.Forms.Timer selectionChangeTimer = new() { Interval = 250 };
  private bool clipboardListenerRegistered;
  private string? clipboardWarning;
  private uint? lastClipboardSequenceNumber;
  private Task<ExcelDiscoveryResult>? pendingDiscoveryTask;
  private Task<IReadOnlyList<string>>? pendingMonitorTask;
  private WorkbookIdentity[] pendingMonitorWorkbooks = [];
  private string? pendingMonitorKey;
  private bool clipboardPreviewOpen;
  private bool clipboardPreviewPending;
  private uint clipboardRetrySequence;
  private int clipboardRetryAttempt;
  private readonly Stack<HistoryEntry> undoHistory = new();
  private readonly Stack<HistoryEntry> redoHistory = new();
  private EvidenceCrafterSettings settings = new();
  private GlobalShortcutRegistration? globalShortcut;
  private int mutationInProgress;
  private bool screenCaptureInProgress;
  private int placementContextRequestVersion;
  private bool updatingPlacementContext;
  private bool placementTargetOverridden;
  private bool placementContextRefreshInProgress;
  private bool placementContextRefreshPending;
  private bool placementContextRefreshPendingForce;
  private ExcelSelectionChangedEventArgs? latestSelectionChange;
  private AutomaticPlacementAnalysisResult? cachedPlacementContext;

  private bool CanUpdateUi =>
    IsHandleCreated &&
    !IsDisposed &&
    !Disposing &&
    !statusLabel.IsDisposed;

  public MainForm(IExcelSessionCatalog sessionCatalog)
  {
    this.sessionCatalog = sessionCatalog ?? throw new ArgumentNullException(nameof(sessionCatalog));
    caseMaintenanceService = new ExcelCaseMaintenanceService(rowMutationService);
    replacementLayoutService = new ExcelManagedReplacementLayoutService(rowMutationService);
    settings = settingsStore.Load();
    clipboardRetryTimer.Tick += (_, _) =>
    {
      clipboardRetryTimer.Stop();
      QueueClipboardPreview();
    };
    selectionChangeTimer.Tick += async (_, _) =>
    {
      selectionChangeTimer.Stop();
      await RefreshPlacementContextAsync();
    };
    sessionMonitor.SelectionChanged += SessionMonitorSelectionChanged;
    InitializeUi();
  }

  private EvidenceSide SelectedSide => oldSideButton.Checked ? EvidenceSide.Old : EvidenceSide.New;

  protected override void OnHandleCreated(EventArgs e)
  {
    base.OnHandleCreated(e);
    clipboardListenerRegistered = NativeClipboard.AddClipboardFormatListener(Handle);
    if (settings.GlobalShortcutEnabled &&
      !GlobalShortcutRegistration.TryRegister(
        Handle,
        id: 1,
        ShortcutModifiers.Control | ShortcutModifiers.Shift | ShortcutModifiers.NoRepeat,
        Keys.E,
        out globalShortcut,
        out var shortcutError))
    {
      clipboardWarning = $"ショートカット Ctrl+Shift+E を登録できません: {shortcutError}";
    }
    if (!clipboardListenerRegistered)
    {
      clipboardWarning = "Clipboard監視を開始できませんでした。";
      statusLabel.Text = clipboardWarning;
    }
  }

  protected override void OnHandleDestroyed(EventArgs e)
  {
    if (clipboardListenerRegistered)
    {
      _ = NativeClipboard.RemoveClipboardFormatListener(Handle);
      clipboardListenerRegistered = false;
    }

    globalShortcut?.Dispose();
    globalShortcut = null;

    base.OnHandleDestroyed(e);
  }

  protected override void Dispose(bool disposing)
  {
    if (disposing)
    {
      ClearHistoryStack(undoHistory);
      ClearHistoryStack(redoHistory);
      sessionMonitor.Dispose();
      clipboardRetryTimer.Dispose();
      selectionChangeTimer.Dispose();
      try
      {
        settingsStore.Save(settings);
      }
      catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
      {
        WriteDiagnostic(DiagnosticEventKind.Application, DiagnosticOutcome.Failed, exception: exception);
      }
    }

    base.Dispose(disposing);
  }

  protected override void WndProc(ref Message message)
  {
    base.WndProc(ref message);
    if (globalShortcut?.Matches(ref message) is true)
    {
      BeginInvoke(async () => await CaptureScreenAsync());
      return;
    }
    if (message.Msg == NativeClipboard.ClipboardUpdateMessage &&
      IsHandleCreated &&
      !IsDisposed &&
      !Disposing)
    {
      try
      {
        QueueClipboardPreview();
      }
      catch (InvalidOperationException) when (IsDisposed || Disposing)
      {
        // The form was closed between receiving the native message and dispatching the callback.
      }
    }
  }

  private void InitializeUi()
  {
    Text = "EvidenceCrafter";
    StartPosition = FormStartPosition.CenterScreen;
    MinimumSize = new Size(760, 370);
    Size = new Size(760, 370);
    AutoScaleMode = AutoScaleMode.Dpi;
    Font = new Font("Meiryo UI", 9F);
    UiTheme.StyleForm(this);

    var layout = new TableLayoutPanel
    {
      Dock = DockStyle.Fill,
      Padding = new Padding(14, 12, 14, 12),
      ColumnCount = 1,
      RowCount = 5,
      AutoScroll = false,
    };
    layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    for (var row = 0; row < 5; row++)
    {
      layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
    }

    layout.Controls.Add(new Label
    {
      AutoSize = true,
      Text = "EvidenceCrafter",
      Font = new Font("Meiryo UI", 15F, FontStyle.Bold),
      ForeColor = UiTheme.Text,
      Anchor = AnchorStyles.Left,
      Margin = new Padding(7, 2, 0, 8),
    }, 0, 0);

    var workbookRow = CreateCard(3);
    workbookRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    workbookRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    workbookRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    workbookRow.Controls.Add(CreateLabel("対象ブック"), 0, 0);
    workbookSelector.Dock = DockStyle.Fill;
    workbookSelector.DropDownStyle = ComboBoxStyle.DropDownList;
    UiTheme.StyleComboBox(workbookSelector);
    workbookSelector.ItemHeight = 21;
    workbookSelector.Margin = new Padding(0);
    workbookSelector.DisplayMember = nameof(WorkbookIdentity.DisplayLabel);
    workbookSelector.SelectedIndexChanged += async (_, _) => await RefreshPlacementContextAsync(force: true);
    workbookRow.Controls.Add(workbookSelector, 1, 0);
    refreshButton.Text = "更新";
    StyleButton(refreshButton);
    refreshButton.Margin = new Padding(6, 0, 0, 0);
    refreshButton.Click += async (_, _) => await RefreshWorkbooksAsync();
    workbookRow.Controls.Add(refreshButton, 2, 0);
    layout.Controls.Add(workbookRow, 0, 1);

    var targetCard = CreateCard(6);
    targetCard.Margin = new Padding(0, 8, 0, 8);
    targetCard.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    targetCard.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
    targetCard.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    targetCard.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    targetCard.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    targetCard.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
    var sheetLabel = CreateLabel("Sheet");
    sheetLabel.Margin = new Padding(3, 0, 6, 0);
    targetCard.Controls.Add(sheetLabel, 0, 0);
    worksheetNameBox.Width = 120;
    worksheetNameBox.Anchor = AnchorStyles.Left;
    StyleTextBox(worksheetNameBox);
    worksheetNameBox.PlaceholderText = "シート名";
    worksheetNameBox.TextChanged += (_, _) => MarkPlacementTargetOverridden();
    worksheetNameBox.Validated += async (_, _) => await RefreshManualSideLayoutAsync();
    targetCard.Controls.Add(worksheetNameBox, 1, 0);

    var placementLabel = CreateLabel("配置先");
    placementLabel.Margin = new Padding(12, 0, 6, 0);
    targetCard.Controls.Add(placementLabel, 2, 0);
    var sidePanel = new FlowLayoutPanel
    {
      AutoSize = true,
      WrapContents = false,
      Margin = new Padding(0),
      Padding = new Padding(0),
    };
    newSideButton.Text = "NEW";
    StyleSideButton(newSideButton);
    newSideButton.Checked = true;
    oldSideButton.Text = "OLD";
    StyleSideButton(oldSideButton);
    newSideButton.CheckedChanged += (_, _) =>
    {
      UpdateSideButtonColors();
      MarkPlacementTargetOverridden();
    };
    oldSideButton.CheckedChanged += (_, _) =>
    {
      UpdateSideButtonColors();
      MarkPlacementTargetOverridden();
    };
    sidePanel.Controls.Add(newSideButton);
    sidePanel.Controls.Add(oldSideButton);
    targetCard.Controls.Add(sidePanel, 3, 0);

    var caseLabel = CreateLabel("CASE");
    caseLabel.Margin = new Padding(12, 0, 6, 0);
    targetCard.Controls.Add(caseLabel, 4, 0);
    caseLabelBox.Width = 86;
    StyleTextBox(caseLabelBox);
    caseLabelBox.PlaceholderText = "自動";
    caseLabelBox.Validated += async (_, _) => await RefreshManualSideLayoutAsync();
    caseLabelBox.TextChanged += (_, _) =>
    {
      if (!updatingPlacementContext)
      {
        placementTargetOverridden = true;
      }
    };
    targetCard.Controls.Add(caseLabelBox, 5, 0);
    previousCaseButton.Text = "前のCASE";
    StyleButton(previousCaseButton);
    previousCaseButton.Margin = new Padding(8, 0, 3, 0);
    previousCaseButton.Click += async (_, _) => await NavigateCaseAsync(CaseNavigationDirection.Previous);
    nextCaseButton.Text = "次のCASE";
    StyleButton(nextCaseButton);
    nextCaseButton.Click += async (_, _) => await NavigateCaseAsync(CaseNavigationDirection.Next);

    var advanceLabel = CreateLabel("配置後");
    advanceLabel.Margin = new Padding(3, 0, 6, 0);
    targetCard.Controls.Add(advanceLabel, 0, 1);
    advanceModeBox.Dock = DockStyle.Fill;
    advanceModeBox.DropDownStyle = ComboBoxStyle.DropDownList;
    UiTheme.StyleComboBox(advanceModeBox);
    advanceModeBox.ItemHeight = 21;
    advanceModeBox.Margin = new Padding(0);
    advanceModeBox.Items.AddRange(["同じCASEの反対Side", "同じSideの次CASE"]);
    advanceModeBox.SelectedIndex = settings.AdvanceMode is PlacementAdvanceMode.NextCaseSameSide ? 1 : 0;
    advanceModeBox.SelectedIndexChanged += (_, _) => SaveAdvanceMode();
    targetCard.Controls.Add(advanceModeBox, 1, 1);
    targetCard.SetColumnSpan(advanceModeBox, 3);
    targetCard.Controls.Add(previousCaseButton, 4, 1);
    targetCard.Controls.Add(nextCaseButton, 5, 1);
    layout.Controls.Add(targetCard, 0, 2);

    var actions = new FlowLayoutPanel
    {
      AutoSize = true,
      Dock = DockStyle.Fill,
      BackColor = UiTheme.Surface,
      Padding = new Padding(10, 7, 10, 7),
      WrapContents = false,
      Margin = new Padding(0),
    };
    captureScreenButton.Text = "画像をキャプチャ";
    StyleButton(captureScreenButton, primary: true);
    captureScreenButton.MinimumSize = new Size(172, 32);
    captureScreenButton.Click += async (_, _) => await CaptureScreenAsync();
    actions.Controls.Add(captureScreenButton);
    actions.Controls.Add(new Label { AutoSize = true, Text = "履歴", Margin = new Padding(18, 7, 4, 0), ForeColor = UiTheme.TextMuted });
    undoButton.Text = "元に戻す";
    StyleButton(undoButton);
    undoButton.Enabled = false;
    undoButton.Click += async (_, _) => await UndoAsync();
    redoButton.Text = "やり直す";
    StyleButton(redoButton);
    redoButton.Enabled = false;
    redoButton.Click += async (_, _) => await RedoAsync();
    actions.Controls.Add(undoButton);
    actions.Controls.Add(redoButton);
    actions.MinimumSize = actions.PreferredSize;
    layout.Controls.Add(actions, 0, 3);

    var statusPanel = new Panel
    {
      Height = 36,
      Dock = DockStyle.Fill,
      BackColor = UiTheme.SurfaceMuted,
      Padding = new Padding(10, 7, 10, 7),
      Margin = new Padding(0, 8, 0, 0),
    };
    statusLabel.AutoEllipsis = true;
    statusLabel.Dock = DockStyle.Fill;
    statusLabel.TextAlign = ContentAlignment.MiddleLeft;
    statusLabel.Text = "Ready";
    statusLabel.ForeColor = UiTheme.TextMuted;
    statusPanel.Controls.Add(statusLabel);
    layout.Controls.Add(statusPanel, 0, 4);

    Controls.Add(layout);
    UpdateSideButtonColors();
    EnsureResponsiveLayout(layout);
    Shown += async (_, _) =>
    {
      EnsureResponsiveLayout(layout);
      await RefreshWorkbooksAsync();
    };
    DpiChanged += (_, _) =>
    {
      if (CanUpdateUi) BeginInvoke(() => EnsureResponsiveLayout(layout));
    };
  }

  private void EnsureResponsiveLayout(TableLayoutPanel layout)
  {
    if (IsDisposed || Disposing) return;

    layout.PerformLayout();
    var preferred = layout.GetPreferredSize(Size.Empty);
    var requiredClientSize = new Size(
      Math.Max(760, preferred.Width),
      Math.Max(370, preferred.Height));
    MinimumSize = SizeFromClientSize(requiredClientSize);
    if (ClientSize.Width < requiredClientSize.Width || ClientSize.Height < requiredClientSize.Height)
    {
      ClientSize = new Size(
        Math.Max(ClientSize.Width, requiredClientSize.Width),
        Math.Max(ClientSize.Height, requiredClientSize.Height));
    }
  }

  private static TableLayoutPanel CreateCard(int columnCount) => new()
  {
    AutoSize = true,
    Dock = DockStyle.Fill,
    BackColor = UiTheme.Surface,
    Padding = new Padding(10, 8, 10, 8),
    ColumnCount = columnCount,
    RowCount = 3,
    Margin = new Padding(0),
  };

  private void StyleButton(Button button, bool primary = false) => UiTheme.StyleButton(button, Font, primary);

  private static void StyleSideButton(RadioButton button)
  {
    button.Appearance = Appearance.Button;
    button.AutoSize = true;
    button.MinimumSize = new Size(64, 32);
    button.Padding = new Padding(12, 2, 12, 2);
    button.TextAlign = ContentAlignment.MiddleCenter;
    button.UseCompatibleTextRendering = false;
    button.FlatStyle = FlatStyle.Flat;
    button.FlatAppearance.BorderColor = UiTheme.Border;
    button.FlatAppearance.CheckedBackColor = UiTheme.Primary;
    button.FlatAppearance.MouseOverBackColor = UiTheme.SurfaceMuted;
    button.Margin = new Padding(0, 0, 4, 0);
  }

  private static void StyleTextBox(TextBox textBox)
  {
    UiTheme.StyleTextBox(textBox);
    textBox.Margin = new Padding(0);
  }

  private void UpdateSideButtonColors()
  {
    ApplySideButtonColor(newSideButton);
    ApplySideButtonColor(oldSideButton);
  }

  private static void ApplySideButtonColor(RadioButton button)
  {
    var selected = button.Checked;
    button.BackColor = selected ? UiTheme.Primary : UiTheme.Surface;
    button.ForeColor = selected ? Color.White : UiTheme.Text;
    button.FlatAppearance.BorderColor = selected
      ? UiTheme.PrimaryHover
      : UiTheme.Border;
  }

  private static Label CreateLabel(string text) => new()
  {
    AutoSize = true,
    MinimumSize = new Size(0, 32),
    Text = text,
    Font = new Font("Meiryo UI", 9F, FontStyle.Bold),
    ForeColor = UiTheme.Text,
    TextAlign = ContentAlignment.MiddleLeft,
    UseCompatibleTextRendering = false,
    Anchor = AnchorStyles.Left,
  };

  private async Task RefreshPlacementContextAsync(bool force = false)
  {
    if (placementContextRefreshInProgress)
    {
      placementContextRefreshPending = true;
      placementContextRefreshPendingForce |= force;
      return;
    }

    placementContextRefreshInProgress = true;
    try
    {
      await RefreshPlacementContextCoreAsync(force);
    }
    finally
    {
      placementContextRefreshInProgress = false;
      if (placementContextRefreshPending && CanUpdateUi)
      {
        var pendingForce = placementContextRefreshPendingForce;
        placementContextRefreshPending = false;
        placementContextRefreshPendingForce = false;
        var appliedLatestSelection = !pendingForce && latestSelectionChange is not null &&
          TryApplyCachedSelection(latestSelectionChange);
        if (appliedLatestSelection)
        {
          latestSelectionChange = null;
        }
        else
        {
          BeginInvoke(async () => await RefreshPlacementContextAsync(pendingForce));
        }
      }
    }
  }

  private async Task RefreshPlacementContextCoreAsync(bool force)
  {
    var requestVersion = Interlocked.Increment(ref placementContextRequestVersion);
    if (workbookSelector.SelectedItem is not WorkbookIdentity workbook ||
      Volatile.Read(ref mutationInProgress) != 0 ||
      (!force && (!settings.FollowExcelSelection || placementTargetOverridden)))
    {
      return;
    }

    var analysis = await StaTask.Run(() => automaticPlacementService.Analyze(
      workbook,
      "ActiveSheet",
      SelectedSide,
      [new AutomaticPlacementImage("context", new ImageDimensions(1, 1))],
      preferActiveGap: false,
      horizontalMarginPoints: settings.HorizontalMarginPoints,
      autoDetectSide: true,
      sameCaseThenNext: settings.AdvanceMode is PlacementAdvanceMode.SameCaseThenNext));
    if (requestVersion != Volatile.Read(ref placementContextRequestVersion) || !CanUpdateUi)
    {
      return;
    }

    if (analysis.Succeeded)
    {
      SetPlacementContext(analysis);
    }
    else
    {
      SetStatus($"配置先を解析できません: {analysis.Message}");
    }
  }

  private string? RequestedCaseLabel =>
    !string.IsNullOrWhiteSpace(caseLabelBox.Text)
      ? caseLabelBox.Text.Trim()
      : null;

  private void SetPlacementContext(AutomaticPlacementAnalysisResult analysis)
  {
    cachedPlacementContext = analysis;
    ApplySideLayout(analysis.LayoutAnalysis!.Layout!.Kind);
    SetPlacementContext(analysis.WorksheetName, analysis.CaseLabel, analysis.ResolvedSide, overridden: false);
  }

  private void ApplySideLayout(SideLayoutKind kind)
  {
    var wasUpdating = updatingPlacementContext;
    updatingPlacementContext = true;
    try
    {
      oldSideButton.Enabled = kind == SideLayoutKind.Both;
      sideToolTip.SetToolTip(oldSideButton, oldSideButton.Enabled ? "旧側に配置" : "このシートはNewのみ");
      if (!oldSideButton.Enabled) newSideButton.Checked = true;
    }
    finally { updatingPlacementContext = wasUpdating; }
  }

  private async Task RefreshManualSideLayoutAsync()
  {
    if (updatingPlacementContext || !placementTargetOverridden ||
      workbookSelector.SelectedItem is not WorkbookIdentity workbook ||
      string.IsNullOrWhiteSpace(worksheetNameBox.Text) || Volatile.Read(ref mutationInProgress) != 0) return;
    var version = Interlocked.Increment(ref placementContextRequestVersion);
    var sheet = worksheetNameBox.Text.Trim();
    var label = RequestedCaseLabel;
    var captured = await StaTask.Run(() => new ExcelSheetSnapshotService().Capture(workbook, sheet));
    if (!CanUpdateUi || version != Volatile.Read(ref placementContextRequestVersion)) return;
    if (captured.Snapshot is not { } snapshot) { SetStatus(captured.Message); return; }
    var row = string.IsNullOrWhiteSpace(label) ? snapshot.ActiveCell.Row :
      ExcelAutomaticPlacementService.ConfirmedAnchors(snapshot.LayoutSignals)
        .FirstOrDefault(anchor => ExcelAutomaticPlacementService.FormatCaseLabel(anchor) == CaseAnchorNormalizer.NormalizeCaseLabel(label))?.Row ?? 0;
    var layout = placementContextLayoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = row });
    if (layout.Layout is { } resolved)
    {
      ApplySideLayout(resolved.Kind);
      SetStatus($"{sheet} / CASE {label} / {SelectedSide}  構成: {(resolved.Kind == SideLayoutKind.NewOnly ? "Newのみ" : "New/Old")}");
    }
    else SetStatus(string.Join(" ", layout.Reasons));
  }

  private void SetPlacementContext(
    string worksheetName,
    string caseLabel,
    EvidenceSide side,
    bool overridden)
  {
    updatingPlacementContext = true;
    try
    {
      worksheetNameBox.Text = worksheetName;
      caseLabelBox.Text = caseLabel;
      oldSideButton.Checked = side is EvidenceSide.Old;
      newSideButton.Checked = side is EvidenceSide.New;
      placementTargetOverridden = overridden;
    }
    finally
    {
      updatingPlacementContext = false;
    }
  }

  private void MarkPlacementTargetOverridden()
  {
    if (!updatingPlacementContext)
    {
      placementTargetOverridden = true;
    }
  }

  private void SessionMonitorSelectionChanged(object? sender, ExcelSelectionChangedEventArgs eventArgs)
  {
    if (!IsHandleCreated || IsDisposed || Disposing)
    {
      return;
    }

    BeginInvoke(() =>
    {
      if (!settings.FollowExcelSelection || placementTargetOverridden ||
        workbookSelector.SelectedItem is not WorkbookIdentity workbook ||
        workbook.ProcessId != eventArgs.ProcessId ||
        !WorkbookPathMatches(workbook, eventArgs.WorkbookFullPath) ||
        Volatile.Read(ref mutationInProgress) != 0)
      {
        return;
      }

      latestSelectionChange = eventArgs;
      if (TryApplyCachedSelection(eventArgs))
      {
        latestSelectionChange = null;
        return;
      }

      selectionChangeTimer.Stop();
      selectionChangeTimer.Start();
    });
  }

  private bool TryApplyCachedSelection(ExcelSelectionChangedEventArgs selection)
  {
    var signals = cachedPlacementContext?.LayoutSignals;
    if (signals is null || selection.Row < 1 || selection.Column < 1 ||
      !string.Equals(cachedPlacementContext!.WorksheetName, selection.WorksheetName, StringComparison.Ordinal))
    {
      return false;
    }

    var analyzed = placementContextLayoutAnalyzer.Analyze(signals with { ActiveRow = selection.Row });
    if (!analyzed.IsSafe || analyzed.Layout is null)
    {
      return false;
    }

    var anchor = ExcelAutomaticPlacementService.ConfirmedAnchors(signals)
      .Where(candidate => candidate.Row <= selection.Row)
      .OrderBy(candidate => candidate.Row)
      .LastOrDefault();
    if (anchor is null)
    {
      return false;
    }

    ApplySideLayout(analyzed.Layout.Kind);
    var side = analyzed.Layout.OldRegion is { } old && selection.Column >= old.FirstColumn &&
      selection.Column <= old.LastColumn
        ? EvidenceSide.Old
        : selection.Column >= analyzed.Layout.NewRegion.FirstColumn &&
          selection.Column <= analyzed.Layout.NewRegion.LastColumn
            ? EvidenceSide.New
            : SelectedSide;
    var label = ExcelAutomaticPlacementService.FormatCaseLabel(anchor);
    SetPlacementContext(selection.WorksheetName, label, side, overridden: false);
    return true;
  }

  private static bool WorkbookPathMatches(WorkbookIdentity workbook, string eventPath) =>
    string.IsNullOrWhiteSpace(eventPath) ||
    string.Equals(workbook.FullPath, eventPath, StringComparison.OrdinalIgnoreCase) ||
    string.Equals(workbook.Name, eventPath, StringComparison.OrdinalIgnoreCase);

  private void SaveAdvanceMode()
  {
    if (advanceModeBox.SelectedIndex < 0)
    {
      return;
    }

    settings = settings with
    {
      AdvanceMode = advanceModeBox.SelectedIndex == 1
        ? PlacementAdvanceMode.NextCaseSameSide
        : PlacementAdvanceMode.SameCaseThenNext,
    };
    TrySaveSettings();
  }

  private void TrySaveSettings()
  {
    try
    {
      settingsStore.Save(settings);
    }
    catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
    {
      WriteDiagnostic(DiagnosticEventKind.Application, DiagnosticOutcome.Failed, exception: exception);
    }
  }

  private async Task RefreshWorkbooksAsync()
  {
    refreshButton.Enabled = false;
    workbookSelector.Enabled = false;
    try
    {
      var previousId = (workbookSelector.SelectedItem as WorkbookIdentity)?.ConnectionId;
      statusLabel.Text = "Excel Workbookを検索しています…";
      pendingDiscoveryTask ??= StaTask.Run(sessionCatalog.Discover);
      var discoveryTask = pendingDiscoveryTask!;
      var completedTask = await Task.WhenAny(
        discoveryTask,
        Task.Delay(ExcelDiscoveryTimeout));
      if (completedTask != discoveryTask)
      {
        statusLabel.Text = "Excel探索がタイムアウトしました。Excelの応答後に更新を再試行してください。";
        return;
      }

      pendingDiscoveryTask = null;
      var result = await discoveryTask;
      if (IsDisposed || Disposing)
      {
        return;
      }

      var monitorWarnings = await RefreshSessionMonitorAsync(result.Workbooks);
      if (IsDisposed || Disposing)
      {
        return;
      }

      workbookSelector.BeginUpdate();
      try
      {
        workbookSelector.Items.Clear();
        foreach (var workbook in result.Workbooks)
        {
          _ = workbookSelector.Items.Add(workbook);
        }
      }
      finally
      {
        workbookSelector.EndUpdate();
      }

      var previousItem = WorkbookSelectionPolicy.Resolve(previousId, result.Workbooks);
      workbookSelector.SelectedItem = previousItem;
      if (previousId is not null && previousItem is null)
      {
        workbookSelector.SelectedIndex = -1;
      }

      statusLabel.Text = result.Workbooks.Count == 0
        ? "起動中のExcel Workbookが見つかりません。"
        : previousItem is null
          ? $"{result.Workbooks.Count} Workbookを検出しました。対象を選択してください。"
          : $"{result.Workbooks.Count} Workbookを検出しました。選択を維持しています。";
      var warningCount = result.Warnings.Count + monitorWarnings.Count;
      if (warningCount > 0)
      {
        statusLabel.Text += $" 警告: {warningCount}件 / {result.Warnings.Concat(monitorWarnings).First()}";
      }

      if (clipboardWarning is not null)
      {
        statusLabel.Text += $" {clipboardWarning}";
      }

      WriteDiagnostic(
        DiagnosticEventKind.ExcelDiscovery,
        DiagnosticOutcome.Succeeded,
        itemCount: result.Workbooks.Count);
      if (previousItem is not null)
      {
        await RefreshPlacementContextAsync(force: true);
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      statusLabel.Text = $"Excel列挙エラー: {exception.Message}";
      WriteDiagnostic(DiagnosticEventKind.ExcelDiscovery, DiagnosticOutcome.Failed, exception: exception);
    }
    finally
    {
      if (!IsDisposed && !Disposing)
      {
        refreshButton.Enabled = true;
        workbookSelector.Enabled = true;
      }
    }
  }

  private async Task<IReadOnlyList<string>> RefreshSessionMonitorAsync(
    IReadOnlyList<WorkbookIdentity> workbooks)
  {
    var warnings = new List<string>();
    var requestKey = CreateWorkbookSetKey(workbooks);
    var timeoutTask = Task.Delay(ExcelMonitorTimeout);

    if (pendingMonitorTask is not null)
    {
      var previousKey = pendingMonitorKey;
      if (await Task.WhenAny(pendingMonitorTask, timeoutTask) != pendingMonitorTask)
      {
        InvalidateMonitorRequests(workbooks);
        return ["Excel終了監視がタイムアウトしました。操作時の直接検証を使用します。"];
      }

      warnings.AddRange(await ConsumePendingMonitorAsync());
      if (string.Equals(previousKey, requestKey, StringComparison.Ordinal))
      {
        return warnings;
      }
    }

    pendingMonitorWorkbooks = workbooks.ToArray();
    pendingMonitorKey = requestKey;
    pendingMonitorTask = sessionMonitor.RefreshAsync(pendingMonitorWorkbooks);
    if (await Task.WhenAny(pendingMonitorTask, timeoutTask) != pendingMonitorTask)
    {
      InvalidateMonitorRequests(workbooks);
      warnings.Add("Excel終了監視がタイムアウトしました。操作時の直接検証を使用します。");
      return warnings;
    }

    warnings.AddRange(await ConsumePendingMonitorAsync());
    return warnings;
  }

  private async Task<IReadOnlyList<string>> ConsumePendingMonitorAsync()
  {
    var task = pendingMonitorTask ?? throw new InvalidOperationException("No session monitor refresh is pending.");
    var monitoredWorkbooks = pendingMonitorWorkbooks;
    try
    {
      return await task;
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      sessionMonitor.InvalidateSessions(monitoredWorkbooks);
      return [$"Excel終了監視に失敗しました。操作時の直接検証を使用します: {exception.Message}"];
    }
    finally
    {
      if (ReferenceEquals(pendingMonitorTask, task))
      {
        pendingMonitorTask = null;
        pendingMonitorWorkbooks = [];
        pendingMonitorKey = null;
      }
    }
  }

  private void InvalidateMonitorRequests(IReadOnlyList<WorkbookIdentity> currentWorkbooks)
  {
    sessionMonitor.InvalidateSessions(
      pendingMonitorWorkbooks
        .Concat(currentWorkbooks)
        .DistinctBy(item => item.ProcessId)
        .ToArray());
  }

  private static string CreateWorkbookSetKey(IReadOnlyList<WorkbookIdentity> workbooks) =>
    string.Join('|', workbooks.Select(item => item.ConnectionId).Order(StringComparer.Ordinal));

  private void QueueClipboardPreview()
  {
    if (!IsHandleCreated || IsDisposed || Disposing)
    {
      return;
    }

    if (clipboardPreviewOpen)
    {
      clipboardPreviewPending = true;
      return;
    }

    try
    {
      BeginInvoke(TryPreviewClipboardImage);
    }
    catch (InvalidOperationException) when (IsDisposed || Disposing)
    {
      // The form closed while the preview callback was being queued.
    }
  }

  private async void TryPreviewClipboardImage()
  {
    if (clipboardPreviewOpen)
    {
      clipboardPreviewPending = true;
      return;
    }

    Image? clipboardImage = null;
    uint sequenceBeforeRead = 0;
    try
    {
      sequenceBeforeRead = NativeClipboard.GetClipboardSequenceNumber();
      if (clipboardRetrySequence != 0 && sequenceBeforeRead != clipboardRetrySequence)
      {
        ResetClipboardRetry();
      }

      if (sequenceBeforeRead != 0 && lastClipboardSequenceNumber == sequenceBeforeRead)
      {
        ResetClipboardRetry();
        return;
      }

      if (!Clipboard.ContainsImage())
      {
        var sequenceAfterRead = NativeClipboard.GetClipboardSequenceNumber();
        if (sequenceBeforeRead != 0 && sequenceAfterRead != 0 && sequenceBeforeRead != sequenceAfterRead)
        {
          clipboardPreviewPending = true;
          return;
        }

        if (sequenceAfterRead != 0)
        {
          lastClipboardSequenceNumber = sequenceAfterRead;
        }

        ResetClipboardRetry();
        return;
      }

      clipboardImage = Clipboard.GetImage();
      if (clipboardImage is null)
      {
        ScheduleClipboardRetry(sequenceBeforeRead, "Clipboard画像をまだ取得できません。");
        return;
      }

      using var imageCopy = new Bitmap(clipboardImage);
      var sequenceAfterImageRead = NativeClipboard.GetClipboardSequenceNumber();
      if (sequenceBeforeRead != 0 &&
        sequenceAfterImageRead != 0 &&
        sequenceBeforeRead != sequenceAfterImageRead)
      {
        clipboardPreviewPending = true;
        return;
      }

      if (sequenceAfterImageRead != 0)
      {
        lastClipboardSequenceNumber = sequenceAfterImageRead;
      }

      ResetClipboardRetry();
      await ShowImagePreviewAsync(imageCopy, "Clipboard");
    }
    catch (ExternalException exception)
    {
      ScheduleClipboardRetry(sequenceBeforeRead, $"Clipboardを読み取れませんでした: {exception.Message}");
    }
    finally
    {
      clipboardImage?.Dispose();
      if (!clipboardPreviewOpen && clipboardPreviewPending && IsHandleCreated && !IsDisposed && !Disposing)
      {
        clipboardPreviewPending = false;
        QueueClipboardPreview();
      }
    }
  }

  private async Task CaptureScreenAsync()
  {
    if (screenCaptureInProgress || clipboardPreviewOpen ||
      Volatile.Read(ref mutationInProgress) != 0 || IsDisposed || Disposing)
    {
      return;
    }

    screenCaptureInProgress = true;
    captureScreenButton.Enabled = false;
    try
    {
      Hide();
      await Task.Delay(150);
      using var capture = new ScreenCaptureDialog();
      if (capture.ShowDialog() != DialogResult.OK)
      {
        SetStatus("画面キャプチャをキャンセルしました。");
        return;
      }

      using var image = capture.TakeCapturedImage();
      Show();
      Activate();
      await ShowImagePreviewAsync(image, "キャプチャ");
    }
    catch (Exception exception) when (exception is ExternalException or InvalidOperationException)
    {
      SetStatus($"画面をキャプチャできませんでした: {exception.Message}");
    }
    finally
    {
      if (!IsDisposed && !Disposing)
      {
        Show();
        Activate();
        captureScreenButton.Enabled = true;
      }
      screenCaptureInProgress = false;
    }
  }

  private async Task ShowImagePreviewAsync(Image image, string sourceLabel)
  {
    var workbook = workbookSelector.SelectedItem as WorkbookIdentity;
    var worksheetName = string.IsNullOrWhiteSpace(worksheetNameBox.Text)
      ? "ActiveSheet"
      : worksheetNameBox.Text.Trim();
    var requestedCaseLabel = RequestedCaseLabel;
    var temporaryDirectory = Path.Combine(Path.GetTempPath(), "EvidenceCrafter");
    var imagePath = Path.Combine(temporaryDirectory, $"preview-{Guid.NewGuid():N}.png");
    try
    {
      Directory.CreateDirectory(temporaryDirectory);
      using var imageCopy = new Bitmap(image);
      imageCopy.Save(imagePath, ImageFormat.Png);
      var request = new AutomaticPlacementImage(imagePath, ToImageDimensions(imageCopy));
      SetStatus("配置予定のCASE／Sideを解析しています…");
      var analysis = workbook is null
        ? AutomaticPlacementAnalysisResult.Failed("Workbookが選択されていません。")
        : await StaTask.Run(() => automaticPlacementService.Analyze(
          workbook,
          worksheetName,
          SelectedSide,
          [request],
          preferActiveGap: false,
          horizontalMarginPoints: settings.HorizontalMarginPoints,
          requestedCaseLabel: requestedCaseLabel,
          sameCaseThenNext: settings.AdvanceMode is PlacementAdvanceMode.SameCaseThenNext));

      if (analysis.Succeeded)
      {
        SetPlacementContext(analysis);
      }

      using var preview = new PreviewDialog(
        new Bitmap(image),
        workbook?.DisplayLabel ?? "未選択",
        worksheetName,
        SelectedSide,
        analysis);
      clipboardPreviewOpen = true;
      DialogResult previewResult;
      try
      {
        previewResult = preview.ShowDialog(this);
      }
      finally
      {
        clipboardPreviewOpen = false;
      }

      if (previewResult == DialogResult.Yes && workbook is not null)
      {
        await PlaceClipboardImageAutomaticallyAsync(
          workbook, worksheetName, SelectedSide, image, requestedCaseLabel, imagePath, analysis);
        imagePath = string.Empty;
      }
      else if (previewResult == DialogResult.Retry && workbook is not null)
      {
        using var editor = new ImageEditorDialog(image);
        if (editor.ShowDialog(this) == DialogResult.OK)
        {
          using var editedImage = editor.GetEditedImage();
          await PlaceClipboardImageAutomaticallyAsync(
            workbook, worksheetName, SelectedSide, editedImage, requestedCaseLabel);
        }
        else
        {
          SetStatus("画像編集をキャンセルしました。Excelは変更していません。");
        }
      }
      else
      {
        SetStatus($"{sourceLabel}画像を確認しました。Excelは変更していません。");
      }
    }
    finally
    {
      if (!string.IsNullOrEmpty(imagePath))
      {
        try { File.Delete(imagePath); } catch (IOException) { }
      }
    }
  }

  protected override bool ProcessCmdKey(ref Message message, Keys keyData)
  {
    if (keyData == (Keys.Control | Keys.Z))
    {
      _ = UndoAsync();
      return true;
    }

    if (keyData == (Keys.Control | Keys.Y))
    {
      _ = RedoAsync();
      return true;
    }

    if (keyData == Keys.F5)
    {
      _ = RefreshWorkbooksAsync();
      return true;
    }

    return base.ProcessCmdKey(ref message, keyData);
  }

  private async Task PlaceClipboardImageAutomaticallyAsync(
    WorkbookIdentity workbook,
    string worksheetName,
    EvidenceSide side,
    Image image,
    string? requestedCaseLabel = null,
    string? preparedImagePath = null,
    AutomaticPlacementAnalysisResult? preparedAnalysis = null)
  {
    if (!TryBeginMutation())
    {
      return;
    }

    var temporaryDirectory = Path.Combine(Path.GetTempPath(), "EvidenceCrafter");
    var imagePath = preparedImagePath ?? Path.Combine(temporaryDirectory, $"automatic-{Guid.NewGuid():N}.png");
    try
    {
      using var imageCopy = new Bitmap(image);
      if (preparedImagePath is null)
      {
        Directory.CreateDirectory(temporaryDirectory);
        imageCopy.Save(imagePath, ImageFormat.Png);
      }
      var dimensions = ToImageDimensions(imageCopy);
      SetStatus("CASE／Sideを解析して自動配置しています…");
      var result = await StaTask.Run(() => automaticPlacementService.PlaceImages(
        workbook,
        worksheetName,
        side,
        [new AutomaticPlacementImage(imagePath, dimensions)],
        preferActiveGap: false,
        settings.HorizontalMarginPoints,
        preparedAnalysis,
        requestedCaseLabel));
      SetStatus(result.Message);
      if (!result.Succeeded)
      {
        return;
      }

      using var stream = new MemoryStream();
      imageCopy.Save(stream, ImageFormat.Png);
      RowDeletionSnapshot? cleanupSnapshot = null;
      if (result.Analysis?.CompletesCaseAfterPlacement == true)
      {
        SetStatus("New／OldがそろったためCASE末尾を整理しています…");
        var cleanup = await StaTask.Run(() => caseMaintenanceService.TrimCompletedCaseTail(
          workbook,
          result.PlacedImages[^1].WorksheetName,
          result.PlacedImages[^1].Target.Metadata.AnchorCell.Row));
        cleanupSnapshot = cleanup.Changed ? cleanup.DeletionSnapshot : null;
      }
      AddAutomaticPlacementHistory(
        workbook,
        side,
        [new HistoryImage(stream.ToArray(), dimensions)],
        result.AppliedInsertions,
        result.PlacedImages,
        cleanupSnapshot,
        result.Analysis!.LayoutAnalysis!.Layout!);

      WriteDiagnostic(
        DiagnosticEventKind.MutationResult,
        DiagnosticOutcome.Succeeded,
        checked((int)workbook.ProcessId),
        result.PlacedImages.Count);
      await AdvanceAfterPlacementAsync(
        workbook,
        result.PlacedImages[^1].WorksheetName,
        result.Analysis?.CaseLabel ?? caseLabelBox.Text,
        side);
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"自動配置に失敗しました: {exception.Message}");
      WriteDiagnostic(
        DiagnosticEventKind.MutationResult,
        DiagnosticOutcome.Failed,
        checked((int)workbook.ProcessId),
        exception: exception);
    }
    finally
    {
      try
      {
        File.Delete(imagePath);
      }
      catch (IOException)
      {
      }
      EndMutation();
    }
  }

  private async Task PlaceClipboardImageAsync(
    WorkbookIdentity workbook,
    string worksheetName,
    EvidenceSide side,
    Image image)
  {
    if (!TryBeginMutation())
    {
      return;
    }

    var temporaryDirectory = Path.Combine(Path.GetTempPath(), "EvidenceCrafter");
    var imagePath = Path.Combine(temporaryDirectory, $"placement-{Guid.NewGuid():N}.png");
    try
    {
      if (!CanUpdateUi)
      {
        return;
      }

      Directory.CreateDirectory(temporaryDirectory);
      using (var imageCopy = new Bitmap(image))
      {
        imageCopy.Save(imagePath, ImageFormat.Png);
        SetStatus("Excelへ画像を配置しています…");
        var dimensions = ToImageDimensions(imageCopy);
        var result = await StaTask.Run(() => imagePlacementService.PlaceImage(
          workbook,
          worksheetName,
          requestedCell: null,
          side,
          imagePath,
          dimensions,
          horizontalMarginPoints: settings.HorizontalMarginPoints));
        if (CanUpdateUi)
        {
          SetStatus(result.Message);
          if (result.Succeeded)
          {
            using var stream = new MemoryStream();
            imageCopy.Save(stream, ImageFormat.Png);
            AddImagePlacementHistory(
              workbook,
              result.WorksheetName,
              side,
              result.FocusCell,
              dimensions,
              stream.ToArray(),
              result.Target!,
              settings.HorizontalMarginPoints);
          }
        }
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"画像配置に失敗しました: {exception.Message}");
    }
    finally
    {
      try
      {
        if (File.Exists(imagePath))
        {
          File.Delete(imagePath);
        }
      }
      catch (IOException)
      {
        // A temporary image left open by Excel is harmless and will be cleaned by the OS.
      }
      EndMutation();
    }
  }

  private static ImageDimensions ToImageDimensions(Image image)
  {
    // Screen captures are pixel data. Image DPI metadata varies by monitor and
    // encoder and must not change their visible size in Excel.
    return ScreenImageSizing.FromPixels(image.Width, image.Height);
  }

  private async Task InsertRowsAsync()
  {
    if (!TryGetMutationTarget(out var workbook, out var worksheetName))
    {
      return;
    }

    var count = checked((int)insertRowCountBox.Value);
    if (MessageBox.Show(
          $"Excelで現在選択中のセルの上へ {count} 行を挿入します。Workbookは自動保存されません。続行しますか？",
          "行挿入の確認",
          MessageBoxButtons.YesNo,
          MessageBoxIcon.Warning,
          MessageBoxDefaultButton.Button2) is not DialogResult.Yes)
    {
      return;
    }

    if (!TryBeginMutation())
    {
      return;
    }

    SetRowActionsEnabled(false);
    try
    {
      SetStatus("Excelへ行を挿入しています…");
      var result = await StaTask.Run(() => rowMutationService.InsertRowsAtActiveCell(
        workbook,
        worksheetName,
        count));
      SetStatus(result.Message);
      if (result.Succeeded && result.Changed)
      {
        AddRowInsertionHistory(workbook, result.WorksheetName, result.StartRow, result.Count);
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"行挿入に失敗しました: {exception.Message}");
    }
    finally
    {
      SetRowActionsEnabled(true);
      EndMutation();
    }
  }

  private async Task NavigateCaseAsync(CaseNavigationDirection direction)
  {
    if (!TryGetMutationTarget(out var workbook, out var worksheetName) || !TryBeginMutation())
    {
      return;
    }

    try
    {
      SetStatus(direction is CaseNavigationDirection.Previous ? "前のCASEを検索しています…" : "次のCASEを検索しています…");
      var result = await StaTask.Run(() => caseNavigationService.Navigate(
        workbook,
        worksheetName,
        direction,
        caseLabelBox.Text,
        SelectedSide,
        settings.AdvanceMode is PlacementAdvanceMode.SameCaseThenNext));
      SetStatus(result.Message);
      ApplyNavigationResult(result);
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"CASE移動に失敗しました: {exception.Message}");
    }
    finally
    {
      EndMutation();
    }
  }

  private async Task AdvanceAfterPlacementAsync(
    WorkbookIdentity workbook,
    string worksheetName,
    string caseLabel,
    EvidenceSide side)
  {
    SetStatus("次の空いているCASE／Sideを検索しています…");
    var result = await StaTask.Run(() => caseNavigationService.Navigate(
      workbook,
      worksheetName,
      CaseNavigationDirection.Next,
      caseLabel,
      side,
      settings.AdvanceMode is PlacementAdvanceMode.SameCaseThenNext));
    SetStatus(result.Message);
    if (result.Succeeded)
    {
      ApplyNavigationResult(result);
    }
  }

  private void ApplyNavigationResult(CaseNavigationResult result)
  {
    if (!result.Succeeded)
    {
      return;
    }

    ApplySideLayout(result.LayoutKind);
    cachedPlacementContext = null;
    updatingPlacementContext = true;
    try
    {
      worksheetNameBox.Text = result.WorksheetName;
      caseLabelBox.Text = result.CaseLabel;
      newSideButton.Checked = result.Side is EvidenceSide.New;
      oldSideButton.Checked = result.Side is EvidenceSide.Old;
      placementTargetOverridden = true;
    }
    finally
    {
      updatingPlacementContext = false;
    }
  }

  private async Task DeleteTrailingRowsAsync()
  {
    if (!TryGetMutationTarget(out var workbook, out var worksheetName))
    {
      return;
    }

    if (!int.TryParse(deleteCaseStartBox.Text, out var caseStartRow) ||
      !int.TryParse(deleteCaseEndBox.Text, out var caseEndRow) ||
      caseStartRow < 1 ||
      caseEndRow < caseStartRow ||
      caseEndRow > ExcelWorksheetLimits.MaximumRow)
    {
      SetStatus("削除対象Caseの開始行と終了行を入力してください。");
      return;
    }

    if (MessageBox.Show(
          $"Case {caseStartRow}～{caseEndRow} の末尾から、安全条件を満たす行だけを削除します。値・数式・コメント・リンク・Shape・結合セルがある行は削除しません。続行しますか？",
          "安全な末尾行削除の確認",
          MessageBoxButtons.YesNo,
          MessageBoxIcon.Warning,
          MessageBoxDefaultButton.Button2) is not DialogResult.Yes)
    {
      return;
    }

    if (!TryBeginMutation())
    {
      return;
    }

    SetRowActionsEnabled(false);
    try
    {
      SetStatus("Excelの行安全性を確認しています…");
      var result = await StaTask.Run(() => rowMutationService.DeleteTrailingRowsWithSnapshot(
        workbook,
        worksheetName,
        caseStartRow,
        caseEndRow,
        tailRows: 4));
      SetStatus(result.Message);
      if (result.Succeeded && result.Changed)
      {
        AddRowDeletionHistory(
          workbook,
          result.WorksheetName,
          result.StartRow,
          result.Count,
          result.DeletionSnapshot);
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"行削除に失敗しました: {exception.Message}");
    }
    finally
    {
      SetRowActionsEnabled(true);
      EndMutation();
    }
  }

  private async Task ReplaceSelectedImageAsync()
  {
    if (!TryGetMutationTarget(out var workbook, out _))
    {
      return;
    }

    if (!TryBeginMutation())
    {
      return;
    }

    Image? clipboardImage = null;
    AppliedRowInsertion? pendingGrowthInsertion = null;
    try
    {
      if (!Clipboard.ContainsImage() || (clipboardImage = Clipboard.GetImage()) is null)
      {
        SetStatus("差し替える画像をClipboardへコピーしてください。");
        return;
      }

      var selection = await StaTask.Run(() => managedShapeService.InspectSelection(workbook));
      if (!selection.Succeeded || selection.Shape is null)
      {
        SetStatus(selection.Message);
        return;
      }

      using var editor = new ImageEditorDialog(clipboardImage);
      if (editor.ShowDialog(this) != DialogResult.OK)
      {
        SetStatus("画像の差し替えをキャンセルしました。");
        return;
      }

      using var replacement = editor.GetEditedImage();
      if (MessageBox.Show(
          $"{selection.Shape.WorksheetName} の選択画像を差し替えます。自動保存されません。続行しますか？",
          "管理画像の差し替え",
          MessageBoxButtons.YesNo,
          MessageBoxIcon.Warning,
          MessageBoxDefaultButton.Button2) != DialogResult.Yes)
      {
        return;
      }

      var imagePath = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"replace-{Guid.NewGuid():N}.png");
      var originalPath = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"replace-original-{Guid.NewGuid():N}.png");
      try
      {
        Directory.CreateDirectory(Path.GetDirectoryName(imagePath)!);
        var exported = await ExportManagedShapePreservingClipboardAsync(
          workbook,
          selection.Shape,
          originalPath);
        if (!exported.Succeeded)
        {
          SetStatus(exported.Message);
          return;
        }

        var originalPng = await File.ReadAllBytesAsync(originalPath);
        using var replacementStream = new MemoryStream();
        replacement.Save(replacementStream, ImageFormat.Png);
        var replacementPng = replacementStream.ToArray();
        var replacementDimensions = ToImageDimensions(replacement);
        replacement.Save(imagePath, ImageFormat.Png);
        var prepared = await StaTask.Run(() => replacementLayoutService.EnsureSpace(
          workbook,
          selection.Shape,
          replacementDimensions,
          settings.HorizontalMarginPoints));
        if (!prepared.Succeeded)
        {
          if (prepared.Insertion is not null)
          {
            _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
              workbook,
              prepared.Insertion.WorksheetName,
              prepared.Insertion.StartRow,
              prepared.Insertion.Count));
          }
          SetStatus(prepared.Message);
          return;
        }
        pendingGrowthInsertion = prepared.Insertion;
        var replacementGeometry = prepared.FittedImage is null
          ? null
          : selection.Shape with
          {
            LeftPoints = prepared.LeftPoints ?? selection.Shape.LeftPoints,
            WidthPoints = prepared.FittedImage.WidthPoints,
            HeightPoints = prepared.FittedImage.HeightPoints,
          };
        var result = await StaTask.Run(() => managedShapeService.Replace(
          workbook,
          selection.Shape,
          imagePath,
          replacementDimensions,
          replacementGeometry));
        SetStatus(result.Message);
        if (result.Succeeded && result.After is not null)
        {
          var cleanup = await StaTask.Run(() => caseMaintenanceService.TrimCaseTail(
            workbook,
            result.After.WorksheetName,
            result.After.Metadata.AnchorCell.Row));
          AddManagedReplacementHistory(
            workbook,
            selection.Shape,
            result.After,
            originalPng,
            replacementPng,
            replacementDimensions,
            prepared.Insertion,
            cleanup.Changed ? cleanup.DeletionSnapshot : null);
          pendingGrowthInsertion = null;
        }
        else if (prepared.Insertion is not null)
        {
          _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
            workbook,
            prepared.Insertion.WorksheetName,
            prepared.Insertion.StartRow,
            prepared.Insertion.Count));
          pendingGrowthInsertion = null;
        }
        WriteDiagnostic(
          DiagnosticEventKind.MutationResult,
          result.Succeeded ? DiagnosticOutcome.Succeeded : DiagnosticOutcome.Rejected,
          checked((int)workbook.ProcessId));
      }
      finally
      {
        try
        {
          File.Delete(imagePath);
          File.Delete(originalPath);
        }
        catch (IOException)
        {
        }
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"画像の差し替えに失敗しました: {exception.Message}");
      WriteDiagnostic(DiagnosticEventKind.MutationResult, DiagnosticOutcome.Failed, exception: exception);
    }
    finally
    {
      if (pendingGrowthInsertion is not null)
      {
        _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
          workbook,
          pendingGrowthInsertion.WorksheetName,
          pendingGrowthInsertion.StartRow,
          pendingGrowthInsertion.Count));
      }
      clipboardPreviewOpen = false;
      clipboardImage?.Dispose();
      EndMutation();
    }
  }

  private async Task DeleteSelectedImageAsync()
  {
    if (!TryGetMutationTarget(out var workbook, out _))
    {
      return;
    }

    if (!TryBeginMutation())
    {
      return;
    }

    try
    {
      var selection = await StaTask.Run(() => managedShapeService.InspectSelection(workbook));
      if (!selection.Succeeded || selection.Shape is null)
      {
        SetStatus(selection.Message);
        return;
      }

      if (MessageBox.Show(
          $"{selection.Shape.WorksheetName} の選択画像を削除します。自動保存されません。続行しますか？",
          "管理画像の削除",
          MessageBoxButtons.YesNo,
          MessageBoxIcon.Warning,
          MessageBoxDefaultButton.Button2) != DialogResult.Yes)
      {
        return;
      }

      var originalPath = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"delete-original-{Guid.NewGuid():N}.png");
      try
      {
        var exported = await ExportManagedShapePreservingClipboardAsync(
          workbook,
          selection.Shape,
          originalPath);
        if (!exported.Succeeded)
        {
          SetStatus(exported.Message);
          return;
        }

        var originalPng = await File.ReadAllBytesAsync(originalPath);
        var result = await StaTask.Run(() => managedShapeService.Delete(workbook, selection.Shape));
        SetStatus(result.Message);
        if (result.Succeeded)
        {
          var cleanup = await StaTask.Run(() => caseMaintenanceService.TrimCaseTail(
            workbook,
            selection.Shape.WorksheetName,
            selection.Shape.Metadata.AnchorCell.Row));
          AddManagedDeletionHistory(
            workbook,
            selection.Shape,
            originalPng,
            cleanup.Changed ? cleanup.DeletionSnapshot : null);
        }
        WriteDiagnostic(
          DiagnosticEventKind.MutationResult,
          result.Succeeded ? DiagnosticOutcome.Succeeded : DiagnosticOutcome.Rejected,
          checked((int)workbook.ProcessId));
      }
      finally
      {
        clipboardPreviewOpen = false;
        try { File.Delete(originalPath); } catch (IOException) { }
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"画像削除に失敗しました: {exception.Message}");
      WriteDiagnostic(DiagnosticEventKind.MutationResult, DiagnosticOutcome.Failed, exception: exception);
    }
    finally
    {
      clipboardPreviewOpen = false;
      EndMutation();
    }
  }

  private bool TryGetMutationTarget(out WorkbookIdentity workbook, out string worksheetName)
  {
    workbook = null!;
    worksheetName = string.IsNullOrWhiteSpace(worksheetNameBox.Text)
      ? "ActiveSheet"
      : worksheetNameBox.Text.Trim();
    if (workbookSelector.SelectedItem is not WorkbookIdentity selectedWorkbook)
    {
      SetStatus("先に対象Workbookを選択してください。");
      return false;
    }

    workbook = selectedWorkbook;
    return true;
  }

  private async Task<ManagedShapeExportResult> ExportManagedShapePreservingClipboardAsync(
    WorkbookIdentity workbook,
    ManagedShapeTarget target,
    string outputPath)
  {
    using var clipboard = ClipboardSnapshot.Capture();
    clipboardPreviewOpen = true;
    try
    {
      var result = await StaTask.Run(() => managedShapeService.Export(workbook, target, outputPath));
      var exportSequence = NativeClipboard.GetClipboardSequenceNumber();
      if (clipboard.Restore(exportSequence))
      {
        lastClipboardSequenceNumber = NativeClipboard.GetClipboardSequenceNumber();
        clipboardPreviewPending = false;
      }
      else
      {
        clipboardPreviewPending = true;
      }
      return result;
    }
    finally
    {
      clipboardPreviewOpen = false;
      if (clipboardPreviewPending && CanUpdateUi)
      {
        clipboardPreviewPending = false;
        QueueClipboardPreview();
      }
    }
  }

  private void SetRowActionsEnabled(bool enabled)
  {
    if (!CanUpdateUi)
    {
      return;
    }

    insertRowCountBox.Enabled = enabled;
    deleteCaseStartBox.Enabled = enabled;
    deleteCaseEndBox.Enabled = enabled;
    insertRowsButton.Enabled = enabled;
    deleteRowsButton.Enabled = enabled;
  }

  private void AddImagePlacementHistory(
    WorkbookIdentity workbook,
    string worksheetName,
    EvidenceSide side,
    CellReference cell,
    ImageDimensions dimensions,
    byte[] png,
    ManagedShapeTarget initialTarget,
    double horizontalMarginPoints)
  {
    var target = initialTarget;
    AddHistory(new HistoryEntry(
      "画像配置",
      async () =>
      {
        var result = await StaTask.Run(() => managedShapeService.Delete(workbook, target));
        SetStatus(result.Message);
        return result.Succeeded;
      },
      async () =>
      {
        var imagePath = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"redo-{Guid.NewGuid():N}.png");
        try
        {
          Directory.CreateDirectory(Path.GetDirectoryName(imagePath)!);
          await File.WriteAllBytesAsync(imagePath, png);
          var result = await StaTask.Run(() => imagePlacementService.PlaceImage(
            workbook,
            worksheetName,
            cell,
            side,
            imagePath,
            dimensions,
            horizontalMarginPoints: horizontalMarginPoints));
          SetStatus(result.Message);
          if (result.Succeeded)
          {
            target = result.Target!;
          }

          return result.Succeeded;
        }
        finally
        {
          try
          {
            File.Delete(imagePath);
          }
          catch (IOException)
          {
          }
        }
      }));
  }

  private void AddAutomaticPlacementHistory(
    WorkbookIdentity workbook,
    EvidenceSide side,
    IReadOnlyList<HistoryImage> historyImages,
    IReadOnlyList<AppliedRowInsertion> insertions,
    IReadOnlyList<AutomaticPlacedImage> images,
    RowDeletionSnapshot? cleanupSnapshot,
    EvidenceCaseLayout expectedLayout)
  {
    if (historyImages.Count != images.Count)
    {
      throw new ArgumentException("History image count must match placed image count.", nameof(historyImages));
    }

    var targets = images.Select(image => image.Target).ToArray();
    var horizontalMarginPoints = settings.HorizontalMarginPoints;

    async Task<bool> ValidateHistoryLayoutAsync()
    {
      var captured = await StaTask.Run(() => new ExcelSheetSnapshotService().Capture(workbook, images[0].WorksheetName));
      var current = captured.Snapshot is { } snapshot
        ? placementContextLayoutAnalyzer.Analyze(snapshot.LayoutSignals with { ActiveRow = images[0].FocusCell.Row }).Layout
        : null;
      if (current is not null && current.Kind == expectedLayout.Kind &&
        current.NewRegion == expectedLayout.NewRegion && current.OldRegion == expectedLayout.OldRegion) return true;
      SetStatus("画像配置時からシート構成が変わったか確認できないため、Undo/Redoを停止しました。");
      return false;
    }

    async Task<bool> PlaceAtAsync(int index)
    {
      var imagePath = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"history-auto-{Guid.NewGuid():N}.png");
      try
      {
        Directory.CreateDirectory(Path.GetDirectoryName(imagePath)!);
        await File.WriteAllBytesAsync(imagePath, historyImages[index].Png);
        var placed = await StaTask.Run(() => imagePlacementService.PlaceImage(
          workbook,
          images[index].WorksheetName,
          images[index].FocusCell,
          side,
              imagePath,
          historyImages[index].Dimensions,
              images[index].Plan.Image.WidthPoints,
          horizontalMarginPoints));
        if (placed.Succeeded)
        {
          targets[index] = placed.Target!;
        }
        else
        {
          SetStatus(placed.Message);
        }

        return placed.Succeeded;
      }
      finally
      {
        try { File.Delete(imagePath); } catch (IOException) { }
      }
    }

    AddHistory(new HistoryEntry(
      "自動配置",
      async () =>
      {
        if (!await ValidateHistoryLayoutAsync()) return false;
        if (cleanupSnapshot is not null &&
          !await RunRowHistoryOperationAsync(() => rowMutationService.RestoreDeletedRows(workbook, cleanupSnapshot)))
        {
          return false;
        }
        var deletedShapeIndexes = new List<int>();
        var deletedInsertions = new List<AppliedRowInsertion>();
        for (var index = targets.Length - 1; index >= 0; index--)
        {
          var deleted = await StaTask.Run(() => managedShapeService.Delete(workbook, targets[index]));
          if (!deleted.Succeeded)
          {
            foreach (var restoreIndex in deletedShapeIndexes.AsEnumerable().Reverse())
            {
              _ = await PlaceAtAsync(restoreIndex);
            }
            if (cleanupSnapshot is not null)
            {
              _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot));
            }
            SetStatus(deleted.Message);
            return false;
          }
          deletedShapeIndexes.Add(index);
        }

        for (var index = insertions.Count - 1; index >= 0; index--)
        {
          var insertion = insertions[index];
          if (!await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
            workbook,
            insertion.WorksheetName,
            insertion.StartRow,
            insertion.Count)))
          {
            foreach (var deletedInsertion in deletedInsertions.AsEnumerable().Reverse())
            {
              _ = await RunRowHistoryOperationAsync(() => rowMutationService.InsertRows(
                workbook,
                deletedInsertion.WorksheetName,
                new RowInsertion(deletedInsertion.StartRow, deletedInsertion.Count, "Compensate failed automatic Undo.")));
            }
            foreach (var restoreIndex in deletedShapeIndexes.AsEnumerable().Reverse())
            {
              _ = await PlaceAtAsync(restoreIndex);
            }
            if (cleanupSnapshot is not null)
            {
              _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot));
            }
            return false;
          }
          deletedInsertions.Add(insertion);
        }

        SetStatus("自動配置を元に戻しました。");
        return true;
      },
      async () =>
      {
        if (!await ValidateHistoryLayoutAsync()) return false;
        var inserted = new List<AppliedRowInsertion>();
        var placedIndexes = new List<int>();
        foreach (var insertion in insertions)
        {
          if (!await RunRowHistoryOperationAsync(() => rowMutationService.InsertRows(
            workbook,
            insertion.WorksheetName,
            new RowInsertion(insertion.StartRow, insertion.Count, "Redo automatic placement."))))
          {
            foreach (var applied in inserted.AsEnumerable().Reverse())
            {
              _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
                workbook, applied.WorksheetName, applied.StartRow, applied.Count));
            }
            return false;
          }
          inserted.Add(insertion);
        }

        for (var index = 0; index < images.Count; index++)
        {
          if (!await PlaceAtAsync(index))
          {
            foreach (var placedIndex in placedIndexes.AsEnumerable().Reverse())
            {
              _ = await StaTask.Run(() => managedShapeService.Delete(workbook, targets[placedIndex]));
            }
            foreach (var applied in inserted.AsEnumerable().Reverse())
            {
              _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
                workbook, applied.WorksheetName, applied.StartRow, applied.Count));
            }
            return false;
          }
          placedIndexes.Add(index);
        }

        if (cleanupSnapshot is not null &&
          !await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot)))
        {
          foreach (var placedIndex in placedIndexes.AsEnumerable().Reverse())
          {
            _ = await StaTask.Run(() => managedShapeService.Delete(workbook, targets[placedIndex]));
          }
          foreach (var applied in inserted.AsEnumerable().Reverse())
          {
            _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
              workbook, applied.WorksheetName, applied.StartRow, applied.Count));
          }
          return false;
        }

        SetStatus("自動配置をやり直しました。");
        return true;
      },
      cleanupSnapshot is null ? null : cleanupSnapshot.Dispose));
  }

  private void AddManagedReplacementHistory(
    WorkbookIdentity workbook,
    ManagedShapeTarget original,
    ManagedShapeTarget replacement,
    byte[] originalPng,
    byte[] replacementPng,
    ImageDimensions replacementDimensions,
    AppliedRowInsertion? growthInsertion,
    RowDeletionSnapshot? cleanupSnapshot)
  {
    var current = replacement;
    AddHistory(new HistoryEntry(
      "画像差し替え",
      async () =>
      {
        if (cleanupSnapshot is not null &&
          !await RunRowHistoryOperationAsync(() => rowMutationService.RestoreDeletedRows(workbook, cleanupSnapshot)))
        {
          return false;
        }
        var result = await ReplaceManagedFromBytesAsync(
          workbook,
          current,
          originalPng,
          new ImageDimensions(original.WidthPoints, original.HeightPoints),
          original);
        if (result.Succeeded && result.After is not null)
        {
          current = result.After;
          if (growthInsertion is not null &&
            !await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
              workbook,
              growthInsertion.WorksheetName,
              growthInsertion.StartRow,
              growthInsertion.Count)))
          {
            var compensated = await ReplaceManagedFromBytesAsync(
              workbook,
              current,
              replacementPng,
              replacementDimensions,
              replacement);
            if (compensated.After is not null)
            {
              current = compensated.After;
            }
            if (cleanupSnapshot is not null)
            {
              _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot));
            }
            return false;
          }
        }
        else if (cleanupSnapshot is not null)
        {
          _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot));
        }
        SetStatus(result.Message);
        return result.Succeeded;
      },
      async () =>
      {
        if (growthInsertion is not null &&
          !await RunRowHistoryOperationAsync(() => rowMutationService.InsertRows(
            workbook,
            growthInsertion.WorksheetName,
            new RowInsertion(growthInsertion.StartRow, growthInsertion.Count, "Redo managed-image growth."))))
        {
          return false;
        }
        var result = await ReplaceManagedFromBytesAsync(
          workbook,
          current,
          replacementPng,
          replacementDimensions,
          replacement);
        if (result.Succeeded && result.After is not null)
        {
          current = result.After;
          if (cleanupSnapshot is not null &&
            !await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot)))
          {
            var compensated = await ReplaceManagedFromBytesAsync(
              workbook,
              current,
              originalPng,
              new ImageDimensions(original.WidthPoints, original.HeightPoints),
              original);
            if (compensated.After is not null)
            {
              current = compensated.After;
            }
            if (growthInsertion is not null)
            {
              _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
                workbook,
                growthInsertion.WorksheetName,
                growthInsertion.StartRow,
                growthInsertion.Count));
            }
            return false;
          }
        }
        else if (growthInsertion is not null)
        {
          _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
            workbook,
            growthInsertion.WorksheetName,
            growthInsertion.StartRow,
            growthInsertion.Count));
        }
        SetStatus(result.Message);
        return result.Succeeded;
      },
      cleanupSnapshot is null ? null : cleanupSnapshot.Dispose));
  }

  private void AddManagedDeletionHistory(
    WorkbookIdentity workbook,
    ManagedShapeTarget original,
    byte[] originalPng,
    RowDeletionSnapshot? cleanupSnapshot)
  {
    ManagedShapeTarget? current = null;
    AddHistory(new HistoryEntry(
      "画像削除",
      async () =>
      {
        if (cleanupSnapshot is not null &&
          !await RunRowHistoryOperationAsync(() => rowMutationService.RestoreDeletedRows(workbook, cleanupSnapshot)))
        {
          return false;
        }
        var path = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"undo-delete-{Guid.NewGuid():N}.png");
        try
        {
          Directory.CreateDirectory(Path.GetDirectoryName(path)!);
          await File.WriteAllBytesAsync(path, originalPng);
          var result = await StaTask.Run(() => managedShapeService.Restore(workbook, original, path));
          current = result.After;
          SetStatus(result.Message);
          if ((!result.Succeeded || current is null) && cleanupSnapshot is not null)
          {
            _ = await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot));
          }
          return result.Succeeded && current is not null;
        }
        finally
        {
          try { File.Delete(path); } catch (IOException) { }
        }
      },
      async () =>
      {
        if (current is null)
        {
          return false;
        }

        var result = await StaTask.Run(() => managedShapeService.Delete(workbook, current));
        SetStatus(result.Message);
        if (!result.Succeeded)
        {
          return false;
        }
        if (cleanupSnapshot is not null &&
          !await RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, cleanupSnapshot)))
        {
          var path = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"compensate-delete-{Guid.NewGuid():N}.png");
          try
          {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.WriteAllBytesAsync(path, originalPng);
            var restored = await StaTask.Run(() => managedShapeService.Restore(workbook, original, path));
            current = restored.After;
          }
          finally
          {
            try { File.Delete(path); } catch (IOException) { }
          }
          return false;
        }
        return true;
      },
      cleanupSnapshot is null ? null : cleanupSnapshot.Dispose));
  }

  private async Task<ManagedShapeMutationResult> ReplaceManagedFromBytesAsync(
    WorkbookIdentity workbook,
    ManagedShapeTarget current,
    byte[] png,
    ImageDimensions dimensions,
    ManagedShapeTarget exactGeometry)
  {
    var path = Path.Combine(Path.GetTempPath(), "EvidenceCrafter", $"history-replace-{Guid.NewGuid():N}.png");
    try
    {
      Directory.CreateDirectory(Path.GetDirectoryName(path)!);
      await File.WriteAllBytesAsync(path, png);
      return await StaTask.Run(() => managedShapeService.Replace(
        workbook,
        current,
        path,
        dimensions,
        exactGeometry));
    }
    finally
    {
      try { File.Delete(path); } catch (IOException) { }
    }
  }

  private void AddRowInsertionHistory(
    WorkbookIdentity workbook,
    string worksheetName,
    int startRow,
    int count) =>
    AddHistory(new HistoryEntry(
      "行挿入",
      () => RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
        workbook,
        worksheetName,
        startRow,
        count)),
      () => RunRowHistoryOperationAsync(() => rowMutationService.InsertRows(
        workbook,
        worksheetName,
        new RowInsertion(startRow, count, "Redo EvidenceCrafter row insertion.")))));

  private void AddRowDeletionHistory(
    WorkbookIdentity workbook,
    string worksheetName,
    int startRow,
    int count,
    RowDeletionSnapshot? snapshot) =>
    AddHistory(snapshot is null
      ? new HistoryEntry(
      "行削除",
      () => RunRowHistoryOperationAsync(() => rowMutationService.InsertRows(
        workbook,
        worksheetName,
        new RowInsertion(startRow, count, "Undo EvidenceCrafter row deletion."))),
      () => RunRowHistoryOperationAsync(() => rowMutationService.DeleteRowsIfSafe(
        workbook,
        worksheetName,
        startRow,
        count)))
      : new HistoryEntry(
        "行削除",
        () => RunRowHistoryOperationAsync(() => rowMutationService.RestoreDeletedRows(workbook, snapshot)),
        () => RunRowHistoryOperationAsync(() => rowMutationService.DeleteRestoredRows(workbook, snapshot)),
        snapshot.Dispose));

  private async Task<bool> RunRowHistoryOperationAsync(Func<RowMutationResult> operation)
  {
    var result = await StaTask.Run(operation);
    SetStatus(result.Message);
    return result.Succeeded && result.Changed;
  }

  private void AddHistory(HistoryEntry entry)
  {
    if (undoHistory.Count >= 20)
    {
      var discarded = undoHistory.Last();
      var retained = undoHistory.Take(19).Reverse().ToArray();
      undoHistory.Clear();
      discarded.Cleanup?.Invoke();
      foreach (var item in retained)
      {
        undoHistory.Push(item);
      }
    }

    undoHistory.Push(entry);
    ClearHistoryStack(redoHistory);
    UpdateHistoryButtons();
  }

  private async Task UndoAsync()
  {
    if (undoHistory.Count == 0)
    {
      return;
    }

    if (!TryBeginMutation())
    {
      return;
    }

    var entry = undoHistory.Peek();
    SetMutationActionsEnabled(false);
    try
    {
      SetStatus($"{entry.Label}を元に戻しています…");
      if (await entry.Undo())
      {
        _ = undoHistory.Pop();
        redoHistory.Push(entry);
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"Undoに失敗しました: {exception.Message}");
    }
    finally
    {
      SetMutationActionsEnabled(true);
      EndMutation();
    }
  }

  private async Task RedoAsync()
  {
    if (redoHistory.Count == 0)
    {
      return;
    }

    if (!TryBeginMutation())
    {
      return;
    }

    var entry = redoHistory.Peek();
    SetMutationActionsEnabled(false);
    try
    {
      SetStatus($"{entry.Label}をやり直しています…");
      if (await entry.Redo())
      {
        _ = redoHistory.Pop();
        undoHistory.Push(entry);
      }
    }
    catch (Exception exception) when (exception is not OutOfMemoryException)
    {
      SetStatus($"Redoに失敗しました: {exception.Message}");
    }
    finally
    {
      SetMutationActionsEnabled(true);
      EndMutation();
    }
  }

  private void SetMutationActionsEnabled(bool enabled)
  {
    SetRowActionsEnabled(enabled);
    if (CanUpdateUi)
    {
      undoButton.Enabled = enabled && undoHistory.Count > 0;
      redoButton.Enabled = enabled && redoHistory.Count > 0;
    }
  }

  private bool TryBeginMutation()
  {
    if (Interlocked.CompareExchange(ref mutationInProgress, 1, 0) == 0)
    {
      SetMutationActionsEnabled(false);
      replaceImageButton.Enabled = false;
      deleteImageButton.Enabled = false;
      previousCaseButton.Enabled = false;
      nextCaseButton.Enabled = false;
      captureScreenButton.Enabled = false;
      return true;
    }

    SetStatus("別のExcel操作が完了するまでお待ちください。");
    return false;
  }

  private void EndMutation()
  {
    cachedPlacementContext = null;
    Interlocked.Exchange(ref mutationInProgress, 0);
    SetMutationActionsEnabled(true);
    if (CanUpdateUi)
    {
      replaceImageButton.Enabled = true;
      deleteImageButton.Enabled = true;
      previousCaseButton.Enabled = true;
      nextCaseButton.Enabled = true;
      captureScreenButton.Enabled = !screenCaptureInProgress;
    }
  }

  private void UpdateHistoryButtons()
  {
    if (CanUpdateUi)
    {
      undoButton.Enabled = undoHistory.Count > 0;
      redoButton.Enabled = redoHistory.Count > 0;
    }
  }

  private static void ClearHistoryStack(Stack<HistoryEntry> history)
  {
    while (history.TryPop(out var entry))
    {
      entry.Cleanup?.Invoke();
    }
  }

  private void ScheduleClipboardRetry(uint sequence, string reason)
  {
    if (!CanUpdateUi)
    {
      return;
    }

    var currentSequence = sequence == 0
      ? NativeClipboard.GetClipboardSequenceNumber()
      : sequence;
    if (currentSequence == 0)
    {
      SetStatus(reason);
      return;
    }

    if (clipboardRetrySequence != currentSequence)
    {
      clipboardRetrySequence = currentSequence;
      clipboardRetryAttempt = 0;
    }

    if (clipboardRetryAttempt >= ClipboardRetryLimit)
    {
      clipboardRetryTimer.Stop();
      SetStatus($"{reason} 再試行上限に達しました。再度キャプチャしてください。");
      return;
    }

    clipboardRetryAttempt++;
    clipboardRetryTimer.Interval = Math.Min(1_200, 150 * (1 << (clipboardRetryAttempt - 1)));
    clipboardRetryTimer.Stop();
    clipboardRetryTimer.Start();
    SetStatus($"{reason} 再試行します ({clipboardRetryAttempt}/{ClipboardRetryLimit})。");
  }

  private void SetStatus(string message)
  {
    if (CanUpdateUi)
    {
      statusLabel.Text = message;
      sideToolTip.SetToolTip(statusLabel, message);
    }
  }

  private void WriteDiagnostic(
    DiagnosticEventKind eventKind,
    DiagnosticOutcome outcome,
    int? processId = null,
    int? itemCount = null,
    TimeSpan? duration = null,
    Exception? exception = null)
  {
    if (settings.DiagnosticLoggingEnabled)
    {
      diagnosticLog.Write(eventKind, outcome, processId, itemCount, duration, exception);
    }
  }

  private void ResetClipboardRetry()
  {
    clipboardRetryTimer.Stop();
    clipboardRetrySequence = 0;
    clipboardRetryAttempt = 0;
  }

  private sealed record HistoryEntry(
    string Label,
    Func<Task<bool>> Undo,
    Func<Task<bool>> Redo,
    Action? Cleanup = null);

  private sealed record HistoryImage(byte[] Png, ImageDimensions Dimensions);

}
