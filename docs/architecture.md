# CraftEvidence アーキテクチャ

## 依存方向

```text
CraftEvidence.App  ───────> CraftEvidence.Core
        │
        └───────────────> CraftEvidence.Excel ─────> CraftEvidence.Core

CraftEvidence.Tests ─────> CraftEvidence.Core
```

`Core` はExcel COM、WinForms、Clipboardへ依存しない。Excelから読み取った情報はimmutableなsignal/DTOへ変換し、解析と配置計画をCoreで行う。

## Project責務

### CraftEvidence.Core

- Case Anchorと最終Case境界の解析
- New/Old領域の解析
- 判定信頼度と理由の生成
- 画像サイズ計算
- 占有範囲と非画像帯を考慮した配置計画

### CraftEvidence.Excel

- Running Object Tableから起動中Excelを読み取り専用で探索
- COMオブジェクトをCore DTOへ変換
- 接続切断を通常状態として扱う
- ExcelMutationサービスがCoreの計画を変更直前に再検証して実行
- 明示承認された単一画像を、検証済みWorkbook/SheetのActiveCellへShapeとして配置
- 明示確認されたActiveCell直上の行挿入と、ライブ安全Snapshotに基づく連続末尾行削除
- 配置検証後に対象Workbook/Sheetと画像左上セルへフォーカスを移す
- Workbook終了イベントを監視し、閉じた接続世代を即時に無効化する

COMオブジェクトはadapter外へ公開しない。通常の列挙・フォーカスRCWは取得スコープ内で解放し、Core DTOや別Apartmentへ渡さない。

Workbook終了イベントの購読だけは明示的な例外である。`ExcelApplicationSessionMonitorHost` がメッセージループ付きの専用background STAを所有し、そのSTA内のMonitorだけがExcel `Application` と `IConnectionPoint` を購読期間中保持する。Refresh、callback、Unadvise、RCW解放は同じSTAで行う。UIは15秒を上限に結果を待ち、timeout時はPID配下の接続トークンを直ちに無効化して操作を継続する。終了時は有効tokenとPID固有のWorkbook Window propertyを同期的に除去してから、同STAへ購読解除を非同期通知し、応答しないExcelがUI終了を妨げないようにする。Process終了時にも同じnative property cleanupを再実行する。

### CraftEvidence.App

- Workbook選択
- New/Oldの明示選択
- Clipboard listener
- 画像プレビュー
- Preview承認と配置結果の表示
- 安全状態と診断理由の表示

## 解析結果

解析は成功/失敗だけでなく、`LayoutConfidence` と理由一覧を返す。

- `High`: 独立した複数信号が一致
- `Medium`: 安全に一意化できるが補助信号が不足
- `Unsafe`: 信号なし、または信号が矛盾

書き込み層は `Unsafe` を必ず拒否する。preview.12ではCase自動解析、Side列幅、既存コンテンツ占有、必要行挿入を単一の自動配置フローへ接続する。明示されたActiveCellへの単一Shape配置と、確認付き行Mutationも提供する。

## Excel変更の境界

自動書き込みは次の一方向フローとする。

```text
Snapshot -> Analyze -> Plan -> Revalidate -> Apply -> Verify -> Journal -> Focus
```

preview.12の実機フローは `ROT再接続 -> immutable Snapshot -> Case/Side計画 -> fingerprint再検証 -> EnableEvents=false -> 行挿入/Shape追加 -> Shape検証 -> 状態復元 -> ROT再接続Focus`。途中失敗時はShapeと挿入行を逆順補償する。明示行操作は `live safety snapshot -> TrailingRowDeletionPlanner -> identity再検証 -> native undo snapshot -> Rows.Resize.Insert/Delete -> 状態復元` とし、保存は行わない。

Snapshot fingerprintが変わった場合はPlanを破棄する。Excel COMをdatabase transactionとして扱わず、既知エラーの事前拒否と、例外時の逆順補償を分ける。

`Focus` は配置の `Verify` 成功後だけ実行する後処理である。`VerifiedPlacementFinalizer` が `PlacementPlan.FocusCell` を `ExcelPlacementFocusService` へ渡し、配置検証結果とフォーカス結果を別フィールドで返す。Focus service は保存済みの HWND を直接 COM 入口にせず、ROT から Workbook を再取得し、PID、名前、フルパス、ROT moniker、Workbook固有Window、接続トークンを再検証する。ROTの変更時刻は通常編集や保存でも変化し、close/reopen時の更新も保証されないため接続世代には使わない。`WorkbookBeforeClose`はclose確定通知ではないため、callback後もdocument windowとWindow propertyを非COMで監視し、Window消滅後だけcallback時点の接続トークンを無効化する。Closeがcancelされた間は元identityを維持する。取消後にfail closedと再探索が続いた場合も、旧watcherはWorkbook keyから引き直さず保存済みtokenだけを失効するため、新tokenを削除しない。監視解除・callback失敗・`EnableEvents=false`・timeout時はPID配下を一括無効化するため、同一Excelプロセス・同一パス・HWND再利用時も旧identityを拒否する。token registryは現在有効なtokenだけを保持し、置換・個別無効化・PID解除で各索引と対応するnative Window propertyから回収する。property除去前には現在値が対象tokenと一致することを確認し、再発行済みpropertyを消さない。イベント監視を確立できないExcelプロセスではFocus自体を無効にしてfail closedとする。

Workbook/SheetのActivateと `Application.Goto(..., true)` は `Application.EnableEvents=false` の区間内で実行し、成功・失敗のいずれでも元の値を復元する。前面化は必須条件にせず、`Goto` による対象Workbook／Sheet／セルの選択結果を返す。これはセル内容を書き換えず、利用者マクロのWorkbook/Worksheetイベントも発火させない。

Workbook探索はbackground STAの同一Taskを完了まで保持する。15秒を超えた場合はUI操作を復旧するがTaskを破棄せず、次回更新で完了結果を回収する。Clipboardはsequenceの読取前後一致を条件とし、一時競合時だけ回数上限付き指数backoffで再試行する。
