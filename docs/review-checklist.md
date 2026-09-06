# 0.1.0-preview.18 再レビューチェックリスト

## 仕様

- [ ] A/B列のすべての非空行をAnchor候補にしてよいか
- [ ] 最終Caseの下罫線/縦罫線不一致をUnsafeとするか
- [ ] Old見出しだけによるSide判定をMediumとして許可するか
- [ ] 2行を「空白」ではなく「非画像帯」とする定義でよいか
- [ ] 管理外Shapeを変更対象外とするか

## 技術

- [ ] Core DTOにExcel固有型が漏れていないか
- [ ] ROT列挙のCOM所有権が明確か
- [ ] UIがClipboard取得だけでExcelを書き換えないか
- [ ] package/RID/.NET 10方針が配布先に合うか

## 品質

- [ ] Unit testの境界ケースが十分か
- [ ] Reference hash guardが再現可能か
- [ ] single-file成果物がレビュー環境で起動するか
- [ ] 次の書き込みPhaseへ進む条件が明確か

## 再レビュー時の自動確認

- [x] Case開始行より前を末尾行削除候補にしない
- [x] 予約帯を跨ぐShape/Mergeへ画像を重ねない
- [x] New/Old双方の最下端から4行を確保する
- [x] Gap配置の前後2行を検証する
- [x] 非構造AnchorとOld見出し単独判定をUnsafeにする
- [x] Workbookを暗黙に選び直さない
- [x] Clipboard通知の再コピーと競合を区別する
- [x] 一時WorkbookでROT列挙と配置先フォーカスを実機検証する
- [x] Publish先を清掃し、EXEとSHA-256以外を拒否する
- [x] 差し替え後の共通・個別・テストケースWorkbookを読み取り専用で再確認する
- [x] 過去の狭幅Fixtureと現行実例 `C:Q` / `R:AF` の可変幅を回帰テストする
- [x] 参照ベースライン6件が再発行後も一致する
- [x] raw UsedRangeと論理Evidence終端を分離する
- [x] 同一Excelインスタンスの複数Workbookを別HWNDで識別する
- [x] Workbook固有ROT登録や終了イベント監視がなくても、操作時の直接照合が成功すればSnapshotとフォーカスを実行する
- [x] EnableEventsが無効でも接続トークンを維持し、フォーカス・画像配置・Undo後に元の無効状態を復元する
- [x] ROT内の無関係な取得不能モニカーがあっても対象WorkbookのSnapshotを継続する
- [x] 判定済みSheet名とCase番号を表示し、存在するCase番号への手動変更を配置計画へ反映する
- [x] 最小化状態でも配置セルの選択結果を返す（前面化は必須条件にしない）
- [x] 配置検証結果とフォーカス結果を別状態で返す
- [x] 探索中のWorkbook選択を固定し、タイムアウト後に操作を復旧する
- [x] Clipboard読取前後のsequenceを照合し、PreviewDialogの再入を防止する
- [x] 行挿入後のCase終端にもExcel最大行を適用する
- [x] フォーカス中のWorkbook/Worksheetイベントを抑止し、`EnableEvents`を必ず復元する
- [x] Workbook close/reopen後は旧identityを拒否する
- [x] 末尾行計算の巨大入力をExcel上限で拒否し、整数overflowを防止する
- [x] 非最終Caseを含む全Caseの終端を論理Evidence終端内に制限する
- [x] 探索タイムアウト後も同一Taskを保持し、完了結果を次回更新で回収する
- [x] Excel最大列を超えるSide境界と配置列を拒否する
- [x] Clipboard一時競合を回数上限付きbackoffで再試行する
- [x] Excel結合テストの遅延完了通知を破棄済み同期オブジェクトへ送らない
- [x] Clipboard重複対策の計画記述をsequence方式へ統一する
- [x] Session MonitorのROT/COM操作と解除をUI外の専用STAで実行する
- [x] Monitor timeout、イベント無効、callback失敗、Detach時にPID配下tokenを無効化する
- [x] 通常編集では接続IDを維持し、close/reopenではWindow session tokenを更新する
- [x] `Advise`失敗経路を含むConnection Point所有権を明示する
- [x] Excel結合テストでMonitor先行解除と生成Excel PID終了を確認する
- [x] Close取消時は現identityを維持し、確定Close後だけWindow tokenを無効化する
- [x] tokenの置換・個別無効化・プロセス解除でregistry状態を回収する
- [x] Close取消後の旧watcherは保存済みtokenだけを失効し、再発行tokenを維持する
- [x] token失効・アプリ終了時にPID固有のWorkbook Window propertyを除去する
- [x] 結合テストの各cleanupを独立実行し、生成Excel PIDと一時ファイルを失敗時も回収する
- [x] Windowsの前面化に依存せず、`Application.Goto`で配置セルを選択する
- [x] 終了gate後のDiscoverがtoken／Window propertyを再登録しない
- [x] 結合テストtimeout時に監督中の生成ExcelとSTAを回収する
- [x] Preview承認後に単一Shapeを追加し、Shape名を検証してから配置セルへFocusする
- [x] ReadOnlyまたは保護SheetではShapeを追加せず安全停止する
- [x] SideをShapeメタデータへ記録し、自動配置では解析済みSide列へ移動する
- [x] ActiveCellの上へ確認済みの指定行数を挿入する
- [x] 削除直前のlive snapshotで内容、数式、結合、コメント、ハイパーリンク、Shapeを安全判定する
- [x] 指定Case内の安全な末尾連続行だけを削除し、Case開始行より前は変更しない
- [x] 行Mutation中に抑止したExcelイベント、画面更新、警告表示を必ず復元する
- [x] 画像配置、行挿入、安全な行削除を同一Workbook接続世代内でUndo/Redoする
- [x] 挿入後に内容が追加された行はUndoで削除せず安全停止する
- [x] 自動配置PreviewでSheet、Case、Side、開始セル、方式、追加行数、画像幅を確認できる
- [x] Previewから選べる配置操作を編集あり／なしに限定し、どちらも自動配置する
- [x] Preview解析を配置処理へ渡し、配置直前に同一Snapshot fingerprintを検証する
- [x] 画面画像を96 DPIのWindows論理サイズへ正規化し、高DPIメタデータによる過小配置を防ぐ
- [x] 個別Evidence実例の30行Case、C:Q／R:AF境界、書式だけのUsedRange末尾を回帰テストする
- [x] 占有セルと行高の走査をActiveCellが属するCase内に限定する
- [x] UsedRangeに結合セルがない場合、セル単位のMergeCells COM呼び出しを省略する
