# EvidenceCrafter 実装計画

- 文書版: 0.1
- 作成日: 2026-08-31
- 実装先: `C:\work\EvidenceCrafter`
- 参照専用: `C:\work\Macro\Case&Evidence`
- 状態: 完成仕様を実装済み。自動・一括配置、管理画像操作、行調整、前後Case移動、操作単位の補償付きUndo/Redoをレビュー候補として検証中

## 1. 結論

EvidenceCrafter は、既存の Case&Evidence が生成した Excel エビデンスシートを毎操作時に読み直し、スクリーンショットの取得、編集、配置、差し替え、削除、行調整を支援する Windows 専用 WinForms アプリとして新規開発する。

実装は次の順序で進める。

1. 参照資産と Excel レイアウトの契約を固定する。
2. Excel を変更しない読み取り・解析機能を完成させる。
3. COM に依存しない配置計画エンジンを単体テストで固める。
4. プレビュー付きの「そのまま貼る」で最小の利用可能版を作る。
5. 行追加・行削除、虫食い挿入、差し替え等の破壊的操作を段階的に追加する。
6. 画像編集、設定、ログ、ショートカットを追加する。
7. 単一 EXE、CI、Release 用成果物を検証する。

最大の技術リスクは UI ではなく、ユーザー編集後のシートから Case 境界と New/Old 領域を安全に再構築すること、および Excel に対する複数の変更を失敗時に整合した状態へ戻すことである。このため、解析の信頼度が不足する場合は書き込みを中止する設計を最初から採用する。

## 2. 入力資料と適用優先順位

仕様が食い違う場合は、次の順で判断する。

1. 現在のユーザー指示
2. 共有会話で明示的に確定した仕様
3. 現行の `escape/BetaEvidenceGenerator.bas` と `MacroModulesGuide.md`
4. `MacroTest.xlsm` と生成済み `.xlsx` の実レイアウト
5. 固定値や過去挙動からの推測

参照側には版ずれがある。現行 `.bas` とガイドは 2026-08 頃、`MacroTest.xlsm` とサンプルブックは 2026-03 頃の内容であり、CONFIG の項目や既定値が一致しない。したがって、現行コードとガイドを仕様源、サンプルブックをレイアウト検証資料として扱う。

### 2.1 参照資産から確認できたレイアウト

- Case の先頭行は既定で 3 行目、以降は既定 50 行刻みだが、行オフセットは設定可能である。
- Case の識別情報は Case 先頭行の A 列または B 列に書かれる。
- 実サンプルでは B 列のアンカーが 3、53、103 行目のように並び、先頭以外には全幅の上罫線がある。
- New 側は C 列開始で、既定は C:Q。Old 側はその直後の R 列開始で、既定は R:AF。
- New 側列数は変更できる。`NewOnly` では Old 側は既定列数を維持し、`Both` では Old 側も New 側と同じ列数へ追従する。
- New/Old 境界の縦罫線、Case 境界の横罫線、Old 見出しはオプションで消える可能性がある。
- 最終 Case の下端は、横罫線、New/Old 境界縦罫線の終端、実コンテンツ、UsedRange を組み合わせて判断する必要がある。
- 過去の差し替えブックでは可変幅を確認済み。2026-09-06の現行共通・個別ブックはNew `C:Q` / Old `R:AF`で、個別ブックは30行Caseと書式だけのUsedRange末尾を持つ。
- 共通エビデンスではCase下端より後ろに書式だけの行が残るため、SheetSnapshotは`RawUsedLastRow`と`LogicalEvidenceLastRow`を分けて収集する。
- `FIRST_DEST_ROW = 3`、`SLOT_HEIGHT = 50`、C/Q/R/AF は互換性フォールバック情報であり、主要判定条件にはしない。

### 2.2 参照フォルダの保護

- `C:\work\Macro\Case&Evidence` 内では作成、更新、削除、整形、テスト出力を行わない。
- 結合テストで参照サンプルを使う場合は、一時フォルダへコピーした複製だけを開く。
- テスト終了時に参照元の SHA-256 が変わっていないことを検証できるスクリプトを用意する。
- EvidenceCrafter へ移すのは外部仕様と匿名化した fixture であり、VBA 実装そのものは複製しない。

調査時点の基準ハッシュは次のとおり。

| 参照ファイル | SHA-256 |
|---|---|
| `escape/BetaEvidenceGenerator.bas` | `0C1B0F5D8108D74D738BB9BE9C34893A37F407B5634FE9BCF199BE57B74ECCF5` |
| `MacroModulesGuide.md` | `25319391FF27B63412B78CF6A2ED66CF2FDC18003436FBE8F706D2FA201D9B2B` |
| `MacroTest.xlsm` | `72618E3D34ED63EB716FCE7508A1F2D5CE93A6FBA7D03A46A6A346CBA45E4A3F` |
| `S00-000-00TESTX_【共通】02TEST2_単体テストエビデンス_初期開発.xlsx` | `4701BECC5D5DA6EA30D444118B301FB8A06D35FB1A559FC7012585063ED6B9DC` |
| `S00-000-00TESTX_【個別】02TEST2_単体テストエビデンス_初期開発.xlsx` | `DDA819E06E1BD51D931E3F60FABA915B7FA2940603C5AF6DD020C6507C68D571` |
| `S00-000-00TESTX_単体テストケース_初期開発_003.xlsx` | `C10C6D120FAF51E1FB1484EBEC7DE52F343BB8FB9265C76C8FCFC85D135FF666` |

## 3. v1 の機能範囲

### 3.0 対応環境の基準

- 正式サポート対象は、.NET 9で動作するWindowsデスクトップ版、およびサポート期間内のデスクトップExcelとする。
- 配布物は.NET 9ランタイムを同梱する自己完結型とし、利用端末への.NETランタイムの事前導入を不要とする。
- 現在の開発ホストは Windows 11 Pro build 22000、Office 2019 x86 である。どちらも正式リリースのサポート基準には採用せず、後方互換性を確認する参考環境として扱う。
- Phase 0 で実利用先の Windows edition/version、Excel version、Office bitness を棚卸しし、サポート対象表を確定する。
- ライフサイクル外の OS/Office が業務上必須の場合は、対応可否を実機試験し「best effort 互換」として明記する。公式サポート環境と同じ保証はしない。

### 3.1 必須機能

- 起動中の複数 Excel インスタンスと Workbook を列挙し、対象を明示選択する。
- 同名 Workbook はフルパスと Excel インスタンスで区別する。
- 選択 Workbook はセッション中固定し、対象変更はユーザー操作に限定する。
- 対象 Workbook の ActiveSheet と ActiveCell を操作直前に再取得する。
- ActiveCell から Evidence Case を自動判定する。
- New/Old はユーザーが明示切り替えし、画像追加後も自動切り替えしない。
- `Win + Shift + S` 後の Clipboard 画像を検知し、Excel を変更する前にプレビューする。
- 「編集せず配置」「編集して配置」「キャンセル」を提供し、配置方法は自動判定する。
- 同一 Case、同一 Side へ複数画像を縦配置する。
- 画像同士の間には実 Excel 行で最低 2 行分の非画像帯を確保し、Case 末尾は最低 4 行確保する。非画像帯のセルには補足文を入力できる。
- 横幅優先で縮小し、小さい画像は拡大しない。左右余白は初期値 6pt とする。
- 高さ不足時は Worksheet の行全体を追加する。
- Case 末尾の連続した不要行だけを安全条件付きで自動削除する。
- ActiveCell が Case 中央の空白を指す場合はその位置を優先し、空間不足ならプレビューに追加行数を表示して行を追加する。
- 編集機能として色選択対応の枠、矢印、枠付きテキスト、モザイク、トリミング、Undo、Redo、元画像へ戻すを提供する。
- Excel で選択中の管理画像を差し替え、削除できる。
- アプリから最後に追加した画像を取り消せる。
- Workbook を自動保存しない。
- ローカル設定と、画像本体を含まない診断ログを保存する。
- Primary RID は `win-x64` とし、self-contained の単一 `EvidenceCrafter.exe` を生成する。ただし Phase 2 で Office x86 との接続互換性を必ず検証する。

### 3.2 初版対象外

- macOS、Excel Online、Google Sheets
- Evidence Case 自体の生成、および Case&Evidence VBA の置換
- OCR、AI 画像認識、録画、クラウド同期、データベース
- Workbook の自動保存
- 管理外 Shape の編集・削除
- 複数画像の任意並べ替え
- `win-arm64` 配布物

## 4. 技術判断

| 項目 | 採用方針 | 理由 |
|---|---|---|
| Framework | `.NET 9` / `net9.0-windows` | .NET 9系の開発・実行環境に対応し、配布物は自己完結型とするため。 |
| UI | WinForms | Windows/Excel 専用で、Clipboard、グローバルホットキー、COM の STA 実行と相性がよい。 |
| Excel | ROT と late-bound COM を薄い adapter に隔離 | 既に起動中の Excel を扱いつつ、Office PIA の配布・バージョン依存を避けるため。実体は Excel adapter 内に閉じ込める。 |
| テスト | MSTest | 共有会話の確定事項。Core 単体テストと Office 必須の結合テストを分離する。 |
| 画像 | WinForms/System.Drawing 系、メモリ内処理 | 初版の編集要件を満たし、画像をディスクへ恒久保存しない。 |
| 配布 | Primary `win-x64`、self-contained、single-file、trim 無効 | 利用者の .NET 導入を不要にする。WinForms/COM/Interop の互換性リスクを避けるため trimming は使わない。Office x86 互換性は早期検証する。 |
| 設定 | `%LocalAppData%\EvidenceCrafter\settings.json` | ユーザー単位で持ち運び不要な設定を保存する。Side は起動時 New に戻す。 |
| ログ | `%LocalAppData%\EvidenceCrafter\logs` | Excel 接続、解析結果、変更計画、結果、例外を記録する。画像やセル本文は既定で記録しない。 |
| 管理 Shape | `EST_IMG_<GUID>` と `AlternativeText` のバージョン付きメタデータ | 管理画像だけを安全に差し替え・削除し、再起動後も識別する。 |

`PublishSingleFile=true`、`SelfContained=true`、`RuntimeIdentifier=win-x64`、`IncludeNativeLibrariesForSelfExtract=true` を primary publish profile に持たせる。単一 EXE であっても Microsoft Excel のインストールは必要であり、Office を同梱するものではない。

.NET の single-file は RID/アーキテクチャ別である。調査環境の Office は x86 のため、Phase 2 で少なくとも次を比較する。

- `win-x64` アプリから Office x86 の ROT/Workbook/Shape を扱う。
- `win-x86` アプリから Office x86 を扱う。
- Office x64 の検証環境が用意できる場合は `win-x64` でも同じ結合テストを行う。

`win-x64` と Office x86 の全機能が安定する場合は primary EXE 1 本を維持する。互換性に問題がある場合は、primary RID を実利用 Office に合わせるか、`EvidenceCrafter-win-x86.exe` と `EvidenceCrafter-win-x64.exe` の各単一ファイルを配布する。この判断は Phase 2 の実測をリリースゲートとする。

## 5. リポジトリ構成

```text
EvidenceCrafter/
├─ .github/
│  └─ workflows/
│     ├─ ci.yml
│     └─ release.yml
├─ docs/
│  ├─ implementation-plan.md
│  ├─ specification.md
│  ├─ architecture.md
│  ├─ reference-baseline.md
│  └─ testing.md
├─ src/
│  ├─ EvidenceCrafter.App/
│  ├─ EvidenceCrafter.Core/
│  └─ EvidenceCrafter.Excel/
├─ tests/
│  └─ EvidenceCrafter.Tests/
│     └─ Fixtures/
├─ artifacts/                 # Git 管理外
├─ bootstrap.ps1
├─ EvidenceCrafter.sln
├─ Directory.Build.props
├─ Directory.Packages.props
├─ global.json
├─ .editorconfig
├─ .gitignore
└─ README.md
```

### 5.1 依存方向

| Project | 責務 | 依存先 |
|---|---|---|
| `EvidenceCrafter.Core` | Case/Layout 解析、占有範囲、画像サイズ、配置、行変更計画、検証 | BCL のみ。WinForms、COM、Excel 型へ依存しない。 |
| `EvidenceCrafter.Excel` | ROT、Workbook 接続、Worksheet の snapshot 化、計画の適用、COM 解放 | Core の DTO/契約 |
| `EvidenceCrafter.App` | WinForms、Clipboard、プレビュー、画像編集、設定、ログ、操作調停 | Core、Excel |
| `EvidenceCrafter.Tests` | Core 単体、fixture 回帰、Excel 結合 | 対象 Project |

COM オブジェクトを Core やバックグラウンドスレッドへ渡さない。Excel から必要情報を一括取得して `SheetSnapshot` へ変換し、それ以降の解析と計画作成は純粋な C# ロジックで行う。

## 6. 主要モデルと境界

### 6.1 Core DTO

- `WorkbookIdentity`: Excel インスタンス、フルパス、Workbook 名、接続世代
- `SheetSnapshot`: Sheet、ActiveCell、UsedRange、行高、列幅、セル使用状態、罫線、結合範囲、Shape
- `CaseAnchor`: 候補行、A/B 値の有無、構造シグナル
- `EvidenceCaseLayout`: StartRow、EndRow、New/Old 領域、判定信頼度
- `ContentSpan`: セルまたは Shape が占有する行・列・Side
- `AffectedRowSnapshot`: 行全体の UsedRange、結合、Shape、リッチテキスト run を含む変更前 fingerprint
- `ImageSpec`: 元サイズ、編集後サイズ、縦横比
- `PlacementRequest`: Case、Side、ActiveCell ヒント、配置モード
- `PlacementPlan`: 画像位置、scale、追加行、削除候補、期待する事前状態
- `MutationPlan`: 行追加、行削除、Shape 操作、罫線/書式補修の順序
- `OperationJournal`: 直前取消と例外時の補償に必要な最小情報

### 6.2 主要サービス

- `IExcelSessionCatalog`
- `IWorksheetSnapshotReader`
- `ICaseLayoutAnalyzer`
- `IContentOccupancyAnalyzer`
- `IPlacementPlanner`
- `IRowMutationPlanner`
- `IWorksheetMutationExecutor`
- `IClipboardImageSource`
- `IImageEditorSession`
- `ISettingsStore`

## 7. 解析・配置アルゴリズム

### 7.1 Snapshot 取得

操作直前に以下を一括取得する。

- 対象 Workbook の存在、ReadOnly、接続状態
- ActiveSheet、ActiveCell、Worksheet 保護状態
- A/B 列の Case 候補値
- Case 近傍の罫線、行高、列幅、結合範囲
- 値、数式、コメント/メモ、ハイパーリンクの有無
- Shape の名前、種類、TopLeftCell/BottomRightCell、位置、サイズ、Placement
- 選択 Shape

値本文は Case 判定に必要な A/B と構造情報に限定し、ログへは出さない。COM 呼び出し回数を抑えるため、Range は可能な限り配列で読み込む。

### 7.2 Case 判定

1. A/B のいずれかに Case 情報がある行を Anchor 候補にする。
2. 先頭候補または上罫線、左右領域の境界、周辺レイアウトを用いて候補を検証する。
3. `Anchor.Row <= ActiveCell.Row` を満たす最も近い Anchor を Case 開始行とする。
4. 次 Anchor があれば `EndRow = NextAnchor.Row - 1` とする。
5. 最終 Case は、下罫線、縦境界線の終端、実コンテンツ、UsedRange を照合して終端を求める。
6. 固定 50 行は他のシグナルと一致した場合だけ補助に使う。
7. 複数候補が矛盾する、または最終 Case の終端が安全に決められない場合は `UnsafeLayout` として書き込みを中止する。

判定結果には `Confidence` と根拠を持たせ、プレビューと診断ログで説明できるようにする。

### 7.3 New/Old 領域判定

1. New 開始は互換契約として C 列を基点にする。
2. Case 内を通る縦罫線から New の右端を探す。
3. ヘッダー、新旧の列幅、Case 下罫線範囲を補助にする。
4. Old は New 右端の次列から開始する。
5. `NewOnly` と `Both` の両列構成を検証する。
6. 「旧」の文字列や Q/R 固定値だけでは判定しない。
7. 罫線 OFF、Old 見出し削除等により一意に決められない場合は書き込みを中止する。

### 7.4 使用中領域

以下のいずれかがあれば使用中とみなす。

- 値、数式
- コメント、メモ、ハイパーリンク
- 管理内外を問わない Shape/画像
- 結合セル

書式だけのセルは画像配置では空きとみなせるが、行削除時は罫線再構築が安全であることを別途検証する。

### 7.5 画像配置

- 配置可能幅は対象 Side のポイント幅から左右 6pt ずつを引く。
- `scale = min(1.0, availableWidth / imageWidth)` とし、縦横比を維持する。
- Shape は `LockAspectRatio = true`、`Placement = xlMove` とし、行移動には追従するが行高変更で画像自体を自動伸縮させない。
- 高さは各行の実 `RowHeight` を積算して必要行数へ変換する。
- Case が空なら Case 開始行を候補にする。
- 各画像の直後 2 行を「予約された非画像帯」として扱う。セル文字は許可するが、次画像や Shape は置かない。
- 補足文が予約 2 行内に収まる場合、次画像はその 2 行の直後から配置できる。予約帯より下にもセル/Shape コンテンツがある場合は、その最下端の後ろに新たな 2 行の非画像帯を確保してから配置する。
- これにより「2 行は完全な空セル」とはせず、「画像同士が最低 2 行分離され、セル内容と画像が重ならない」を不変条件にする。
- ActiveCell が空白帯を指す場合は虫食い候補を先に評価する。
- 虫食いが狭い場合は不足行をその位置へ追加する計画を作り、プレビューで「中央へ N 行追加」を明示する。
- 配置後は New/Old の低い方ではなく、両 Side の最下端の最大値から 4 行を残す。

### 7.6 行追加・削除

- Side 単独ではなく Worksheet の行全体を挿入・削除する。
- 隣接する Evidence 行を基準に高さ、書式、罫線を補完する。
- 削除対象は Case 末尾に連続する空行だけとし、中央の空行は削除しない。
- 行全体の削除は Evidence 領域外のセルも消すため、New/Old だけでなく対象行の UsedRange 全列について、値、数式、コメント/メモ、リンク、Shape、結合影響がないことを確認する。
- 次 Case へ侵入せず、末尾 4 行を下回らない範囲に限定する。
- 行変更後に Case/Side/Shape を再 snapshot し、期待した構造と一致することを検証する。
- 変更対象外セルのリッチテキスト、結合、行高、罫線、印刷レイアウトが意図せず変わらないことを結合テストで確認する。

## 8. Excel 変更の安全設計

### 8.1 Plan → Revalidate → Apply → Verify

1. Snapshot から不変な `MutationPlan` を作る。
2. 書き込み直前に軽量 snapshot を再取得し、Workbook/Sheet/Case/境界/対象 Shape の fingerprint を照合する。
3. 不一致なら計画を破棄して再解析する。
4. `ScreenUpdating`、`EnableEvents`、`DisplayAlerts`、`Calculation` の元状態を保存する。
5. 計画順に変更し、管理 Shape 名とメタデータを設定する。
6. 再解析して配置、余白、境界を検証する。
7. Application 状態を `finally` で必ず復元する。

### 8.2 失敗時契約

Excel COM は外部アプリであり、複数操作を完全なデータベース transaction にはできない。そのため「どの例外でも物理的に部分変更が絶対残らない」とは約束しない。

v1 では次を保証目標とする。

- 既知の妥当性エラーでは 1 件も書き込まない。
- 書き込み中の例外では `OperationJournal` による補償を逆順に試みる。
- 補償を完全に確認できない場合は自動保存せず、対象 Workbook、操作、復旧案を明示する。
- Excel 標準 Undo に依存しない。Interop 操作で Undo 履歴が維持される保証はないため、「直前取消」はアプリ管理機能として実装する。
- 変更中に Workbook/Sheet が切り替わっても、選択済み Workbook 以外には書き込まない。

### 8.3 COM ライフサイクル

- Excel 接続と全 COM 操作は STA で直列化する。
- `Range`、`Worksheet`、`Workbook` 等の操作RCWを長期保持しない。Workbook終了イベント用の`Application`/`IConnectionPoint`だけは専用STA内で購読期間を明示して保持する。
- chained property access と COM collection の暗黙 enumerator を避ける。
- COM 参照は取得元が明確な順序で解放し、強制 GC を通常の制御手段にはしない。
- Workbook 終了、Excel プロセス終了、RPC 切断を通常の状態遷移として扱う。

## 9. 実装フェーズ

### Phase 0: 仕様・参照契約の固定

成果物:

- `docs/specification.md`
- `docs/reference-baseline.md`
- 匿名化した Case/Layout JSON fixture
- .NET 9 採用、安全な書き込み契約、Shape メタデータのADR

作業:

- 現行 VBA/ガイドとサンプルブックの版ずれを文書化する。
- 実利用先の Windows/Excel/Office bitness を棚卸しし、正式サポートと best effort 互換の境界を確定する。
- 既定 50 行、行オフセット変更、NewOnly、Both、罫線 ON/OFF、Old 見出し削除の期待構造を fixture 化する。
- 虫食い挿入、左右 6pt、画像削除 v1 採用を仕様へ確定する。
- 参照側のハッシュ検証を用意する。

完了条件:

- 固定列・固定 50 行に依存しない判定ルールが例付きで説明されている。
- 判定不能時の停止条件とユーザー表示が定義されている。
- 参照フォルダに変更がない。

### Phase 1: リポジトリとビルド基盤

成果物:

- Solution、App/Core/Excel/Test project
- `Directory.Build.props`、中央 package 管理、`global.json`
- nullable 有効、警告をエラーとして扱う基本設定
- `bootstrap.ps1`、`.gitignore`、README、CI

作業:

- `bootstrap.ps1` は SDK 確認、restore、build、単体 test、任意 publish を再実行可能にする。
- `artifacts/`、ユーザー設定、ログ、Office 一時ファイルを Git 対象外にする。
- CI は Windows runner で build と Excel 非依存 test を実行する。

完了条件:

- clean checkout 相当から bootstrap が成功する。
- `dotnet build` と Excel 非依存 `dotnet test` が成功する。
- 空の App を single-file publish できる。

### Phase 2: Excel 読み取りスパイク

成果物:

- ROT ベースの `IExcelSessionCatalog`
- 読み取り専用 `SheetSnapshot` 作成器
- Case/Side 検出の診断画面または CLI テスト harness
- Excel 結合テスト基盤

作業:

- 複数 Excel インスタンス、複数 Workbook、同名 Workbook、未保存 Workbook を検証する。
- `win-x64`/`win-x86` アプリと 32/64-bit Office の組み合わせ、Workbook 終了競合、ReadOnly、保護 Sheet を検証する。
- COM 参照が残って Excel 終了を妨げないことを確認する。
- 参照サンプルは一時コピーでのみ開く。

完了条件:

- Excel を一切変更せず、選択 Workbook/ActiveSheet/ActiveCell/Case/Side 候補を表示できる。
- 複数インスタンス列挙の可否と制限が実機結果として文書化されている。
- 接続解除後に誤書き込みせず Disconnected へ遷移する。

### Phase 3: Core 解析・配置計画

成果物:

- `CaseLayoutAnalyzer`
- `ContentOccupancyAnalyzer`
- `ImageSizingService`
- `PlacementPlanner`
- `RowMutationPlanner`
- fixture ベース MSTest

完了条件:

- Case 先頭/末尾、最終 Case、動的 New/Old、罫線なし、判定不能をテストできる。
- 0/1/複数画像、最低 2 行の非画像帯、帯内の補足文、4 行余白、大小/縦長/横長画像をテストできる。
- 虫食い十分/不足/使用中 ActiveCell、中央行追加、末尾削除安全条件をテストできる。
- Core test は Excel プロセスなしで決定的に通る。

### Phase 4: 最小利用可能版 — そのまま貼る

成果物:

- Compact WinForms shell
- Workbook 選択、New/Old、前/次 Case
- Clipboard listener
- プレビュー
- 行追加不要な範囲での管理 Shape 貼り付け

完了条件:

- `Win + Shift + S` からプレビューを経て正しい Workbook/Sheet/Case/Side に貼れる。
- 通常の画像コピーを誤検知しても、承認前に Excel を変更しない。
- Clipboard の重複通知で同じ画像を二重追加しない。
- 画像は aspect ratio を保ち、小画像を拡大しない。

preview.12では上記に加え、Case自動解析を画像配置へ接続し、必要行の挿入、明示的な行操作、管理画像の差し替え・削除、画像本体と行書式を含むUndo/Redoを実装する。保存は安全境界としてExcel利用者の明示操作に残す。

### Phase 5: 行調整・複数画像・虫食い

成果物:

- 全行挿入/削除 executor
- 書式/罫線再構築
- 同一 Side 複数画像
- 虫食い位置への挿入と末尾 override

完了条件:

- New/Old の長い側を基準に末尾 4 行を維持する。
- 2 行の非画像帯へ補足文字を入力しても保持され、次画像が文字と重ならない。
- 行追加、削除、行高変更、Shape 手動移動後に再解析して追従する。
- Evidence 領域外を含む UsedRange のどこかに値、数式、コメント、リンク、結合セル、管理外 Shape がある行を削除しない。
- 次 Case を壊さず、変更後の罫線が参照レイアウトと一致する。
- 行操作後も対象外セルのリッチテキスト run、行高、結合、印刷レイアウトを保持する。

### Phase 6: 画像編集

成果物:

- 非破壊 editor session
- 色選択対応の枠、矢印、枠付きテキスト、モザイク、トリミング
- editor 内 Undo/Redo、元画像へ戻す

完了条件:

- 全編集を確定前に取り消せる。
- 100%、150%、200% DPI と複数モニターで座標がずれない。
- 編集前後で画像の縦横比と配置計算が正しい。
- キャンセル時に画像ファイルや Excel 変更を残さない。

### Phase 7: 差し替え・削除・直前取消

成果物:

- 選択管理 Shape の差し替え
- 選択管理 Shape の削除
- アプリ管理の直前貼り付け取消

完了条件:

- 管理外 Shape、複数選択、非画像選択では変更しない。
- 差し替え後も元の上端を維持し、高さ増減に応じて行を安全に調整する。
- 削除/取消後も Case 末尾 4 行と次 Case 境界を維持する。
- アプリ再起動後も Shape メタデータから管理画像を識別できる。

### Phase 8: 運用機能と堅牢化

成果物:

- LocalAppData 設定
- 診断ログとローテーション
- グローバルショートカット
- エラー/Disconnected UX
- 長時間稼働試験

完了条件:

- 前回 Workbook は自動接続せず、起動中の同一パスを候補表示するだけにする。
- Side は起動時 New に戻る。
- ホットキー競合時に GUI 操作は継続できる。
- ログに画像、セル本文、個人情報を既定で出さない。
- Excel の起動/終了を繰り返しても COM 参照とメモリが増え続けない。

### Phase 9: 配布・リリース

成果物:

- `artifacts/publish/EvidenceCrafter.exe`
- publish profile
- CI workflow
- tag 用 Release workflow
- 利用者向け README

完了条件:

- .NET Runtime 未導入かつ正式サポート対象の検証用 Windows 環境で EXE が起動する。
- 対応 Excel がない場合は明確な案内を出して終了する。
- EXE 以外のアプリ配布ファイルを要求しないことを clean VM で確認する。
- primary RID と Office x86/x64 の互換表を README に記載し、Phase 2 の判定どおりの成果物名にする。
- build、unit test、publish、ハッシュ生成が自動化されている。
- ExcelIntegration test は通常 CI から除外し、Office 導入済み環境で明示実行する。

## 10. テスト戦略

### 10.1 Core 単体テスト

最低限、次の組み合わせを MSTest で網羅する。

- Case: 既定、行追加後、行削除後、最終 Case、不正/曖昧レイアウト
- Side: 既定 15 列、New 列数変更、NewOnly、Both、Old 見出しなし、罫線なし
- Image: 大、小、縦長、横長、crop 後、aspect ratio
- Placement: 0/1/複数、2 行の非画像帯、帯内補足文、帯外補足文、4 行、New/Old の高さ差
- Gap: 十分、不足、ActiveCell が使用中、中央への行追加
- Delete: Evidence 内外の値、数式、コメント、リンク、Shape、結合、次 Case、中央空白
- Replace: 同高、高くなる、低くなる、下に画像/補足セルあり
- Safety: snapshot fingerprint 不一致、接続世代不一致、判定信頼度不足

### 10.2 Excel 結合テスト

`[TestCategory("ExcelIntegration")]` を付け、通常の unit test から除外する。

- Workbook/ActiveSheet/ActiveCell 取得
- 複数 Excel インスタンス列挙
- Shape 貼り付けと `xlMove`
- 行挿入/削除時の Shape anchor
- 罫線、書式、行高の再構築
- Evidence 領域外のセル内リッチテキスト run、結合、印刷設定の非破壊性
- 保護、ReadOnly、Workbook close、RPC 切断
- Close取消後のfail-closed・再探索と、旧watcherによる新token誤失効がないこと
- token失効・アプリ終了時のWorkbook Window property回収
- Windows前面化に依存しない対象Workbook/Sheet/Cell選択
- 一時コピー上での差し替え、削除、補償処理

### 10.3 手動 E2E

- Excel/Office の対応版と bitness
- 正式サポート Windows/Excel と、必要に応じた best effort 互換環境
- Windows 表示倍率 100/150/200%
- 単一/複数モニター
- 画像を 20 枚以上連続追加する長時間操作
- Excel 側で行、セル、Shape を途中編集した後の再開
- Workbook 未保存、同名 Workbook、ネットワーク/OneDrive 上の ReadOnly
- EXE の clean VM 起動

## 11. 品質ゲート

各 Phase の完了には次を要求する。

- 追加した仕様に対応する MSTest がある。
- Excel 非依存 test がすべて成功する。
- Excel 書き込みを含む Phase は、一時コピー上の結合テストと手動 E2E が成功する。
- nullable 警告と analyzer 警告を未説明のまま残さない。
- public API と複雑な private 処理には XML コメントまたは「なぜ」のコメントがある。
- COM 参照と Application 状態の復元が review 済みである。
- 参照フォルダの基準ハッシュが変わっていない。
- docs/specification、architecture、testing、README が実装と一致する。

## 12. リスク登録簿

| リスク | 影響 | 対策 |
|---|---|---|
| 参照コードとサンプルの版ずれ | 誤った fixture/判定 | 参照優先順位と SHA-256 を固定し、版ごとに fixture を持つ。 |
| Case/Side が一意に判定不能 | 別 Case/Side を破壊 | 信頼度付き解析、固定値単独フォールバック禁止、UnsafeLayout で停止。 |
| 行削除が Evidence 領域外の内容も消す | データ損失 | 対象行の UsedRange 全列を検査し、末尾連続行限定、結合/Shape/次 Case guard。 |
| COM 操作途中の例外 | 部分変更 | 事前検証、操作 journal、逆順補償、事後検証、自動保存禁止。 |
| Excel 標準 Undo が使えない | 復旧不能 | 直前取消をアプリ管理し、標準 Undo を受入条件にしない。 |
| 複数 Excel/COM 解放不良 | 誤接続、Excel が終了しない | ROT monikerとWorkbook固有Window token、Close確定のWindow消滅確認、旧watcherのtoken限定失効、専用STA直列化、PID単位fail-closed、RCW所有ルール、native Window property回収、失敗分離したcleanup、生成Excel PID終了試験。有効tokenだけを索引化して無効化時に回収する。ROT変更時刻は通常編集で変わるため接続世代には使わない。 |
| Clipboard 重複/別画像 | 二重・誤貼付 | Clipboard sequence、読取前後照合、上限付きbackoff、Preview再入guard、必ずプレビュー承認。 |
| DPI/複数モニター | 編集座標ずれ | device-independent 座標と代表 DPI の E2E。 |
| single-file と Interop | 起動失敗、複数ファイル化 | Phase 1 と clean VM で早期検証、trim 無効、publish profile 固定。 |
| EXE/Office の x86/x64 組み合わせ | ROT 列挙や COM 接続失敗 | Phase 2 で両 RID を実測し、primary RID または複数成果物を決定。 |
| 開発ホストの OS/Office がライフサイクル外 | 本番で再現しない、公式サポート不可 | 対応環境を Phase 0 で確定し、正式環境の clean VM でリリース試験。旧環境は best effort と明記。 |
| 行操作でリッチテキスト/帳票書式が変わる | 既存資料の破損 | 対象外セルの run/結合/行高/罫線/印刷設定を fingerprint と結合テストで保護。 |
| ログへの業務情報混入 | 情報漏えい | 画像・セル本文を記録しない、構造情報と ID のみに制限。 |

## 13. 現時点の仮決定

共有会話で再確認候補になっていた項目は、実装を止めないため次を初期値とする。

| 項目 | 計画上の初期決定 |
|---|---|
| Framework | .NET 9 / `net9.0-windows` |
| 狭い虫食い位置 | 行を追加して挿入し、プレビューで追加行数を明示 |
| 左右余白 | 6pt ずつ。設定変更可能 |
| 選択画像削除 | v1 に含める。管理 Shape のみ |
| グローバルショートカット | Phase 8で実装。設定で無効化でき、競合時は警告して通常操作を継続 |
| GitHub Release | Phase 9。通常 CI と Office 必須 test は分離 |
| 対応 OS/Excel | リリース時点の公式サポート対象を正式対応。現ホストの旧環境は best effort 互換試験のみ |

この初期決定を変更する場合でも、Phase 0 の仕様確定までなら後続設計への影響は限定的である。

## 14. 完成の定義

次をすべて満たした時点で v1 完成とする。

- 対応する起動中 Excel Workbook を安全に選択できる。
- ActiveSheet/ActiveCell と現在のシート状態から Case/New/Old を再構築できる。
- スクリーンショットをプレビューし、そのまままたは編集後に正しく配置できる。
- 複数画像、最低 2 行の非画像帯、4 行末尾、虫食い、行増減、差し替え、削除、直前取消が仕様どおり動く。
- 判定不能・保護・ReadOnly・切断・結合影響時に安全停止する。
- Workbook を自動保存せず、参照フォルダを一切変更しない。
- Core unit test、Excel integration test、手動 E2E の品質ゲートを通る。
- Primary RID では `artifacts/publish/EvidenceCrafter.exe` 1 ファイルを配布できる。複数 RID が必要と判定された場合も、各 RID の配布単位は単一 EXE とする。
- 新しい開発 PC で `bootstrap.ps1` により build/test/publish を再現できる。

## 15. 参照 URL

- 共有仕様: https://chatgpt.com/share/6a950606-a918-83ee-81bb-28ec577017bd
- .NET support policy: https://dotnet.microsoft.com/en-us/platform/support/policy
- .NET single-file deployment: https://learn.microsoft.com/dotnet/core/deploying/single-file/overview
- .NET 9 supported OS versions: https://github.com/dotnet/core/blob/main/release-notes/9.0/supported-os.md
- Excel 2019 lifecycle: https://learn.microsoft.com/lifecycle/products/excel-2019
