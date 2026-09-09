# 図形・行高情報読み取りの改善（2026-09-09）

## 変更と維持した動作

画像配置前の解析、配置直前の再検証、CASE移動はいずれも
`ExcelSheetSnapshotService.ReadShapes` でシート全体の図形情報を読む。
60個の図形の読み取りだけで修正前は中央値682.85 msかかっていた。

- 図形の左上・右下について、それぞれRowとColumnを個別取得していた処理を、
  絶対R1C1形式のAddressから両方取得する処理に変更した。
  管理画像1個あたりのプロパティ／Item取得は9回から7回になる。
- 管理画像の命名規則に合わない図形はAlternativeTextを読んでも管理画像にならないため、
  その取得を省略する（9回から6回）。
- Addressが取得できない、形式が異なる、またはExcelの行列上限に収まらない場合は、
  従来のRow／Column取得に戻す。
- 行高は、範囲の先頭行の高さと範囲全体の高さが整合する場合だけ均一と判断する。
  混在時は範囲を二分して確認し、一部だけ異なる行や非表示行も実値として取得する。

Addressの引数は絶対参照・R1C1・外部参照なしを明示する。
Excel画面の参照形式に依存しない指定は
[MicrosoftのRange.Address仕様](https://learn.microsoft.com/en-us/office/vba/api/excel.range.address)に従う。

スナップショットのキャッシュや安全確認の省略は行わない。
配置セル、サイズ、CASE／Side遷移、行操作、Undo／Redo、保存動作は変更していない。
変更前後の配置計画のfingerprintは両方とも以下だった。

```text
174FC28A8E4D42CD61F0E717B9D656A0B2D4CC704A4B358232B67F5DFFFF6766
```

## ベンチマーク

Windows、実Excel、Releaseビルド、.NET SDK 9.0.313。
製品のglobal.json（9.0.304以上の互換.NET 9 SDKを許可）は変更していない。
既存のExcelとは別プロセスに作成した一時ブックで直列計測した。

共通条件はA1:AF722、240行×3 CASE、NEW／OLD、60図形。
図形は矩形で、管理画像と同じ名前・メタデータの30個と通常名の30個を使用する。
この計測はShape情報を読む処理のもので、PNGのデコードや画像編集の速度は含まない。
配置計画には120×30ポイントの画像を指定し、CASE 1-1のD5から配置されることを照合する。

各処理は1回のウォームアップ後、5回計測した中央値。
修正前はコミット0711fddの製品コードに計測テストだけを追加した。
修正後は2回計測したため、表には各回の中央値の範囲を記載する。

| 処理 | 修正前 | 修正後（各回の中央値の範囲） | 時間短縮 |
| --- | ---: | ---: | ---: |
| 配置先解析全体（Analyze） | 1,521.49 ms | 1,018.43～1,272.59 ms | 16.4～33.1% |
| CASE移動用スナップショット | 792.29 ms | 498.43～638.12 ms | 19.5～37.1% |
| 60図形の情報取得 | 682.85 ms | 419.08～518.95 ms | 24.0～38.6% |

CASE移動の計測は情報取得部分であり、Excelの画面移動は含まない。
キャプチャから編集、貼り付け完了までの総時間の短縮率ではない。
実際の効果は図形数、Excelの応答時間、PCの負荷によって変わる。
修正後初回は値の照合と計測が完了した後、試験用Excelの終了に失敗した。
試験側で共有Application参照を全解放していた箇所を1参照の解放に直し、
2回目は終了・一時ファイル回収まで含めて成功した。

生データは以下のTRXの標準出力に保存されている（artifactsはGit管理外）。

- `artifacts/test-results/placement-perf-baseline.trx`
- `artifacts/test-results/placement-perf-after.trx`
- `artifacts/test-results/performance-verified.trx`
- `artifacts/test-results/row-height-final.trx`

## 検証

- 通常テスト109件成功。Address非対応、形式不一致、相対参照、複数セル、行列上限超過のfallbackを含む。
- 実Excelテスト4件成功。画像配置、行挿入・削除、Undo／Redo、保護・接続再検証、既存の罫線・列幅テストを含む。
- 60図形すべての座標と管理画像判定を従来の個別取得と照合。
- A1／R1C1表示、A1・AAA1024・XFD1048576、図形移動とメタデータ変更後の再取得を確認。
- 生成したExcelプロセスは終了し、元から開いていたExcelプロセスは維持された。
- 参照用ブックは開かず、一時生成ブックのみを使用した。

参照ハッシュ検証スクリプトは、既存の保存済み基準と`escape/BetaEvidenceGenerator.bas`・
`MacroTest.xlsm`の2件が一致しなかった。今回これらのファイルは変更しておらず、
基準値の書き換えもしていない。残る4件は基準と一致した。

## 行高の正確性修正

行高の読み取りを検討する過程で、3～242行を15ポイント、123行だけ30ポイントにすると、
既存のReadRowHeightsが123行も15として返すケースを確認した。
`artifacts/test-results/row-perf-before.trx`に修正前の照合失敗を記録している。
範囲全体の高さを併用して誤った均一判定を防ぎ、混在範囲だけを二分探索するよう修正した。
実Excel上で240行を読み、個別のセル高さと全行を照合した結果は次の通り。

| 行高構成 | 修正後中央値 |
| --- | ---: |
| 全行15ポイント | 3.48 ms |
| 1行だけ30ポイント | 46.01 ms |
| 後半120行が30ポイント | 8.96 ms |
| 1行だけ非表示 | 49.29 ms |
| 全行非表示 | 1.62 ms |

全ケースで個別取得したExcelの実値と一致した。高さが細かく交互に異なる場合は分割数が増えるが、
結果の正確性を優先し、全行を常時個別取得する方式には戻していない。

## 再現コマンド

```powershell
dotnet test tests/EvidenceCrafter.Tests/EvidenceCrafter.Tests.csproj -c Release --filter 'FullyQualifiedName~PlacementAnalysis_WithReal' --logger 'trx;LogFileName=placement-perf.trx' --results-directory artifacts/test-results
dotnet test tests/EvidenceCrafter.Tests/EvidenceCrafter.Tests.csproj -c Release --filter 'FullyQualifiedName~RowHeightReads_WithReal'
dotnet test tests/EvidenceCrafter.Tests/EvidenceCrafter.Tests.csproj -c Release --filter 'TestCategory=ExcelIntegration'
.\bootstrap.ps1 -Publish
```
