# ADR 0004: preview.10では行Mutationを明示操作に限定する

- Status: Accepted
- Date: 2026-09-05

## Context

画像配置計画には行挿入候補が含まれるが、Case境界や既存コンテンツを誤認したまま自動実行すると利用者Workbookのデータを破壊し得る。一方、動作検証には行挿入と行削除をツールから実行できる最小経路が必要である。

## Decision

preview.10では次の2操作だけを、確認ダイアログを伴う明示操作として提供する。

- 選択したWorkbook／SheetのActiveCell上へ、指定行数を挿入する。
- 指定したCase範囲についてlive safety snapshotを取得し、安全な末尾連続行だけを削除する。

実行直前にWorkbook identity、接続token、ReadOnly、Sheet保護を再検証する。削除安全性はセル値、数式、結合、コメント／ノート、ハイパーリンク、Shapeを対象とし、判定不能は削除不可とする。Excelのイベント、画面更新、警告表示は操作後に元の値へ戻す。

## Consequences

利用者は行操作を単独で検証できる。画像配置・行挿入・安全な行削除は、同一アプリ／Workbook接続世代内のLIFO履歴でUndo/Redoできる。挿入行に内容が追加された場合はUndoを拒否する。画像配置からの自動行調整、処理途中例外の複数ステップ補償、保存は後続とする。
