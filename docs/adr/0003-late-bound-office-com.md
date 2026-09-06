# ADR 0003: Excel discoveryはOffice PIAへ実行時依存しない

- 状態: Proposed
- 日付: 2026-08-31

## Decision

ROTから取得したExcel COMオブジェクトは、`Microsoft.Office.Interop.Excel`の型へcastせず、adapter内部の遅延バインディングで読み取る。

## Context

型付きInteropパッケージは開発ホストのOffice 2019 x86でも実行時に `office.dll` を解決できず、Workbook列挙前に失敗した。単一EXE配布で利用者PCのPIA/GAC構成へ依存することは避けたい。

## Consequences

- App/CoreはOffice PIAを必要としない。
- COM property名はExcel外部契約としてadapterに閉じ込める。
- 書き込みPhaseへ進む前に、遅延バインディングの型変換と例外分類を結合テストで追加する。
