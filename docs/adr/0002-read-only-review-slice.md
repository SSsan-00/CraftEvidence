# ADR 0002: 初回レビュー版を読み取り専用にする

- 状態: Proposed
- 日付: 2026-08-31

## Decision

0.1.0-previewではWorkbook列挙、Core解析、Clipboardプレビューまでを提供し、Excel書き込みを接続しない。

## Context

Case/Side境界の誤判定や行全体削除は利用者データを破壊し得る。安全停止条件とCOM所有権のレビュー前に書き込み機能を公開すべきではない。

## Consequences

- UI/解析/配布を早期にレビューできる。
- 実用的な貼り付けは次の承認済みPhaseで実装する。
- Clipboard検知だけではExcelへ副作用を発生させない。
