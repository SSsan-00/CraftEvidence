# ADR 0001: .NET 10 LTSを採用する

- 状態: Proposed
- 日付: 2026-08-31

## Decision

EvidenceCrafterのtarget frameworkを `net10.0-windows` とする。

## Context

共有仕様は.NET 9を前提としていたが、2026-11-10にサポート終了予定である。新規製品の実装開始時点で保守期間が短すぎる。

## Consequences

- .NET 10対応Windowsを正式サポート対象とする。
- 古いWindows/Officeはbest-effort互換試験として分離する。
- self-contained single-fileにより利用者のruntime導入を不要にする。
