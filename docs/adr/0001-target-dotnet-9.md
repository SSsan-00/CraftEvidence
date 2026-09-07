# ADR 0001: .NET 9を採用する

- 状態: Accepted
- 日付: 2026-08-31

## Decision

EvidenceCrafterのtarget frameworkを `net9.0-windows` とする。

## Context

利用環境の要件として.NET 9系のSDKで開発・検証できることが必要になった。配布物は自己完結型とし、利用端末に導入済みのランタイムには依存しない。

## Consequences

- .NET 9対応Windowsを正式サポート対象とする。
- 古いWindows/Officeはbest-effort互換試験として分離する。
- self-contained single-fileにより利用者のruntime導入を不要にする。
