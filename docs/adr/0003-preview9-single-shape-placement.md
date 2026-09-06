# ADR 0003: preview.9で単一Shape配置を実機検証する

- 状態: Accepted for hands-on verification
- 日付: 2026-09-05

## Decision

preview.9では、ユーザーがPreviewで明示承認した場合に限り、検証済みWorkbookの指定Sheet（または`ActiveSheet`）をアクティブ化し、そのWorkbook側ActiveCellへ1枚の画像Shapeを配置する。SideはUIとShapeメタデータへ記録するが、列境界への自動移動はこの版では行わない。

配置前後にWorkbook identity、Window session token、`EnableEvents`、Shape名を検証し、配置後はROTから再接続して対象セルを選択する。前面化、自動保存、行追加・行削除、差し替え、Undoはこの版に含めない。

## Consequences

- 実ExcelでClipboardから配置・選択までの一連動作を確認できる。
- 行Mutationをまだ接続しないため、検証用の一時Workbookで安全に試せる。
- Case/Side解析と配置計画をMutationへ接続するPhase 5以降の受入条件は残る。
