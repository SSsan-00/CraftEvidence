# Case&Evidence 参照ベースライン

- 確認日: 2026-09-06
- 参照ルート: `C:\work\Macro\Case&Evidence`
- 取扱い: 読み取り専用

## 仕様源の優先順位

1. `escape/BetaEvidenceGenerator.bas`
2. `MacroModulesGuide.md`
3. `MacroSuiteSingleFileInstaller.bas` 内の埋め込みモジュール
4. `MacroTest.xlsm` と生成済み `.xlsx`

`MacroTest.xlsm` と生成済みWorkbookには現行コードとの版差があり得るため、生成済みWorkbookはレイアウト形状のfixtureとしてのみ使う。2026-08-31の差し替え後は、共通・個別エビデンスと対応するテストケースWorkbookの3冊を一組の参照資料として扱う。

## Evidenceレイアウト契約

- 既定Case開始行: 3
- 既定slot height: 50、ただし設定変更可能
- Case情報: Case開始行のA/B列
- New開始: C列
- New既定右端: Q列
- Old既定開始: R列
- 既定全体右端: AF列
- New列数は可変
- `Both`ではOld列数も追従
- 横罫線、縦境界線、Old見出しは無効化可能
- 現行の共通・個別ブックはNew `C:Q` / Old `R:AF`
- 共通ブックにはCase下端より後ろに書式だけが残る行があるため、raw UsedRange最終行を実データ終端と同一視しない

## 基準ハッシュ

| File | SHA-256 |
|---|---|
| `escape/BetaEvidenceGenerator.bas` | `0C1B0F5D8108D74D738BB9BE9C34893A37F407B5634FE9BCF199BE57B74ECCF5` |
| `MacroModulesGuide.md` | `25319391FF27B63412B78CF6A2ED66CF2FDC18003436FBE8F706D2FA201D9B2B` |
| `MacroTest.xlsm` | `72618E3D34ED63EB716FCE7508A1F2D5CE93A6FBA7D03A46A6A346CBA45E4A3F` |
| `S00-000-00TESTX_【共通】02TEST2_単体テストエビデンス_初期開発.xlsx` | `D2FAC5E708F461AC030D543D81EE88779D347F7C35369BB2E9AB430D038DEF8D` |
| `S00-000-00TESTX_【個別】02TEST2_単体テストエビデンス_初期開発.xlsx` | `F105816F85476CD58399A24A39706F1D172BA9E380EC6D512749C18F52DFDB52` |
| `S00-000-00TESTX_単体テストケース_初期開発_003.xlsx` | `C10C6D120FAF51E1FB1484EBEC7DE52F343BB8FB9265C76C8FCFC85D135FF666` |

## テスト利用規約

- 参照元を直接Excelで更新しない。
- 書き込み試験は一時ディレクトリへ複製してから実施する。
- fixtureは業務値を含まない匿名JSONをCraftEvidence側に置く。
- テスト前後で上記ハッシュを比較する。
