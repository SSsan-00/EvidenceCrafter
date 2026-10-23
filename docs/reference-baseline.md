# Case&Evidence 参照ベースライン

- 確認日: 2026-10-24（生成済みxlsx 3冊の基準は2026-09-06から変更なし）
- 参照ルート: `C:\work\Macro\Case&Evidence`
- 取扱い: 読み取り専用

## 仕様源の優先順位

1. `escape/BetaEvidenceGenerator.bas`
2. `MacroModulesGuide.md`
3. `MacroSuiteSingleFileInstaller.bas` 内の埋め込みモジュール
4. `MacroTest.xlsm` と生成済み `.xlsx`

`MacroTest.xlsm` と生成済みWorkbookには現行コードとの版差があり得るため、生成済みWorkbookはレイアウト形状のfixtureとしてのみ使う。2026-09-06のユーザー差し替え後は、共通・個別エビデンスと対応するテストケースWorkbookの3冊を一組の参照資料として扱う。

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

2026-10-24のユーザー依頼により、更新済みのマクロ本体・運用ガイド・マクロブック3件を現在の参照基準として採用した。マクロのCase開始行3、slot height 50、A/B列、New C:Q／Old R:AF、見出し行2が上記契約と一致することを確認した。運用ガイドの対象モジュールを読み、MacroTest.xlsmの全ZIPエントリの読取、VBAプロジェクトとテンプレートシートA1／A1-1-1の存在を確認した。マクロの実行検証は行っていない。参照元は変更せず、検査スクリプトと本表の期待値のみ更新した。

| File | SHA-256 |
|---|---|
| `escape/BetaEvidenceGenerator.bas` | `7BC85DF208FA9A220714F9D40AE5FFCAF720F58CF8525C998ADC6E114AE11B51` |
| `MacroModulesGuide.md` | `CA8FE73561F6782904AD14481708E4A0A70BDC0592803C3064F46BAC726B0CC3` |
| `MacroTest.xlsm` | `36A7C04CD0D6A50200A07728F691DF5DDE1566C3EC324EAC7E20AFE6A35EBAC8` |
| `S00-000-00TESTX_【共通】02TEST2_単体テストエビデンス_初期開発.xlsx` | `D2FAC5E708F461AC030D543D81EE88779D347F7C35369BB2E9AB430D038DEF8D` |
| `S00-000-00TESTX_【個別】02TEST2_単体テストエビデンス_初期開発.xlsx` | `8C98F2E545DAFF407836E2842DEBE021B5CF24916FC71FFE99A67FD1755873CA` |
| `S00-000-00TESTX_単体テストケース_初期開発_003.xlsx` | `C10C6D120FAF51E1FB1484EBEC7DE52F343BB8FB9265C76C8FCFC85D135FF666` |

## テスト利用規約

- 参照元を直接Excelで更新しない。
- 書き込み試験は一時ディレクトリへ複製してから実施する。
- fixtureは業務値を含まない匿名JSONをEvidenceCrafter側に置く。
- テスト前後で上記ハッシュを比較する。
