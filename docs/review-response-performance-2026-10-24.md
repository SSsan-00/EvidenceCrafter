# セル状態照合の通信削減（2026-10-24）

CASE末尾整理と行操作Undo／Redoの `CaptureRowFingerprint` を改善した。削除・復元の可否、照合対象とhashの順序、画像の配置計画、UI操作は変更しない。

## 変更

- 横・縦配置、フォント名・サイズ・太字・斜体を対象範囲から一度だけ取得する。想定する型の共通値が得られた項目だけセル間で共用し、混在・null・DBNull・取得失敗では従来のセル別取得を使う。既存のNumberFormat／WrapTextも同じ判定に統一した。
- 範囲内の条件付き書式が0件と確認できた場合、各セルには従来と同じ0を記録する。条件付き書式が存在する場合や取得失敗時はセルごとに読む。
- 共通値は1回の照合の中だけで使用する。次の操作では再取得する。色・塗りつぶし・Style・入力規則の読取、全Shapeのライブ安全確認、native退避、退避後の再確認は維持した。
- 公開API、設定、依存パッケージ、保存形式に変更はない。新しいCOM参照も取得したスコープで解放する。

対象範囲全体の条件付き書式を返すことは [Range.FormatConditionsの公式仕様](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formatconditions) を確認し、混在書式の具体的な読取結果は実Excelで旧実装と比較した。

## 検証

- 変更前の通常テスト: 173/173合格。
- 変更後の通常テスト: 174/174合格。共通値がfalseの場合、型違い・null・DBNull・COMエラー時のセル別フォールバックと既存の失敗表現を確認した。
- 追加の実Excelテスト `RowFingerprints_WithMixedFormatting_MatchCellReads`: 合格。独立した旧セル別readerとSHA-256を比較した。フォント名・サイズ・太字・斜体・色・横縦配置の単一セル変更と復元、セル内の一部分だけ異なる書式、条件付き書式、入力規則、結合セル、開始列Zを確認した。
- 既存の実Excel回帰4件: すべて合格（スキップなし）。`DiscoverAndFocus_WithRealTemporaryWorkbook_VerifiesIdentityAndActiveCell`、`SameSideBackfill_AlignsActualTopsAcrossCases`、`RowMutationOptimizations_PreserveSafetyAndFingerprints`、`AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard`。行削除・復元・Undo／Redo、外部編集の検出、退避中の変更、故障注入時の退避保持、自動配置後の移動を確認した。
- Releaseビルドとwin-x64単一EXEの発行が成功。警告・エラーなし。GUIの手動受入と実Excelテスト全件の再実行は行っていない。

## 性能比較

比較元は開始時commit `3038e4bf99b73e3a6064579ce02ac14b56a4ebe4` のReleaseビルド。変更前バイナリは `artifacts/response-performance-baseline` にコピーして保持した。.NET SDK 9.0.313、Windows上の実Excelで既存の `SameSidePerformanceChecks` を使った。測定コードのGit blobは変更前後とも `1f1d6af0ee2d434ecabfdc2d5506ca6f6a1d8f14`。

各条件はウォームアップ1回と計測3回の中央値。毎回専用Excelプロセスと一時ブックを作り、起動・fixture作成・終了を計測から除外した。総時間はOLD配置確定＋末尾整理＋次CASE移動で、プレビューやUI描画の時間を含まない。工程別中央値の和は総時間中央値と一致しない。

表の図形数は追加する対象CASE外の図形数であり、配置するNEW／OLDの管理画像は別に存在する。

| 条件 | 状態照合 前→後 | 末尾整理 前→後 | 総時間 前→後 | 総時間短縮率 |
| --- | ---: | ---: | ---: | ---: |
| 前方・0図形・均一書式・少数行削除 | 1.632→1.118秒 | 2.858→2.296秒 | 3.611→2.944秒 | 18.5% |
| 前方・200図形・均一書式・少数行削除 | 1.724→1.200秒 | 4.001→3.450秒 | 5.288→4.744秒 | 10.3% |
| 前方・0図形・混在書式・多数行削除 | 22.122→14.971秒 | 23.360→16.191秒 | 24.009→16.826秒 | 29.9% |

均一書式の条件では、追加した配置とフォント6項目は各1回の範囲取得、条件付き書式は0件の判定1回を確認した。セル数をNとすると、この部分の通信は従来の8N回から9回になる（範囲Font／FormatConditionsの取得を含む）。検証用の工程通知は0msの回数イベントであり、通信自体の所要時間を表すものではない。

多数行のfixtureでは照合対象が1,312セル。NumberFormat／WrapTextが混在し、従来どおりそれぞれ1,312回のセル別読取を実行している。全条件で状態照合は約30〜32%短縮した。前後計24回の専用プロセス試行（ウォームアップを含む）が成功し、失敗や遅い試行を除外していない。

実測はこの環境とfixtureに限定される。外部Excelの応答やPC負荷、書式の混在で効果は変わる。全項目が混在する場合はセル別取得を維持するため、同じ高速化率は保証しない。

原記録はGit管理外の `artifacts/response-performance-results` に保存する。

再計測では、Releaseビルド後に次の設定で既存テストを実行する。多数行条件は `COUNTS=0`、`FORMAT=mixed`、`DELETION=many` とする。

```powershell
$env:EVIDENCECRAFTER_PERF_COUNTS = '0,200'
$env:EVIDENCECRAFTER_PERF_SAMPLES = '3'
$env:EVIDENCECRAFTER_PERF_WARMUPS = '1'
$env:EVIDENCECRAFTER_PERF_POSITION = 'front'
$env:EVIDENCECRAFTER_PERF_FORMAT = 'uniform'
$env:EVIDENCECRAFTER_PERF_DELETION = 'few'
dotnet test tests/EvidenceCrafter.Tests -c Release --no-build --filter 'FullyQualifiedName~SameSidePerformance_WithGrowingShapeCounts'
```

## 実行用EXE

- 出力: `artifacts/response-performance-publish/win-x64/EvidenceCrafter.exe`
- SHA-256: `AE8D7348A7F37975E1CF275F29B776CE9A96C5E37DDEC8753645FFE27E978DC0`
- EXEとSHA-256 sidecarの2ファイルだけであること、およびハッシュ一致を確認した。

## 参照ベースライン

最初の `scripts/Test-ReferenceBaseline.ps1` の実行では、参照元のxlsx 3冊は一致したが、以下3件は記録済みハッシュと一致しなかった。いずれも作業開始（2026-10-24）より前の更新日時であり、参照元は変更していない。

その後のユーザー依頼により、マクロの配置定数、運用ガイド、マクロブックのZIP構造・VBAプロジェクト・テンプレートシートを確認した上で、この3件の現在のハッシュを新しい期待値として採用した。[参照ベースライン](reference-baseline.md) と検査スクリプトを同時に更新した。これは参照基準の更新であり、元の2026-09-06版との内容同一性やマクロの動作を保証するものではない。

更新後のハッシュ検査は全6件が一致した。

| ファイル | 最終更新日時（ローカル） | 現在のSHA-256 |
| --- | --- | --- |
| escape/BetaEvidenceGenerator.bas | 2026-09-08 16:06:53 | 7BC85DF208FA9A220714F9D40AE5FFCAF720F58CF8525C998ADC6E114AE11B51 |
| MacroModulesGuide.md | 2026-10-07 15:19:10 | CA8FE73561F6782904AD14481708E4A0A70BDC0592803C3064F46BAC726B0CC3 |
| MacroTest.xlsm | 2026-10-01 05:50:15 | 36A7C04CD0D6A50200A07728F691DF5DDE1566C3EC324EAC7E20AFE6A35EBAC8 |
