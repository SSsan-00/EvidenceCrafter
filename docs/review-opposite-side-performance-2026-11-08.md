# 複数枚の先行配置後のレスポンス改善（2026-11-08）

## 原因と変更

同じCASEにNEWを4枚配置してから大きいOLDを1枚置く条件と、4つのCASEにNEWを1枚ずつ配置してから最初のCASEにOLDを置く条件を再現した。比較元は `9c944bb`。画像を置いた後に同期実行する末尾行整理、特にUndo用の `CaptureRowFingerprint` が支配的だった。同一CASEの変更前中央値では、総時間14.83秒に対して末尾整理11.53秒、行状態照合9.50秒だった。

従来はFontの名前・サイズ・太字・斜体が範囲内で共通と分かっていても、色を読むために全セルのFontを取得していた。Interiorも全セルで取得し、色とパターンを個別に読んでいた。

`ExcelRowMutationService` の同じ照合処理内でFont.Color、Interior.Color、Interior.Patternを範囲取得する。Fontの全照合項目、またはInteriorの全照合項目が既知なら、そのオブジェクトのセル別取得を省く。均一な通常書式では、この部分を5回×セル数のCOM読取から範囲全体で4回の追加読取へ置き換える。既存の値・数式・入力規則・Style・行情報・依存関係の照合は維持する。

入力規則なしの場合も、従来は各セルでValidationとType／Formula1／Formula2を読み、未設定を表す同じ3項目を構築していた。[SpecialCells](https://learn.microsoft.com/en-us/office/vba/api/excel.range.specialcells)の`xlCellTypeAllValidation`で範囲内に該当セルがないと確認できたとき、この4回×セル数の読取を範囲検索1回へ置き換える。比較するfingerprintの内容は変えない。

共通Styleの取得と既存の文字列化も範囲ごとに1回にまとめる。値・数式・すべての照合書式が既知で、入力規則と条件付き書式がない場合は、セルのCOMオブジェクト取得自体も省ける。列単位で既知のNumberFormat／WrapTextも既存キャッシュを使う。Style混在や1項目でも未確定の場合はセルを取得し、未知の項目だけ従来どおり読む。StyleのCOM参照も処理後に解放する。

「該当セルなし」はExcelのHRESULT `0x800A03EC`と、日本語または英語の明示的なno-matchメッセージが両方一致した場合だけ扱う。一般的な1004エラー、通信エラー、未知のメッセージ、nullの結果は個別照合へ戻す。1セルでも入力規則があれば全セルを従来どおり照合する。検索先はfingerprint対象範囲に限定し、範囲外の規則による誤判定も実Excelで確認する。

混在、null、DBNull、取得不能は従来のセル別読取へ戻る。実ExcelでInterior.Colorは混在色と黒の両方に0を返すことを確認したため、0を共通色として使用しない。Font.Colorのセル間混在、セル内の部分色変更、塗りつぶしパターン混在も旧readerとの照合対象にした。[Font.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.font.color)と[Interior.Color](https://learn.microsoft.com/en-us/office/vba/api/excel.interior.color)はVariant型のAPIであり、実際に確認できた型の共通値のみを使う。

呼出元は行削除後、Undo前、復元後、Redo前の4箇所。配置順・左右の対応付け・倍率と上端合わせ・行の追加削除・CASE移動・履歴の検証条件は変更しない。キャッシュは単一照合内だけで、外部編集をまたぐ再利用はしない。追加依存関係やユーザー設定はない。

## 再現方法

- Windows、.NET SDK 9.0.313、ローカルExcel。全試行は専用Excelプロセスと一時ブックを使用し、直列実行。
- 書式は均一、行高15pt、A:AFの32列、CASE間隔50行。対象外の追加図形は0個。
- 同一CASE: 120×80のNEWを4枚置いた後、180×240のOLDを1枚配置。参照画像の拡大と行追加を含む。
- 複数CASE: 4CASEに120×80のNEWを1枚ずつ置いた後、最初のCASEに同寸法のOLDを1枚配置。
- 変更前と書式読取改善段階は各条件ウォームアップ1回＋計測3回。10秒以下という追加目標を確認する最終版は、各条件ウォームアップ1回＋計測5回。起動・fixture作成・終了を除外。総時間は配置確定＋末尾整理＋次CASE移動。プレビューは別測定で、UI描画やClipboard取得は含まない。
- 変更前の製品DLLは `artifacts/opposite-side-2026-11-08-baseline` に保存。最終版と複数CASEの変更前測定に用いた`SameSidePerformanceChecks.cs`のGit blobは `b3a68b26053d5659e6489815e46233949a196088`。同一CASEの初回計測後に別CASE用fixture分岐と記録項目を追加したが、その条件で実行する配置と計測区間は同じ。
- 原記録は `artifacts/opposite-side-2026-11-08` のJSONLとTRX。失敗や遅い試行を選別して除外しない。

```powershell
$env:EVIDENCECRAFTER_PERF_COUNTS = '0'
$env:EVIDENCECRAFTER_PERF_SAMPLES = '5'
$env:EVIDENCECRAFTER_PERF_WARMUPS = '1'
$env:EVIDENCECRAFTER_PERF_POSITION = 'front'
$env:EVIDENCECRAFTER_PERF_FORMAT = 'uniform'
$env:EVIDENCECRAFTER_PERF_DELETION = 'many'
$env:EVIDENCECRAFTER_PERF_NEW_IMAGES = '4'
$env:EVIDENCECRAFTER_PERF_LARGE_OLD = '1' # 複数CASEでは0
$env:EVIDENCECRAFTER_PERF_SEPARATE_CASES = '0' # 複数CASEでは1
$env:EVIDENCECRAFTER_PERF_LOG = 'C:/work/EvidenceCrafter/artifacts/opposite-side-2026-11-08/new-run.jsonl'
dotnet test tests/EvidenceCrafter.Tests -c Release --no-build --filter 'FullyQualifiedName~SameSidePerformance'
```

## 測定と検証結果

単位は秒、変更前3回・最終版5回の中央値。工程別中央値の和は総時間中央値とは一致しない。最終版の原記録は`same-complete.jsonl`と`cases-complete.jsonl`。

| 条件 | 状態照合 前→後 | 末尾整理 前→後 | 総時間 前→後 | 総時間の短縮 |
| --- | ---: | ---: | ---: | ---: |
| 同じCASEにNEW4枚→大きいOLD1枚 | 9.502→0.616 | 11.529→2.493 | 14.827→5.755 | 61.2% |
| 4CASEにNEW各1枚→最初のCASEへOLD | 20.067→1.213 | 21.841→3.218 | 22.854→4.293 | 81.2% |

状態照合自体はそれぞれ93.5%、94.0%短縮。最終版の全計測でセル／Font／Interior／Validationの個別取得は0回。変更前と最終版の合計20試行（ウォームアップを含む）はすべて成功し、末尾行の削除も実行された。

最終版は両条件の計測各5回すべて10秒以下。最大値は同一CASEで8.170秒、複数CASEで6.191秒。プレビューを各試行に加えた最大値もそれぞれ8.484秒、6.583秒で、ウォームアップ各1回も10秒以下だった。書式読取だけを改善した中間段階の複数CASE中央値15.913秒からも、最終版では4.293秒へ短縮した。

総時間の各実測値（ウォームアップを除く）:

- 同一CASE・前: 16.411、14.563、14.827。後: 5.586、5.755、8.170、5.926、5.597。
- 複数CASE・前: 22.854、23.752、22.054。後: 5.208、6.191、4.159、4.293、4.179。

プレビューの中央値は同一CASEで0.426→0.314秒、複数CASEで0.260→0.268秒。未変更の工程にも変動があり、少数サンプルで全経路の高速化や非悪化を主張しない。

## 検証

- 通常テスト178件と、実Excelで旧セル別readerのSHA-256と比較する`RowFingerprints_WithMixedFormatting_MatchCellReads`が合格。
- 共通／混在Style、列別および列内混在の書式、フォント色・セル内の部分色、黒・白・色混在の塗りつぶし、パターン、条件付き書式、入力規則の追加・削除・入力メッセージのみ・対象範囲外、単一行、結合セルでhash同値を確認。
- 未知の1004エラー、通信エラー、未対応言語、nullの戻り値を「入力規則なし」と扱わないことを単体テストで確認。
- 最終コードで実Excel回帰5件（PairedImagesの位置合わせと衝突回避、SameSideBackfill、RowMutationOptimizations、AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard）が全件合格、スキップなし。左右上端、参照画像の拡縮、行追加・削除・復元、Undo／Redo、外部編集検出、セル内容の保護、自動CASE移動を確認した。hash比較を含め実Excelは6シナリオ、通常テストと合わせて184件合格。
- 参照資料のSHA-256は全6件一致。既存の3件の不一致は解消済みで、今回も参照ファイル自体は変更していない。
- 参照用xlsx3冊はZIP内のworksheet XMLを読み取り、入力規則の定義がないことを確認した。実ブック全体の性能測定を代替するものではない。

## 実行用EXE

- `artifacts/opposite-side-2026-11-08-publish/win-x64/EvidenceCrafter.exe`
- SHA-256: `6C037CA57179933A6C813819ACAEF459380916BF8F60E6F1049E949BDD854559`
- Releaseのwin-x64単一EXEを発行し、出力がEXEとSHA-256 sidecarの2ファイルだけであること、ハッシュ一致を確認した。

## 制限

ユーザーの実ブック・実Clipboard画像・手動GUI操作は未測定。ローカル測定であり、全環境で同じ短縮率や10秒以下を保証しない。書式が混在する範囲、入力規則が存在する範囲、日本語・英語以外のExcelのno-matchメッセージでは従来の個別照合が残る。全図形のライブ境界確認とnative退避も引き続き時間がかかる。安全確認やUndo用の退避を省く変更はしていない。
