# 性能監査と列単位の書式読取（2026-11-08）

前回の改善を含む `535e174` を基準として、Excel配置・末尾整理・履歴操作、CASE移動、Clipboard画像処理、画像編集の主要経路を調べた。

## 監査結果と変更範囲

| 経路 | 確認した実装・ボトルネック | 対応 |
| --- | --- | --- |
| 末尾整理・行Undo／Redo | `CaptureRowFingerprint` は全範囲で書式が混在するとNumberFormat／WrapTextを全セルから再取得する。列内では同じ書式でも通信回数が行数×列数に比例する | 列内の共通値を取得し、確認できた項目のみ同じ照合内で再利用 |
| 図形の安全確認 | 対象外を含む全図形のライブ境界走査が図形数に比例する | 削除の安全性に必要なため維持。古い図形情報の再利用は行わない |
| CASE移動・選択 | 既に図形なしの軽量Snapshotを使用。行高・列幅にも範囲読取がある | 今回は変更なし |
| Clipboard・画像編集 | Clipboardの世代番号で重複を抑制し、PNGエンコードはTask.Run。編集の描画は既存画像と注釈を使用し、モザイクはGDI+の縮小・拡大 | コード確認のみ。今回のExcel計測を画像編集の性能結果として扱わない |

変更先は `ExcelRowMutationService` の状態照合共通処理。行削除後、Undo前、復元後、Redo前の4呼出箇所に適用される。公開API・設定・依存パッケージ・hashの内容と順序は変更しない。

全範囲の共通値を優先し、必要な場合だけ照合対象行に限定した各列の範囲を読む。列内混在、null、DBNull、取得失敗ではセル別読取に戻る。単一行では列読取を追加しない。値は操作をまたいで保持せず、COM参照は各列の読取後に解放する。

2項目が列ごとに均一なR行×C列では、この部分の通信は2RC回から3C回（列Range取得を含む）になる。全列内で混在する場合は従来のセル別取得に最大3C回が加わるため、その条件で高速化を保証しない。

## 検証方法

- 変更前後の通常テストと、旧セル別readerを独立した比較対象にした実ExcelのSHA-256一致確認。
- 列間の混在、列内の混在、照合範囲外の異なる書式、全列内混在、単一行、結合セルを確認する。取得失敗と不要な項目を再取得しないことは単体テストでも確認する。
- 性能測定は既存の `SameSidePerformanceChecks` を変更せず使用。比較元のReleaseバイナリは `artifacts/performance-2026-11-08-baseline`、原記録は `artifacts/performance-2026-11-08` に保存する。
- 各条件はウォームアップ1回＋計測3回。専用のExcelプロセスと一時ブックを毎回作成し、起動・fixture作成・終了を測定から除く。総時間はOLD配置確定＋末尾整理＋次CASE移動で、プレビューとUI描画は含まない。図形数は対象CASE外に追加する図形の数。

.NET SDKは9.0.313。測定コードのGit blobは前後とも `1f1d6af0ee2d434ecabfdc2d5506ca6f6a1d8f14`。

再現用の環境変数は `EVIDENCECRAFTER_PERF_SAMPLES=3`、`EVIDENCECRAFTER_PERF_WARMUPS=1`、`EVIDENCECRAFTER_PERF_POSITION=front`。均一書式は `COUNTS=0,200`／`FORMAT=uniform`／`DELETION=few`、列別書式は `COUNTS=0`／`FORMAT=mixed`／`DELETION=many`（各変数名に `EVIDENCECRAFTER_PERF_` を付ける）。

```powershell
dotnet test tests/EvidenceCrafter.Tests -c Release --no-build --filter 'FullyQualifiedName~SameSidePerformance_WithGrowingShapeCounts'
```

## 検証結果

- 通常テスト: 変更前174/174件、変更後177/177件合格。
- 実Excel `RowFingerprints_WithMixedFormatting_MatchCellReads`: 合格。旧セル別読取とのhash同値を維持し、列単位の共通書式を読むケースでは4列×2項目の範囲取得、1列だけ混在するケースではその列の2セル×2項目だけがフォールバック、全列内混在では8セル×2項目がフォールバックすることを確認した。
- 取得に失敗する列、null／DBNull、片方の項目だけ取得失敗する場合、全体で既知の項目を再取得しない条件は単体テストで確認した。
- 既存の実Excel回帰4件（DiscoverAndFocus、SameSideBackfill、RowMutationOptimizations、AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard）も全件合格、スキップなし。行削除・復元・Undo／Redo、外部編集検出、故障注入時の退避保持、自動配置後のCASE移動を確認した。上記hash比較を含め、関連実Excelテストは計5件合格。
- 参照元のハッシュは作業前後とも全6件一致。Releaseビルドとwin-x64単一EXEの発行は成功し、警告・エラーなし。

## ベンチマーク

単位は秒、各3回の中央値。工程別中央値の和は総時間中央値と一致しない。

| 条件 | 状態照合 前→後 | 末尾整理 前→後 | 総時間 前→後 | 総時間の変化 |
| --- | ---: | ---: | ---: | ---: |
| 均一書式・0図形・少数行 | 1.678→1.746 | 3.396→3.487 | 4.401→4.612 | 4.8%増加 |
| 均一書式・200図形・少数行 | 1.727→1.672 | 4.990→5.365 | 7.162→7.449 | 4.0%増加 |
| 列別書式・0図形・多数行 | 24.106→19.998 | 25.907→21.829 | 27.064→22.896 | 15.4%短縮 |

列別書式のfixtureは41行×32列＝1,312セル。測定3回すべてで、NumberFormat／WrapTextのセル別取得が各1,312回から0回へ、範囲取得が各32回になった。列Rangeの取得32回を含めても、この部分は2,624回から96回へ減る。状態照合は17.0%、末尾整理は15.7%短縮した。前後合計24回の専用Excelプロセス試行（ウォームアップを含む）はすべて成功し、失敗や遅い試行を除外していない。

均一書式では追加の列読取はなく、Excelへの書式読取回数は前後で同じ。未変更のプレビュー経路にも−6.6%〜+9.4%の差があった。この3回の測定だけで負荷変動と回帰を統計的に区別していないため、均一書式の高速化や全条件での非悪化は主張しない。

全図形のライブ境界確認、入力規則・色などのセル別照合、native退避の負荷は残る。今回の変更で全ボトルネックを解消したとは扱わない。GUIの手動操作・画像編集の時間計測、全実Excelシナリオの再実行は対象外。

## 実行用EXE

- `artifacts/performance-2026-11-08-publish/win-x64/EvidenceCrafter.exe`
- SHA-256: `3ACC36FCBD93D9D4649DFBF8139D409E292F5FC4F4D86010002F1E72767E2F1D`
- 出力がEXEとSHA-256 sidecarの2ファイルだけであること、ハッシュ一致を確認済み。
