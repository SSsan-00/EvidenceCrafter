# EvidenceCrafter テスト方針

2026-09-11のキャプチャ・テキスト編集・最前面機能の検証結果と手動受入手順は
[実装レビュー](review-capture-editor-window.md) を参照。

2026-10-01のNEW先行・OLD後追い配置と性能測定は
[改修検証記録](review-same-side-backfill-2026-10-01.md) を参照。

CASE末尾整理の追加高速化・復旧用退避・故障注入の結果は
[追加改修検証記録](review-excel-com-performance-2026-10-01.md) を参照。

## 通常テスト

`EvidenceCrafter.Tests` はExcelなしで動くCoreテストを既定とする。

```powershell
dotnet test tests\EvidenceCrafter.Tests\EvidenceCrafter.Tests.csproj --filter 'TestCategory!=ExcelIntegration'
```

## Excel結合テスト

Officeを起動するテストには `[TestCategory("ExcelIntegration")]` を付ける。通常CIでは実行せず、Office導入済みの明示環境だけで実行する。

現行の実機テストは、同一Excelプロセスに一時Workbookを2冊作成し、Workbook別HWND、通常編集後の接続ID維持、非アクティブ側のセル選択、`Application.Goto`、`EnableEvents`復元を確認する。さらに一時PNGの手動／Case自動配置、必要行挿入、管理画像のExport・差し替え・削除・元geometry復元、保護Sheetでの拒否を確認する。行操作ではActiveCell上への挿入、live safety snapshotに基づくCase末尾削除、Undo/Redo、編集済み挿入行のUndo拒否を確認する。Close取消、監視token再発行、確定Close後の旧identity拒否、生成Excel PID終了と一時ファイル回収も同一シナリオで検証する。Windows PowerShell 5.1でも参照ハッシュ検証を再現できるよう、スクリプトはUTF-8 BOMで保存する。

`SameSideBackfill` は4CASEの配置・末尾整理・同じSideの次CASE移動後の実Top、外部編集とRedo、エビデンス列外の削除保護を確認する。`RowMutationOptimizations` は旧fingerprintとのhash同値、Shape境界と両Side判定、故障時の退避保持、退避中の変更検出、行削除Undo→Redoの内容保全を実Excelで検証する。

上記fixtureには印刷範囲を設定し、`RowMutationOptimizations` には後続CASEを参照する名前定義も追加する。名前定義を読む・復元する `Names.Item` はメソッドとして呼ぶ。行削除Undo／Redo後は、Excelから直接読み取った名前定義と退避記録を比較する。

`AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard` は一時ブックと非表示のMainFormを使い、実際の自動配置ハンドラーをメッセージ処理付きで実行する。印刷範囲付きCASEの正常整理、別シートの数式による行未変更での整理中止、同じSideの次CASE移動、配置履歴・画像・CASE境界の保全、結果不明の行変更が報告された後の操作停止を確認する。設定・ログは専用一時ディレクトリへ保存する。

このfixtureでは先行行を高くして、通常の安全確認セル数の範囲内で予定Top=937483.25ptを作る。OLD後追い・末尾整理後・実際の画面Undo／Redo後の横並びと次CASE移動を確認する。

`InsertedImages_VerifyGeometryAndPreserveFailedRollback` は既存画像とセル内容があるシートで、座標0・下方の小数座標・非表示行・非表示列・geometry不一致・取り消し失敗を確認する。後続CASEのmarkerと先に配置済みの画像・追加行が保持されることも検証する。行を隠すfixtureは `EntireRow` を使用し、実Hidden・Heightと画像列が表示されていることを確認する。

Topの丸め回帰は50,000／100,000／200,000行の6条件で、直接読んだExcel実Topが調査時の実値と一致することを検証する。遠方列のLeft、削除・再配置、Single刻み幅の境界、隣接値を越えるずれ・追加許容上限超過の拒否も確認する。

挿入画像のgeometry検証・取り消し未確認などの失敗では、予定／実Left・Top・Width・Height、処理工程、Excel版、Side、セル座標、元サイズ・倍率、行／列の非表示状態を `%LOCALAPPDATA%\EvidenceCrafter\logs\diagnostic.jsonl` に記録する。この失敗診断は通常の診断ログ設定が無効でも記録し、セル内容・画像データ・Workbookのパスは含めない。画像の取り消しを確認できない場合はアプリの変更操作を停止し、復旧用PNGのパスをステータスへ表示する。

preview.35からは最寄りSingleのLeft／Top、隣接Singleの追加許容上限、Top位置検証の結果、アプリInformationalVersionも記録する。既存の予定値差0.05pt／最寄りSingle差0.05ptに加えて、最寄り値から0.25pt以内の直前／直後Singleと実値が完全一致する場合だけ追加で許可する。画像同士の実Top差、サイズ検証、外部編集判定は従来の0.05ptを維持する。

preview.36の `InsertedImages_CorrectReportedSizeDriftOnceAndVerifyFinalGeometry` は、新規画像に報告値745.2000122070312ptを注入し、幅・高さ、原寸・縮小、挿入直後・属性設定後、低い・大きい座標で予定745.5ptへ戻せることを独立COM読取で確認する。各辺の差0.5pt以内で位置が妥当な新規画像だけを一度補正し、補正後もサイズの許容差0.05ptを維持する。通常成功時は補正しない。補正の例外・再度のずれ・位置／名前の変更・削除失敗、自動配置・後追い・Redo、複数画像の途中失敗による画像・行の補償も確認する。失敗診断には補正前の実geometryと `sizeCorrectionAttempted` を追加する。自然発生の調査結果と制限は[検証記録](review-size-verification-2026-10-05.md)を参照。

preview.37では新規画像の最終サイズ許容差を各辺1ptへ変更した。0.5pt以内の差は従来どおり一度補正し、それでも1pt以内の差が残る場合はExcel実geometryで成功とする。既存の実Excel試験へ幅／高さの±1ptの成功、±1.25ptの拒否、補正後の約−0.30pt差の成功、配置可能幅の超過拒否を追加した。受入後の実サイズを使うUndoと、0.30ptの外部サイズ編集の拒否も確認する。自動配置・後追い・Redoは補正後にも差を残した状態で確認し、複数画像の補償試験は1ptを超える差で失敗させる。診断に `insertedImageSizeTolerancePoints` を記録する。位置・参照画像の検証を含む詳細は[1pt許容の検証記録](review-size-tolerance-1pt-2026-10-05.md)を参照。

preview.38では上記preview.35〜37の配置比較・サイズ補正を置き換え、予定値と実値を各0.1pt単位へ丸めて差1pt以内を受け入れる。位置・サイズ、NEW／OLDの上端合わせ、参照画像の拡縮、配置可能幅、行高さ不足の比較を共通化した。`InsertedImages_AllowRoundedGeometryAndPreserveFailureRecovery` は低い／大きい座標で4辺の±1pt・0.1pt丸め境界、幅上限、自動配置・後追い・Redo、補償を実Excelで確認する。履歴は丸めない実geometryを保存し、履歴対象の外部編集検知は従来の0.05ptを維持する。診断には丸めた座標と `comparisonTolerancePoints`／`comparisonPrecisionPoints` を記録する。詳細は[配置全体の1pt許容の検証記録](review-rounded-placement-1pt-2026-10-05.md)を参照。

`SameSidePerformance` は対象CASE外の0/50/200/500図形で、専用Excelプロセスをサンプルごとに作り、プレビュー・配置・末尾整理・次CASE移動をウォームアップ1回＋計測10回測る。起動・fixture生成・終了を時間から除外する。全体中央値と最大値、工程時間・読取回数をJSONLに記録できる。

```powershell
New-Item -ItemType Directory -Path artifacts -Force | Out-Null
$env:EVIDENCECRAFTER_PERF_LOG = Join-Path (Get-Location) 'artifacts\performance.jsonl'
dotnet test tests\EvidenceCrafter.Tests -c Release --no-build --filter 'FullyQualifiedName~SameSidePerformance'
```

条件変更用の環境変数は`EVIDENCECRAFTER_PERF_COUNTS`（既定`0,50,200,500`）、`EVIDENCECRAFTER_PERF_SAMPLES`（10）、`EVIDENCECRAFTER_PERF_WARMUPS`（1）、`EVIDENCECRAFTER_PERF_POSITION`（`front`／`middle`／`back`）、`EVIDENCECRAFTER_PERF_FORMAT`（`uniform`／`mixed`）、`EVIDENCECRAFTER_PERF_DELETION`（`few`／`many`／`none`）。JSONLは追記するため、比較する版ごとに別ファイルを指定する。Excelを使うテストは同時に実行せず、性能計測中はビルドもしない。

参照Workbookを使う場合は次を必須とする。

1. SHA-256をベースラインと照合する。
2. `%TEMP%\EvidenceCrafter.Tests\<GUID>` へコピーする。
3. コピーだけをExcelで開く。
4. Excelを閉じ、一時コピーを削除する。
5. 参照元SHA-256を再確認する。

## Review sliceのテスト対象

- Case Anchorの選択
- 次AnchorからのCase終端
- 最終CaseのUsedRange終端と自動行削除停止
- ヘッダーによるNewのみ／New・Old境界、罫線あり／なしの同一配置結果
- 画像の縮小と非拡大
- 画面pixelの96 DPI論理サイズ換算
- 2行非画像帯と4行末尾余白
- 虫食い候補と末尾fallback
- 入力値のguard
- Excel最大行・最大列境界と整数overflow guard
- 非最終Caseを含む論理Evidence終端の整合
- Workbook close/reopenの接続世代とExcelイベント復元
- 明示的な行挿入と、live safety snapshotによるCase末尾行削除
- 個別Evidence実例の30行Case、C:Q／R:AF境界、書式末尾を含むUsedRange
- ROT内のIDispatch非対応COMオブジェクトの読み飛ばしと探索継続

## 完了条件

preview.40では負の縦オフセット−1pt／−0.3ptを0ptとして計画・挿入する。`PlacementPlannerTests` と `AutomaticPlacementServiceTests` で必要行数と後追い解析、非有限値・許容範囲外の拒否を確認し、`InsertedImages_AllowRoundedGeometryAndPreserveFailureRecovery` で実Excelへの負オフセット指定・削除・範囲外の拒否を確認する。

preview.39の `SplitCaseNumbering_ResolvesPlacesNavigatesAndReplays` はA5／B6の別行番号、大番号の変更、従来の同じ行番号、A列行またはB列行の直前にあるSide見出し、CASE開始・終端、指定CASEへの配置、削除・再配置、前後CASE移動、重複CASEの拒否を専用の一時Excelブックで確認する。

- Release buildでwarning 0
- Excel非依存テストが全件成功
- `bootstrap.ps1 -Publish` が単一EXEを生成
- 参照ベースラインのハッシュ不変
