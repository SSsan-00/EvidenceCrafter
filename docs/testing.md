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

`InsertedImages_VerifyGeometryAndPreserveFailedRollback` は既存画像とセル内容があるシートで、座標0・下方の小数座標・非表示行・非表示列・geometry不一致・取り消し失敗を確認する。後続CASEのmarkerと先に配置済みの画像・追加行が保持されることも検証する。行を隠すfixtureは `EntireRow` を使用し、実Hidden・Heightと画像列が表示されていることを確認する。

挿入画像のgeometry検証・取り消し未確認などの失敗では、予定／実Left・Top・Width・Height、処理工程、Excel版、Side、セル座標、元サイズ・倍率、行／列の非表示状態を `%LOCALAPPDATA%\EvidenceCrafter\logs\diagnostic.jsonl` に記録する。この失敗診断は通常の診断ログ設定が無効でも記録し、セル内容・画像データ・Workbookのパスは含めない。画像の取り消しを確認できない場合はアプリの変更操作を停止し、復旧用PNGのパスをステータスへ表示する。

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

- Release buildでwarning 0
- Excel非依存テストが全件成功
- `bootstrap.ps1 -Publish` が単一EXEを生成
- 参照ベースラインのハッシュ不変
