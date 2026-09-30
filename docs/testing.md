# EvidenceCrafter テスト方針

2026-09-11のキャプチャ・テキスト編集・最前面機能の検証結果と手動受入手順は
[実装レビュー](review-capture-editor-window.md) を参照。

2026-10-01のNEW先行・OLD後追い配置と性能測定は
[改修検証記録](review-same-side-backfill-2026-10-01.md) を参照。

## 通常テスト

`EvidenceCrafter.Tests` はExcelなしで動くCoreテストを既定とする。

```powershell
dotnet test tests\EvidenceCrafter.Tests\EvidenceCrafter.Tests.csproj --filter 'TestCategory!=ExcelIntegration'
```

## Excel結合テスト

Officeを起動するテストには `[TestCategory("ExcelIntegration")]` を付ける。通常CIでは実行せず、Office導入済みの明示環境だけで実行する。

現行の実機テストは、同一Excelプロセスに一時Workbookを2冊作成し、Workbook別HWND、通常編集後の接続ID維持、非アクティブ側のセル選択、`Application.Goto`、`EnableEvents`復元を確認する。さらに一時PNGの手動／Case自動配置、必要行挿入、管理画像のExport・差し替え・削除・元geometry復元、保護Sheetでの拒否を確認する。行操作ではActiveCell上への挿入、live safety snapshotに基づくCase末尾削除、Undo/Redo、編集済み挿入行のUndo拒否を確認する。Close取消、監視token再発行、確定Close後の旧identity拒否、生成Excel PID終了と一時ファイル回収も同一シナリオで検証する。Windows PowerShell 5.1でも参照ハッシュ検証を再現できるよう、スクリプトはUTF-8 BOMで保存する。

`SameSideBackfill` は4CASEの配置・末尾整理・同じSideの次CASE移動後の実Top、外部編集とRedo、エビデンス列外の削除保護を確認する。`SameSidePerformance` は対象CASE外の0/50/200/500図形で配置を各10回、末尾整理を各1回測定する。全体中央値と混同しない。

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
