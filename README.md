# EvidenceCrafter

Case&Evidence形式のExcelブックへスクリーンショットを安全に配置するWindowsデスクトップツールです。`0.1.0-preview.30` は完成仕様レビュー候補です。

## 実装済み

- 起動中Excelブックの探索と、接続世代・ウィンドウ・ReadOnly・保護状態の操作時再検証（終了イベント監視が利用できない環境にも対応）
- 通知を出さない独自範囲キャプチャ、Clipboard監視、プレビュー、色選択対応の画像編集（枠・矢印・枠付き文字・モザイク・切り抜き、編集Undo/Redo）
- 数値Case番号と左から順に並ぶ任意文字列のSideヘッダーによる罫線任意の解析、Newのみ／New・Old対応、前後Case移動、Case範囲外では最初の未配置Caseへ補完する単一画像自動配置
- 同じCaseのNew/Oldがそろった場合に、安全な空行だけを削除して末尾2行を維持
- ActiveCell上への行挿入と、Case終端罫線を保持する安全条件付き末尾行削除
- 管理画像の差し替え・削除
- 配置・行操作・差し替え・削除のアプリ内Undo/Redo（削除行のネイティブ書式、同一Sheet参照数式、名前定義、印刷範囲も復元。別Sheetに数式がある場合は行削除を安全停止）
- 左右余白、診断ログ、グローバルショートカットの設定
- 変更途中の補償、20件の履歴上限、Excelの自動保存を行わない安全境界
- 起動中の二重起動を抑止し、2回目の起動で既存ウィンドウ（編集中のモーダル画面を含む）を前面化
- New/Oldの対応する画像を同じCASE内の序数で対応付け、配置幅を超えない共通倍率で新画像を配置し、既存画像も同倍率へ拡縮（旧メタデータなし画像は表示幅を基準に安全に近似）

## ビルド・検証

.NET SDK **9.0.304** と Windows が必要です。`global.json` は互換性のある新しい .NET 9 SDK も許可し、CI・発行検証は 9.0.304 を使用します。SDK のバージョンとアプリ実行用 .NET 9 ランタイムのバージョンは異なります。

```powershell
.\bootstrap.ps1
dotnet test .\tests\EvidenceCrafter.Tests\EvidenceCrafter.Tests.csproj -c Release --no-build --filter "TestCategory=ExcelIntegration"
```

## 発行

```powershell
.\bootstrap.ps1 -Publish -Runtime win-x64
```

成果物は `artifacts/publish/win-x64/EvidenceCrafter.exe` です。SHA-256 sidecarも同じフォルダに生成されます。

## 使い方

1. Excelで対象ブックとシートを開き、書き込み先Case内のセルを選択します。
2. EvidenceCrafterでWorkbookを選ぶと、実際のSheet名とCase番号が表示されます。必要な場合はSheet名またはCase番号を編集し、New/Oldを選びます。
3. `Ctrl + Shift + E`または「画面をキャプチャ」で範囲を選び、表示された配置予定（Sheet、Case、Side、開始セル、追加行数、画像幅）を確認して、編集の有無だけを選びます。配置先は自動判定されます。従来の`Win + Shift + S`も利用できます。
4. 管理画像をExcelで1つ選択すると差し替え・削除できます。`Ctrl+Z` / `Ctrl+Y` または画面ボタンで履歴を操作します。
5. 内容を確認後、保存はExcel側で明示的に行います。

参照元 `C:\work\Macro\Case&Evidence` は変更しません。実Excel結合テストはOS一時フォルダに専用ブックを作成します。詳細は [実装状態](docs/review-status.md) と [テスト手順](docs/testing.md) を参照してください。
