# CASE末尾整理と配置後移動の回帰修正

## 原因と修正

画像配置後、New／Oldがそろうと実行されるCASE末尾整理で、印刷範囲などの名前定義を `Names.Item` のプロパティ呼び出しで取得していた。Excelの `Names.Item` はメソッドであり、名前定義が存在すると `0x80020003` が発生する。新規ブックでは名前定義の件数が0のため、従来のfixtureでは問題を検出できなかった。

Microsoftの仕様: [Names.Item method](https://learn.microsoft.com/en-us/office/vba/api/excel.names.item)。名前定義の取得とUndo時の復元の両方をメソッド呼び出しへ修正した。印刷範囲を設定した実Excel一時ブックで、修正前の同一エラーと修正後の成功を確認した。

もう一つの問題は、末尾整理が行未変更のまま失敗しても、配置後の次CASE／Side移動の前にreturnしていたこと。配置済みの履歴を残したうえで設定どおり移動し、整理の中止理由をステータスに残す。行変更結果が不明、または退避の復旧が必要な場合は従来どおり移動・追加操作を停止し、PNGや行退避を保持する。

変更は共有の名前定義処理2箇所と自動配置後の分岐に限定した。差し替え・削除・行Undo／Redoも共有処理の修正を利用する。行削除の安全条件、厳密なfingerprint、移動設定、キャプチャ時の移動抑止は維持する。新しい依存関係・COM全件走査・再試行は追加していない。

## 検証記録

- 修正前: 印刷範囲を追加した `SameSideBackfill` が実Excelで `0x80020003` により失敗。記録: `artifacts/case-tail-regression/case-tail-names-before.trx`。
- 修正後: 同じfixtureのNEW先行・OLD後追い4CASE配置、末尾整理、横並び、同じSideの次CASE移動が合格。記録: `case-tail-names-after.trx`。
- 印刷範囲・名前定義付きの行削除、Undo／Redo、故障注入、結果不明時の退避保持が合格。名前定義はExcelから直接読み取り、退避前／削除後の記録と一致を確認。
- 自動配置ハンドラーを実Excel＋非表示MainFormで実行し、正常な末尾整理後と、別シートの数式を理由に行未変更で整理を中止した後の両方で、OLDのまま次CASEへ進むこと、履歴が残ることを確認。結果不明の行変更をUIに報告した後の配置停止も確認。
- UIfixtureの初回実行は動作の検証を通過したが、追加した設定・ログディレクトリの後片付け漏れで失敗した。続く実行でも二重Disposeによる設定ファイル再生成を検出した。fixtureを修正した最終実行は後片付けを含めて合格。失敗記録は `case-tail-ui-and-undo.trx`、`case-tail-ui-final.trx` に保持し、合格記録は `case-tail-ui-pass.trx` に保持。

- Releaseビルド: 警告0・エラー0。
- Excel非依存テスト: 170／170合格、スキップ0。
- 影響する実Excelシナリオ3件はすべて合格（SameSideBackfill、RowMutationOptimizations、AutomaticPlacementの画面ハンドラー）。最終画面ハンドラーでは `1-1 → 1-2 / Old` と、整理中止を伴う `1-2 → 1-3 / Old` を確認。
- `bootstrap.ps1 -Publish -Runtime win-x64`: 成功。生成したExcelプロセスの残存なし。今回、性能計測の再実行は行っていない。

ユーザーの元ブックは未提供のため、元ブックそのものの検証は未実施。印刷範囲付き一時ブックで報告と一致するエラーを再現して修正した。

## 配布用EXE

- `artifacts/publish/win-x64/EvidenceCrafter.exe`
- ProductVersion: `0.1.0-preview.34+8d9d76bba8fc6eff6e7dab65a46d1c75b2c16006`
- SHA-256: `40DA9AF124B764CBADEFC2931A11C131385C2C0437933425A8564BD5EDBFB3CC`（sidecar一致）
- 旧preview.33は `artifacts/case-tail-regression/EvidenceCrafter-preview33.exe` に退避済み。
- EXE・生テスト記録はローカルのignored artifact。ソースと検証文書をremote mainへ反映する。
