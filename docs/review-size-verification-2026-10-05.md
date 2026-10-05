# 挿入画像サイズ差への限定補正と検証

対象: [修正計画](implementation-plan-size-verification-2026-10-05.md)。2026-10-05実装・検証・発行完了。アプリ版preview.36。改修ソース: `74ac3514056b898f4f6b285ee4005c9fd40b0525`。

後続のユーザー依頼により、preview.37では新規画像の最終サイズ差を1ptまで許容する。[後続の検証記録](review-size-tolerance-1pt-2026-10-05.md)。以下はpreview.36時点の記録を保持する。

## 報告と調査結果

ユーザー報告は予定745.5pt、実際745.2000122070312pt。差は−0.29998779296875ptで、従来のサイズ許容差0.05ptでは拒否する。幅／高さのどちらか、報告元のExcel版・座標・倍率は未確認。745.5自体はSingleで正確に表せるため、前回の位置丸め判定をサイズへ流用しない。

専用プロセス監督付きharnessを利用し、報告元ブックを操作せず一時ブックとPNGで調査した。手元のExcelはVersion 16.0、Build 19127。

- 994pixelの辺、原寸／2倍の画像を半分へ縮小、縦長／横長／正方形、通常・下方行・遠方列の35条件で、745.5ptは予定どおり取得できた。挿入直後と属性設定後にサイズ差は生じず、旧サービスでも35条件が成功した。
- xlsへSaveAs後の同じ35条件でもサイズ差は生じなかった。この調査の終了時は追加xlsファイルをharnessが削除対象に含めず、一時ディレクトリの削除が失敗した。Excel終了と残存ファイルを確認し、生成したxlsと空ディレクトリを回収した。xlsを再度開いた試験ではない。
- 画像DPI7条件（96／120／144／72／300、縦横で異なるDPI）、Excel表示倍率6条件（50／75／100／125／150／200%）、予定サイズ5条件の210試行でも、自然なサイズ差は観測しなかった。こちらはPerMonitorV2で実行したが、Windows自体の拡大率は変更していない。

**自然発生の原因は未特定。** 本修正は原因が確定した丸め許容ではなく、報告された差を実Excel上の新規Shapeへ注入し、予定サイズの明示設定で正確に戻せることを確認した限定的な補正である。手元の実機で再現しなかったことを、報告元で発生しない根拠にはしない。

原記録・再実行用probeはignored artifactの `artifacts/size-verification-investigation`。`probe.log`、`probe-xls.log`、`probe-sweep.log` とJSONに保存した。

## 実装

`ExcelImagePlacementService.PlaceWorkbook` の既存検証で不一致があった場合、管理名が一致し、位置が既存条件を満たし、各辺が有限・正で予定との差が0.5pt以内の新規Shapeに限り、以下を一度だけ実行する。

1. 縦横比ロックを解除し、予定Width／Heightを明示設定する。
2. 縦横比ロックを戻し、管理名・Left・Top・Width・Heightを再取得する。
3. 従来のgeometry検証を再度適用する。サイズは引き続き0.05pt以内の一致を要求する。

0.5ptは報告差約0.30ptを補正対象とする保守的な操作上限で、Excelの丸め単位や許容誤差だと実測した値ではない。上限を超える不一致、0・負値・非有限値、位置ずれ、管理名の不一致は補正しない。補正後にも約0.30ptの差が残る場合は取り消す。再挿入・待機・複数回の補正は行わない。

自動配置・手動配置・各Redoが同じ共有処理を使う。成功結果には最終実geometryを保存する。参照画像の外部編集判定、NEW／OLDの実Top差0.05pt、CASE整理の条件を維持している。

通常成功経路は既存のCOM読取のまま。補正対象の経路だけに設定・再読取を追加したため、通常配置の性能再計測は実施していない。高速化を主張する変更ではない。

失敗診断に `beforeSizeCorrection` と `sizeCorrectionAttempted` を追加した。補正操作中に例外が起きた場合は最終実geometryを未確認（null）として記録し、補正前の値を補正後の実値として残さない。サイズ不一致の文言は「サイズが予定値と一致しませんでした」に変更し、予定・実際・差を維持する。既存の削除確認、行補償、取り消し未確認時の操作停止・復旧PNG保持を使用する。

## 検証

- Releaseビルド: 警告0・エラー0。
- 非Excelテスト: 171／171合格、スキップ0（`artifacts/size-verification-fix/unit.trx`）。報告値をそのまま成功としないこと、補正上限の内外、不正値・位置ずれの拒否、診断情報を確認した。
- 新規実Excel試験の初回は単独で合格（`size-correction.trx`）。その後、自動配置・後追い・Redo・複数画像の補償を追加して、影響試験と一緒に再検証した。
- 最終の影響実Excel試験は4／4合格、スキップ0（`affected.trx`、合計2分18秒）。

| シナリオ | 確認 | 所要時間 |
| --- | --- | ---: |
| InsertedImages_CorrectReportedSizeDriftOnceAndVerifyFinalGeometry | 報告値の補正、最終geometry、失敗復旧、自動配置・後追い・Redo・行補償 | 15.8秒 |
| InsertedImages_VerifyGeometryAndPreserveFailedRollback | 座標丸め、サイズ0・大きな不一致の拒否、取り消し未確認時の保全 | 13.3秒 |
| SameSideBackfill_AlignsActualTopsAcrossCases | 4CASEの横並び、行整理、参照画像の外部編集拒否 | 80.5秒 |
| AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard | 画面からの配置・Undo／Redo、次CASE移動、復旧ガード | 28.9秒 |

新規試験ではWidth／Height、原寸／縮小、低い／大きい座標、挿入直後／属性設定後の16組で報告値を注入し、COMで直接読んだサイズと履歴用実値が一致することを確認した。各組の補正回数は1回。補正しても差が残る場合、補正中の例外、補正後の管理名／位置の変更、削除失敗を拒否し、通常の745.5pt配置では補正が0回であることも確認した。

自動配置では新規側・後追い側・Redoで約0.30ptのサイズ差を注入し、補正と実Top一致を確認した。複数画像の2枚目で差が残る場合、1枚目と追加行が補償され、CASEアンカーと既存図形・セル内容が保持されることを確認した。検証後のExcelプロセス残存0。

報告元ブックそのものの解消確認は未実施。報告元で再設定後も745.2000122070312ptとなる場合は失敗を維持するため、追加診断で補正前後の値と工程を確認できる。

## 発行

`bootstrap.ps1 -Publish -Runtime win-x64` が成功した。最終のReleaseビルドも警告0・エラー0、非Excelテスト171／171合格、スキップ0。出力ファイルが単一EXEとSHA-256 sidecarの2点だけであることを確認した。

- EXE: `artifacts/publish/win-x64/EvidenceCrafter.exe`
- ProductVersion: `0.1.0-preview.36+74ac3514056b898f4f6b285ee4005c9fd40b0525`
- SHA-256: `A3C707ECE430DA6C993183E30F8AF17FC3655F17AF1C824B46866593A5CF969C`（sidecar一致）
- 発行ログ: `artifacts/size-verification-fix/publish.log`

ソース・計画・検証記録をremote mainへ反映する。EXEと生記録は既存方針どおりignored artifact。

旧preview.35は `artifacts/size-verification-fix/EvidenceCrafter-preview35.exe` に保持した。SHA-256は `49EB861D4AEAE8843D1FD6F5FA63DBD707A1B5F1AB4F0C79631FE32B32D53647`。
