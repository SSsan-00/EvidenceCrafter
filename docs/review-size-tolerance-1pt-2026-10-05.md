# 新規画像のサイズ差1pt許容

2026-10-05のユーザー依頼「1ptまで許容する実装にしてremote push」に対応。アプリ版preview.37。改修ソース: `f1ef65fd19b240115e985774ed6d5290b5459d91`。実装・検証・発行完了。

## 変更

`ExcelImagePlacementService` の新規画像の予定／実Width・Height比較を、それぞれ差1pt以内（境界を含む）で成功とする。有限かつ正であることを引き続き要求する。自動配置・手動配置・各Redoが同じ共有処理を使う。

初回の0.05pt比較と、差0.5pt以内の一度だけのサイズ明示設定は維持した。補正後にも約0.30ptの差が残る場合や、補正対象外の0.5pt超～1pt以内の差の場合も受け入れる。補正中の例外、管理名の不一致、位置ずれ、1ptを超えるサイズ差は既存の取り消し経路へ戻す。

受け入れた画像の実Widthが配置可能幅を0.05pt超えて上回る場合は拒否する。1pt許容で予定幅からはみ出す場合も、配置可能幅の超過を成功扱いにしない。行領域の予約・CASE境界・末尾整理は既存の計画と安全確認を使用する。

縦横比ロックの設定・補正時の復元を維持し、履歴にはExcelの最終実geometryを保存する。NEW／OLDの実Top差0.05pt、参照画像のサイズ検証0.05pt、既存画像の外部編集検知を変更していない。許容差は新規画像の予定サイズとの比較に限定した。

失敗診断に `insertedImageSizeTolerancePoints: 1` を追加する。通常成功経路のCOM往復は増やしていないため、性能の再計測は実施していない。

## 検証

- Releaseビルド: 警告0・エラー0。
- 非Excelテスト: 171／171合格、スキップ0（`artifacts/size-tolerance-1pt/unit.trx`）。幅／高さの±1pt受入、境界外の±1.001pt拒否、0・負値・非有限値、位置ずれの拒否、補正上限と受入上限の違い、診断項目を確認した。
- 影響する実Excel試験: 4／4合格、スキップ0（`artifacts/size-tolerance-1pt/affected.trx`、2分21秒）。`InsertedImages_CorrectReportedSizeDriftOnceAndVerifyFinalGeometry`、`InsertedImages_VerifyGeometryAndPreserveFailedRollback`、`SameSideBackfill_AlignsActualTopsAcrossCases`、`AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard` が合格した。

実Excel試験では、幅／高さの±1pt受入と±1.25pt拒否、補正しても報告値745.2000122070312ptが残る場合の成功を確認した。受入後の実geometryをCOMで独立に読み、履歴用geometryと一致することを確認した。配置可能幅を超える+1pt幅は拒否した。

受入後の画像をさらに0.30pt幅変更した場合、元の履歴対象での削除（Undo）を拒否することを確認した。実サイズへ戻して削除できることも確認し、挿入時の許容差が既存画像の外部編集検知へ波及していないことを検証した。

自動配置・後追い・Redoの試験では補正後にも約0.30ptの高さ差を残し、配置成功・実Top一致・実サイズを使った履歴を確認した。複数画像の途中失敗は1ptを超える差で発生させ、画像・追加行の補償とCASEアンカー・既存内容の保全を確認した。既存の画面操作試験でUndo／Redo・次CASE移動・復旧ガードも合格した。

報告元ブックを操作した検証ではなく、専用Excelプロセスと一時ブックでの試験である。

## 発行

`bootstrap.ps1 -Publish -Runtime win-x64` が成功した。最終Releaseビルドも警告0・エラー0、非Excelテスト171／171合格、スキップ0。単一EXEとSHA-256 sidecarのみの出力を確認し、ハッシュが一致した。

- EXE: `artifacts/publish/win-x64/EvidenceCrafter.exe`
- ProductVersion: `0.1.0-preview.37+f1ef65fd19b240115e985774ed6d5290b5459d91`
- SHA-256: `5A02C63D46134B677C63B4B229B1F05E966E6549BF563C9BFC834E8D916C43F8`
- 発行ログ: `artifacts/size-tolerance-1pt/publish.log`

ソースと検証記録をremote mainへ反映する。EXEと生記録は既存方針どおりignored artifact。

旧preview.36は `artifacts/size-tolerance-1pt/EvidenceCrafter-preview36.exe` に退避済み。SHA-256は `A3C707ECE430DA6C993183E30F8AF17FC3655F17AF1C824B46866593A5CF969C`。
