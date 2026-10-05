# 挿入画像サイズ検証エラーの修正計画

作成日: 2026-10-05。調査対象: `282a2eb`、ソース上のアプリ版 `0.1.0-preview.35`。
状態: 計画に基づく実装を実施。実測と採用した補正、検証・配布の結果は[検証記録](review-size-verification-2026-10-05.md)を参照。以下は当初の計画と報告値を保持する。手元のExcelでは自然発生条件を特定できず、報告値を注入した実Excelで補正と拒否・復旧を検証する方針を採用した。

## 報告と確認できたこと

報告: 「自動配置に失敗したため変更を取り消しました: 挿入画像サイズを確認できませんでした…」

現行コードでは `ExcelAutomaticPlacementService` → `ExcelImagePlacementService.PlaceImage` → `PlaceWorkbook` → `VerifyGeometry` の順で画像を配置・検証する。画像挿入と属性設定の後、Width／Heightについて、予定値・実値が有限かつ正であることと、差が0.05pt以内であることを要求している。失敗すると作成画像を削除し、自動配置側で既に配置した画像・挿入行を戻す。

この文言は必ずしも「サイズを読み取れない」を意味しない。読み取れたサイズが予定値と一致しない場合や0の場合にも表示する。COM読取自体の例外は別の工程付きメッセージになる。位置から順に検証し、最初の不一致だけをメッセージへ表示する。

前回のTop修正では位置の丸め判定だけを変更し、サイズは0.05ptのまま維持している。位置の修正が今回の原因と確定したわけではない。画像サイズ計算は `ImageSizingService` にあり、画面画像は `ScreenImageSizing` でpixelを96 DPI基準のptへ換算し、配置幅と倍率から予定サイズを計算する。

ユーザーから追加で「予定745.5pt、実際745.2000122070312pt、最新バージョン使用」と報告された。差は **−0.29998779296875pt（約−0.30pt）** で、現行の0.05pt判定を約6倍超える。745.5は単精度でも正確に表現でき、この値付近の単精度の刻みは0.00006103515625ptである。今回の差は4915刻み相当なので、入力値の単純な単精度化や隣接値への丸めだけでは説明できない。前回のTop修正の流用や、0.30ptを無条件に許容する変更を先に決める根拠にはならない。

Width／Heightのどちらか、配置座標・倍率、正確なアプリ版番号・Excel版は未確認。手元の診断ログは最終更新が2026-09-30で、今回のplacement失敗記録はなかった。現時点でサイズ変化が生じる工程や原因は未確定。

## 1. 失敗条件を確定する

- エラー全文、アプリ版、`%LOCALAPPDATA%\EvidenceCrafter\logs\diagnostic.jsonl` の該当placement記録を照合する。既存ログには予定／実geometry、元サイズ、倍率、セル、Side、工程、Excel版、配置セルの行・列非表示状態があるため、新しいログ基盤は追加しない。
- **予定745.5ptを最優先の再現条件にする。** Width／Heightが未特定なのでそれぞれを745.5ptにする条件を用意する。画面画像の換算では994pixelが745.5ptに相当するため、該当辺994pixel・倍率1の条件と、大きな元画像を縮小して同じ745.5ptにする条件を比較する。もう一方の辺は複数の縦横比を使い、報告元の画像寸法と同一だとは扱わない。
- 報告元の倍率・座標が得られた場合は一致させる。未入手でも一時ブックで調査を進め、報告元条件での解消確認は未完了として扱う。
- 既存の専用Excelプロセス監督付きテストを使い、一時ブックに画像を配置する。故障注入による不一致と、通常の挿入で自然に起こる不一致を区別する。
- 既存の `PlacementStageObserved` の `AfterInsert`／`AfterAttributes` を使って実Left・Top・Width・Heightを記録する。属性設定後に変化する場合のみ、調査用コードでName、AlternativeText、LockAspectRatio、Placementのどの設定で変わるかを絞る。
- 同じ画像・同じ予定サイズを低い座標と大きい座標へ配置して比較する。単純な整数サイズだけでなく、縮小後の小数サイズ、縦長・横長、端数pixelを使う。画像自体の大きさと配置座標の影響を分離する。
- 非表示行・列、結合セル、行挿入後、NEW先行→OLD後追いも切り分ける。非表示fixtureは実際のHidden・Height・Widthを確認してから試験する。

## 2. 再現結果に応じて共通処理を最小修正する

主な変更先は `ExcelImagePlacementService.cs`。最初から許容誤差の拡大や再設定を入れず、以下の分岐で必要な変更だけを採用する。

| 再現結果 | 修正方針 |
| --- | --- |
| 予定サイズの計算・受け渡しが誤っている | `ImageSizingService` または呼び出し元の単位・倍率の不整合を直す。96 DPI基準の既存仕様は維持する。 |
| 挿入時または属性設定時にサイズが変わり、再設定で正しく戻せる | 新規Shapeだけを対象に、縦横比ロックを考慮したサイズ設定順を修正する。補正は必要な場合の1回に限定し、属性復元後に位置・幅・高さを再検証する。 |
| 再設定でも変わらない微小差が、通常挿入で再現する | 実測した数値変換に基づき、サイズ専用の比較規則と上限を設ける。位置用 `PositionMatches` をそのまま流用しない。未確認の差は拒否する。 |
| 実サイズが0、明確な変形、配置領域不足など | 配置失敗を維持する。事前に判定可能な条件だけ変更前に拒否し、理由を具体化する。 |

サイズの丸めを許す変更が必要な場合は、Width／Heightの個別差だけでなく、縦横比、Sideの右端、CASE下端、他の画像との間隔へ及ぶ影響も確認する。計画サイズと実サイズの差が既存の配置安全条件を越える場合は成功扱いにしない。

無条件の1pt許容、検証削除、値の64への置換、待機・再挿入ループは追加しない。根拠のない丸め式や相対誤差率を先に決めない。

## 3. 共通経路・履歴・取り消しを維持する

- 共通処理の修正を自動配置、手動配置、それぞれのRedoへ反映する。呼び出し元ごとに同じ補正を追加しない。
- 成功時の `ManagedShapeTarget` には補正・検証後のExcel実値を格納する。メタデータに記録する元サイズ・倍率との整合も確認する。
- NEW／OLDの実Top差0.05pt、参照画像の外部編集検知、既存画像のサイズ検証を緩めない。予定サイズの丸め対応と、既存画像が変更されていないことの判定を分ける。
- 失敗時の画像削除確認、既存の行補償、取り消し未確認時の操作停止・復旧PNG保持を維持する。サイズ補正の途中で例外が発生した場合も同じ失敗経路へ戻す。
- エラーは、取得済みサイズの不一致と読み取り失敗を区別できる表現にする。予定・実際・差は残す。追加診断は調査で不足が判明した項目だけとし、Workbookパス・セル内容・画像内容は記録しない。

## 4. 回帰テストと受入確認

既存の `ImagePlacementVerificationTests.cs` を中心に、再現した失敗値を修正前に失敗する回帰条件として追加する。サイズ計算を変更する場合だけ `ImageSizingServiceTests.cs`／`ScreenImageSizingTests.cs` も拡張する。

- 単体: 予定745.5／実際745.2000122070312が現行判定で拒否されることをWidth／Heightの両方で固定する。最終的に再設定で直す場合はこの拒否を維持し、比較規則を直す場合は実Excelで裏付けた条件だけ受入側へ変更する。正常サイズ、許容境界の内外、0・負値・NaN・Infinity、明確な変形も確認する。丸め判定を変える場合は、許す実測例と、その直外の拒否例を対にする。
- 実Excel: 整数／小数サイズ、縮小、ペア配置の倍率、縦長／横長、通常／下方座標、非表示行・列、結合セル、行挿入後を確認する。製品の比較関数だけで合否を判断せず、COMで独立に実サイズと配置境界を読む。
- 操作: 自動／手動配置、NEW→OLD後追い、末尾整理、Undo／Redo、同じSideの次CASE移動を確認する。
- 失敗: 複数画像の途中失敗、補正途中の例外、画像削除失敗、行補償失敗で、既存画像・値・数式・書式と復旧情報を確認する。
- 既存の `InsertedImages_VerifyGeometryAndPreserveFailedRollback`、`SameSideBackfill`、`AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard` を再実行する。行操作を変更した場合は `RowMutationOptimizations` も対象にする。

実装後の基本チェック:

```powershell
dotnet build EvidenceCrafter.sln -c Release
dotnet test tests/EvidenceCrafter.Tests/EvidenceCrafter.Tests.csproj -c Release --no-build --filter 'TestCategory!=ExcelIntegration'
dotnet test tests/EvidenceCrafter.Tests/EvidenceCrafter.Tests.csproj -c Release --no-build --filter 'FullyQualifiedName~InsertedImages_VerifyGeometryAndPreserveFailedRollback|FullyQualifiedName~SameSideBackfill|FullyQualifiedName~AutomaticPlacement_AdvancesAfterSafeTailFailureAndPreservesRecoveryGuard'
```

実Excel試験は同時実行しない。通常成功経路へ全Shape走査を増やさず、追加のCOM往復を導入した場合だけ既存性能harnessで影響を確認する。

## 完了条件

1. 自然に再現した正常配置のサイズ検証エラーが解消し、独立読取でも実サイズ・配置範囲が妥当である。
2. 不正なサイズと配置不能条件は引き続き拒否する。
3. 横並び、Undo／Redo、次CASE移動、外部編集検知と失敗時の復旧が維持される。
4. Releaseビルド警告0、非Excelテストと影響する実Excel試験が合格する。
5. 再現値、採用した補正または比較規則、その上限、検証結果を記録する。報告元との照合ができていない場合は明記する。

検証完了後に配布版の更新・単一EXEとSHA-256の生成へ進む。2026-10-05の追加依頼で実装・検証・remoteへのpushが承認された。
